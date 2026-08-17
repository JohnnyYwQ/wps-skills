#!/usr/bin/env python3
"""Run one manifest-backed WPS Action through a temporary loopback service.

Input: ``invoke`` followed by one JSON request object.
Output: one structured JSON result on stdout; diagnostics are reserved for stderr.
Position: public process boundary used by the WPS Skill Package.
"""

import errno
import json
import os
import re
import secrets
import socket
import sys
import time
from contextlib import contextmanager
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

from wps_skill import addon_installer


HOST = "127.0.0.1"
PORT = 58891
MANIFEST_PATH = Path(__file__).resolve().parents[1] / "references" / "action-manifest.json"
TRANSPORT_ERROR_DETAILS = {
    "PORT_IN_USE": (
        "Loopback port {0} is already in use".format(PORT),
        True,
    ),
    "PORT_UNAVAILABLE": ("The loopback port could not be opened", True),
    "ADDIN_NOT_READY": (
        "WPS Add-in did not return a result before the timeout",
        True,
    ),
    "ACTION_TIMEOUT": (
        "WPS Action did not return a result before the timeout",
        True,
    ),
    "ADDIN_DISCONNECTED": (
        "WPS Add-in disconnected while sending its result",
        False,
    ),
    "INVALID_ADDIN_JSON": ("WPS Add-in result must be valid JSON", False),
    "INVALID_ADDIN_RESPONSE": (
        "WPS Add-in result payload must be a JSON object",
        False,
    ),
    "REQUEST_ID_MISMATCH": (
        "WPS Add-in result does not match the pending WPS Action",
        False,
    ),
}


class RunnerError(Exception):
    def __init__(self, code, message, retryable, action=None, operation=None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.retryable = retryable
        self.action = action
        self.operation = operation

    def as_result(self):
        result = {
            "ok": False,
            "error": {
                "code": self.code,
                "message": self.message,
                "retryable": self.retryable,
            },
        }
        if self.operation is not None:
            result["operation"] = self.operation
        else:
            result["action"] = self.action
        return result


class LoopbackBindError(Exception):
    def __init__(self, error_number):
        super().__init__(error_number)
        self.error_number = error_number


def transport_error_details(code):
    message, retryable = TRANSPORT_ERROR_DETAILS[code]
    return {"code": code, "message": message, "retryable": retryable}


def transport_runner_error(code, action):
    details = transport_error_details(code)
    return RunnerError(
        details["code"],
        details["message"],
        details["retryable"],
        action,
    )


def bind_error_code(bind_error):
    if bind_error.error_number == errno.EADDRINUSE:
        return "PORT_IN_USE"
    return "PORT_UNAVAILABLE"


def incomplete_exchange_error_code(exchange_state):
    if exchange_state["action_delivered"]:
        return "ACTION_TIMEOUT"
    return "ADDIN_NOT_READY"


@contextmanager
def action_lock(action=None, operation=None):
    """Hold the current user's non-blocking cross-process WPS Action lock."""
    lock_file = None
    try:
        lock_path = addon_installer.action_lock_path()
        lock_path.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
        if os.name != "nt":
            lock_path.parent.chmod(0o700)
        lock_file = lock_path.open("a+b")
        if os.name != "nt":
            lock_path.chmod(0o600)
    except (OSError, addon_installer.AddinInstallError) as error:
        if lock_file is not None:
            lock_file.close()
        if isinstance(error, addon_installer.AddinInstallError):
            raise RunnerError(
                error.code,
                error.message,
                error.retryable,
                action=action,
                operation=operation,
            )
        raise RunnerError(
            "LOCK_UNAVAILABLE",
            "The WPS Action lock could not be opened",
            True,
            action=action,
            operation=operation,
        )

    acquired = False
    try:
        try:
            if os.name == "nt":
                import msvcrt

                if lock_path.stat().st_size == 0:
                    lock_file.write(b"\0")
                    lock_file.flush()
                lock_file.seek(0)
                msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl

                fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            acquired = True
        except (OSError, IOError) as error:
            busy_error_numbers = {errno.EACCES, errno.EAGAIN}
            if error.errno not in busy_error_numbers:
                raise RunnerError(
                    "LOCK_UNAVAILABLE",
                    "The WPS Action lock could not be acquired",
                    True,
                    action=action,
                    operation=operation,
                )
            raise RunnerError(
                "ACTION_BUSY",
                "Another WPS Action is already running",
                True,
                action=action,
                operation=operation,
            )
        yield
    finally:
        if acquired:
            try:
                if os.name == "nt":
                    import msvcrt

                    lock_file.seek(0)
                    msvcrt.locking(lock_file.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl

                    fcntl.flock(lock_file.fileno(), fcntl.LOCK_UN)
            except OSError:
                pass
        lock_file.close()


def load_action(action_name):
    try:
        with MANIFEST_PATH.open(encoding="utf-8") as manifest_file:
            manifest = json.load(manifest_file)
        actions = manifest["actions"]
    except (OSError, KeyError, TypeError, ValueError) as error:
        raise RunnerError(
            "INVALID_MANIFEST",
            "The WPS Action manifest could not be loaded: {0}".format(error),
            False,
            action_name,
        )
    for action in actions:
        if action["action"] == action_name:
            return action
    raise RunnerError(
        "UNKNOWN_ACTION",
        "Unknown WPS Action: {0}".format(action_name),
        False,
        action_name,
    )


def value_matches_type(value, expected_type):
    type_checks = {
        "array": lambda item: isinstance(item, list),
        "boolean": lambda item: isinstance(item, bool),
        "null": lambda item: item is None,
        "number": lambda item: isinstance(item, (int, float))
        and not isinstance(item, bool),
        "object": lambda item: isinstance(item, dict),
        "string": lambda item: isinstance(item, str),
    }
    return expected_type in type_checks and type_checks[expected_type](value)


def validate_contract(value, schema, path):
    expected_types = schema.get("type")
    if isinstance(expected_types, str):
        expected_types = [expected_types]
    if expected_types and not any(
        value_matches_type(value, expected_type) for expected_type in expected_types
    ):
        if len(expected_types) == 1:
            expected = expected_types[0]
        else:
            expected = "one of {0}".format(", ".join(expected_types))
        return "{0} must be {1}".format(path, expected_type_phrase(expected))

    allowed_values = schema.get("enum")
    if isinstance(allowed_values, list) and value not in allowed_values:
        return "{0} must be one of {1}".format(
            path, ", ".join(str(item) for item in allowed_values)
        )

    pattern = schema.get("pattern")
    if isinstance(value, str) and isinstance(pattern, str):
        try:
            matches_pattern = re.search(pattern, value) is not None
        except re.error:
            matches_pattern = False
        if not matches_pattern:
            return "{0} must match pattern '{1}'".format(path, pattern)

    if isinstance(value, dict):
        required = schema.get("required", [])
        for field in required:
            if field not in value:
                return "{0} is missing required field '{1}'".format(path, field)

        properties = schema.get("properties", {})
        if schema.get("additionalProperties") is False:
            for field in value:
                if field not in properties:
                    return "{0} contains unknown field '{1}'".format(path, field)

        for field, field_value in value.items():
            if field in properties:
                problem = validate_contract(
                    field_value, properties[field], "{0}.{1}".format(path, field)
                )
                if problem:
                    return problem

    if isinstance(value, list) and "items" in schema:
        for index, item in enumerate(value):
            problem = validate_contract(
                item, schema["items"], "{0}[{1}]".format(path, index)
            )
            if problem:
                return problem

    alternatives = schema.get("anyOf", [])
    if alternatives and not any(
        validate_contract(value, alternative, path) is None
        for alternative in alternatives
    ):
        return "{0} must satisfy at least one allowed contract".format(path)
    return None


def expected_type_phrase(expected_type):
    if expected_type == "object":
        return "an object"
    if expected_type == "array":
        return "an array"
    if expected_type.startswith("one of "):
        return expected_type
    return "a {0}".format(expected_type)


def parse_request(raw_request):
    try:
        request = json.loads(raw_request)
    except (TypeError, ValueError):
        raise RunnerError("INVALID_REQUEST", "Request must be valid JSON", False)
    if not isinstance(request, dict):
        raise RunnerError("INVALID_REQUEST", "Request must be a JSON object", False)

    action_name = request.get("action")
    if not isinstance(action_name, str) or not action_name:
        raise RunnerError(
            "INVALID_REQUEST", "Request field 'action' must be a string", False
        )

    timeout_ms = request.get("timeout_ms", 30000)
    if (
        not isinstance(timeout_ms, int)
        or isinstance(timeout_ms, bool)
        or not 1 <= timeout_ms <= 300000
    ):
        raise RunnerError(
            "INVALID_REQUEST",
            "Request field 'timeout_ms' must be an integer between 1 and 300000",
            False,
            action_name,
        )

    if "confirmed" in request and not isinstance(request["confirmed"], bool):
        raise RunnerError(
            "INVALID_REQUEST",
            "Request field 'confirmed' must be a boolean",
            False,
            action_name,
        )
    return request


def make_handler(exchange_state, auth_token):
    class PollingHandler(BaseHTTPRequestHandler):
        def setup(self):
            BaseHTTPRequestHandler.setup(self)
            deadline = exchange_state.get("deadline")
            if deadline is not None:
                self.connection.settimeout(
                    max(0.001, deadline - time.monotonic())
                )

        def _is_authorized(self):
            provided = self.headers.get("Authorization", "")
            expected = "Bearer {0}".format(auth_token)
            return secrets.compare_digest(provided, expected)

        def _reject_unauthorized(self):
            self._send_json(
                401,
                {
                    "ok": False,
                    "error": {
                        "code": "AUTHENTICATION_FAILED",
                        "message": "Loopback request authentication failed",
                    },
                },
            )

        def _reject_protocol_error(self, status, code):
            error = transport_runner_error(
                code, exchange_state["action_request"]["action"]
            )
            exchange_state["protocol_error"] = error
            self._send_json(
                status,
                {
                    "ok": False,
                    "error": {"code": error.code, "message": error.message},
                },
            )

        def do_GET(self):
            if not self._is_authorized():
                self._reject_unauthorized()
                return
            if self.path != "/poll":
                self._send_json(404, {"error": "Not found"})
                return
            if self._send_json(
                200, {"actionRequest": exchange_state["action_request"]}
            ):
                exchange_state["action_delivered"] = True

        def do_POST(self):
            if not self._is_authorized():
                self._reject_unauthorized()
                return
            if self.path != "/result":
                self._send_json(404, {"error": "Not found"})
                return
            try:
                content_length = int(self.headers.get("Content-Length", "0"))
                request_body = self.rfile.read(content_length)
            except socket.timeout:
                exchange_state["protocol_error"] = transport_runner_error(
                    "ACTION_TIMEOUT",
                    exchange_state["action_request"]["action"],
                )
                return
            except (TypeError, ValueError):
                request_body = b""
                content_length = 0
            if len(request_body) != content_length:
                exchange_state["protocol_error"] = transport_runner_error(
                    "ADDIN_DISCONNECTED",
                    exchange_state["action_request"]["action"],
                )
                return
            try:
                payload = json.loads(request_body.decode("utf-8"))
            except (TypeError, ValueError, UnicodeDecodeError, json.JSONDecodeError):
                self._reject_protocol_error(400, "INVALID_ADDIN_JSON")
                return
            if not isinstance(payload, dict):
                self._reject_protocol_error(400, "INVALID_ADDIN_RESPONSE")
                return
            if payload.get("requestId") != exchange_state["action_request"][
                "requestId"
            ]:
                self._reject_protocol_error(409, "REQUEST_ID_MISMATCH")
                return
            exchange_state["result"] = payload.get("result")
            exchange_state["result_received"] = True
            self._send_json(200, {"ok": True})

        def do_OPTIONS(self):
            # Browser CORS preflights cannot carry Authorization. They expose
            # no WPS Action or result and mutate no exchange state.
            self._send_json(200, {})

        def _send_json(self, status, payload):
            body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(status)
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            self.send_header(
                "Access-Control-Allow-Headers", "Content-Type, Authorization"
            )
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            try:
                self.wfile.write(body)
                return True
            except (BrokenPipeError, ConnectionResetError):
                return False

        def log_message(self, format_string, *args):
            del format_string, args

    return PollingHandler


def exchange_with_addin(action_request, auth_token, timeout_ms):
    exchange_state = {
        "action_request": action_request,
        "result": None,
        "result_received": False,
        "protocol_error": None,
        "action_delivered": False,
    }
    try:
        server = HTTPServer((HOST, PORT), make_handler(exchange_state, auth_token))
    except OSError as error:
        raise LoopbackBindError(error.errno)
    try:
        deadline = time.monotonic() + (timeout_ms / 1000.0)
        exchange_state["deadline"] = deadline
        while (
            not exchange_state["result_received"]
            and exchange_state["protocol_error"] is None
            and time.monotonic() < deadline
        ):
            server.timeout = min(0.25, max(0.0, deadline - time.monotonic()))
            server.handle_request()
        return exchange_state
    finally:
        server.server_close()


def check_options(raw_options):
    if raw_options is None:
        return {"timeout_ms": 2000}
    try:
        options = json.loads(raw_options)
    except (TypeError, ValueError):
        raise RunnerError(
            "INVALID_REQUEST",
            "Check options must be valid JSON",
            False,
            operation="check",
        )
    if not isinstance(options, dict):
        raise RunnerError(
            "INVALID_REQUEST",
            "Check options must be a JSON object",
            False,
            operation="check",
        )
    timeout_ms = options.get("timeout_ms", 2000)
    if (
        not isinstance(timeout_ms, int)
        or isinstance(timeout_ms, bool)
        or not 1 <= timeout_ms <= 300000
    ):
        raise RunnerError(
            "INVALID_REQUEST",
            "Check option 'timeout_ms' must be an integer between 1 and 300000",
            False,
            operation="check",
        )
    if set(options) - {"timeout_ms"}:
        raise RunnerError(
            "INVALID_REQUEST",
            "Check options contain an unknown field",
            False,
            operation="check",
        )
    return {"timeout_ms": timeout_ms}


def ping_addin(auth_token, timeout_ms, expected_digest, restart_pending):
    failure_status = (
        "restart_required" if restart_pending else "addin_unavailable"
    )
    try:
        exchange_state = exchange_with_addin(
            {
                "action": "ping",
                "params": {},
                "requestId": "req-{0}".format(secrets.token_urlsafe(18)),
            },
            auth_token,
            timeout_ms,
        )
    except LoopbackBindError as error:
        return failure_status, transport_error_details(bind_error_code(error))
    result = exchange_state["result"]
    if exchange_state["protocol_error"] is not None:
        protocol_error = exchange_state["protocol_error"].as_result()["error"]
        return failure_status, protocol_error
    if not exchange_state["result_received"]:
        error = transport_error_details(
            incomplete_exchange_error_code(exchange_state)
        )
        return failure_status, error
    if (
        isinstance(result, dict)
        and result.get("success") is True
        and result.get("message") == "pong"
    ):
        loaded_digest = result.get("installDigest")
        if loaded_digest == expected_digest:
            return "ready", None
        if isinstance(loaded_digest, str) and loaded_digest:
            return "restart_required", None
    return failure_status, {
        "code": "INVALID_RESULT",
        "message": "WPS Add-in returned an invalid readiness result",
        "retryable": False,
    }


def readiness_context(action=None, operation=None):
    try:
        install_result = addon_installer.install()
    except addon_installer.AddinInstallError as error:
        raise RunnerError(
            error.code,
            error.message,
            error.retryable,
            action=action,
            operation=operation,
        )
    platform_name = install_result["platform"]
    return {
        "install": install_result,
        "wps_running": addon_installer.wps_is_running(platform_name),
        "auth_token": addon_installer.auth_token(platform_name),
        "source_digest": addon_installer.source_digest(),
        "restart_pending": addon_installer.restart_is_pending(platform_name),
    }


def invoke(request):
    action_name = request["action"]
    action = load_action(action_name)
    risk = action.get("risk")
    if risk not in {"read", "write", "destructive"}:
        raise RunnerError(
            "INVALID_ACTION_RISK",
            (
                "WPS Action has an invalid or missing risk classification: {0}"
            ).format(action_name),
            False,
            action_name,
        )
    if risk == "destructive" and request.get("confirmed") is not True:
        raise RunnerError(
            "CONFIRMATION_REQUIRED",
            "Destructive WPS Action requires confirmed=true: {0}".format(
                action_name
            ),
            False,
            action_name,
        )
    params = request.get("params", {})
    params_problem = validate_contract(params, action["parameters"], "params")
    if params_problem:
        raise RunnerError(
            "INVALID_PARAMS", params_problem, False, action_name
        )
    with action_lock(action=action_name):
        return invoke_locked(action, action_name, params, request)


def invoke_locked(action, action_name, params, request):
    readiness = readiness_context(action=action_name)
    install_result = readiness["install"]
    if install_result["restart_required"]:
        raise RunnerError(
            "WPS_RESTART_REQUIRED",
            "WPS Add-in was installed or updated; restart WPS Office before retrying",
            True,
            action_name,
        )
    if not readiness["wps_running"]:
        raise RunnerError(
            "WPS_NOT_RUNNING",
            "WPS Office is not running",
            True,
            action_name,
        )
    if readiness["restart_pending"]:
        raise RunnerError(
            "WPS_RESTART_REQUIRED",
            "WPS Add-in was installed or updated; restart WPS Office before retrying",
            True,
            action_name,
        )
    auth_token = readiness["auth_token"]
    timeout_ms = request.get("timeout_ms", 30000)
    try:
        exchange_state = exchange_with_addin(
            {
                "action": action_name,
                "params": params,
                "requestId": "req-{0}".format(secrets.token_urlsafe(18)),
            },
            auth_token,
            timeout_ms,
        )
    except LoopbackBindError as error:
        raise transport_runner_error(bind_error_code(error), action_name)

    if exchange_state["protocol_error"] is not None:
        raise exchange_state["protocol_error"]
    if not exchange_state["result_received"]:
        error_code = incomplete_exchange_error_code(exchange_state)
        raise transport_runner_error(error_code, action_name)

    result = exchange_state["result"]
    if not isinstance(result, dict):
        raise RunnerError(
            "INVALID_RESULT",
            "WPS Add-in returned an invalid result",
            False,
            action_name,
        )
    if not isinstance(result.get("success"), bool):
        raise RunnerError(
            "INVALID_RESULT",
            "WPS Add-in result must include a boolean 'success' field",
            False,
            action_name,
        )
    if not result["success"]:
        raise RunnerError(
            "WPS_ACTION_FAILED",
            str(result.get("error") or "WPS Action failed"),
            False,
            action_name,
        )
    if "data" in result:
        data = result["data"]
    else:
        data = {
            key: value
            for key, value in result.items()
            if key not in ("success", "installDigest")
        }
    result_problem = validate_contract(data, action["result"], "result")
    if result_problem:
        raise RunnerError(
            "INVALID_RESULT", result_problem, False, action_name
        )
    return {"ok": True, "action": action_name, "data": data}


def check_addin(options):
    readiness = readiness_context(operation="check")
    install_result = readiness["install"]
    if install_result["restart_required"]:
        return {
            "status": "restart_required",
            "ready": False,
            "restart_required": True,
            "platform": install_result["platform"],
            "architecture": install_result["architecture"],
        }
    if not readiness["wps_running"]:
        return {
            "status": "wps_not_running",
            "ready": False,
            "restart_required": False,
            "platform": install_result["platform"],
            "architecture": install_result["architecture"],
        }
    platform_name = install_result["platform"]
    expected_digest = readiness["source_digest"]
    status, transport_error = ping_addin(
        readiness["auth_token"],
        options["timeout_ms"],
        expected_digest,
        readiness["restart_pending"],
    )
    ready = status == "ready"
    if ready:
        addon_installer.acknowledge_loaded_digest(platform_name, expected_digest)
    result = {
        "status": status,
        "ready": ready,
        "restart_required": status == "restart_required",
        "platform": install_result["platform"],
        "architecture": install_result["architecture"],
    }
    if transport_error is not None:
        result["error"] = transport_error
    return result


def main(argv):
    if len(argv) == 2 and argv[1] == "install":
        try:
            install_result = addon_installer.install()
        except addon_installer.AddinInstallError as error:
            raise RunnerError(
                error.code,
                error.message,
                error.retryable,
                operation="install",
            )
        result = {
            "ok": True,
            "operation": "install",
            "data": install_result,
        }
        print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
        return 0
    if 2 <= len(argv) <= 3 and argv[1] == "check":
        options = check_options(argv[2] if len(argv) == 3 else None)
        with action_lock(operation="check"):
            check_result = check_addin(options)
        result = {"ok": True, "operation": "check", "data": check_result}
        print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
        return 0
    if len(argv) != 3 or argv[1] != "invoke":
        raise RunnerError(
            "INVALID_REQUEST", "Usage: wps.py invoke '<json-request>'", False
        )
    request = parse_request(argv[2])
    result = invoke(request)
    print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
    return 0


def action_from_argv(argv):
    if len(argv) < 3:
        return None
    try:
        request = json.loads(argv[2])
    except (TypeError, ValueError):
        return None
    if isinstance(request, dict) and isinstance(request.get("action"), str):
        return request["action"]
    return None


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv))
    except RunnerError as error:
        print(
            json.dumps(error.as_result(), ensure_ascii=False, separators=(",", ":"))
        )
        sys.exit(1)
    except Exception as error:
        unexpected = RunnerError(
            "INTERNAL_ERROR",
            "The WPS Runner failed unexpectedly",
            False,
            action_from_argv(sys.argv),
        )
        print(
            json.dumps(unexpected.as_result(), ensure_ascii=False, separators=(",", ":"))
        )
        print("Runner diagnostic: {0}".format(error), file=sys.stderr)
        sys.exit(1)
