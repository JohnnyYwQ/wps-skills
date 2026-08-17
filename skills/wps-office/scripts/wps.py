#!/usr/bin/env python3
"""Run one manifest-backed WPS Action through a temporary loopback service.

Input: ``invoke`` followed by one JSON request object.
Output: one structured JSON result on stdout; diagnostics are reserved for stderr.
Position: public process boundary used by the WPS Skill Package.
"""

import errno
import json
import secrets
import sys
import time
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

from wps_skill import addon_installer


HOST = "127.0.0.1"
PORT = 58891
MANIFEST_PATH = Path(__file__).resolve().parents[1] / "references" / "action-manifest.json"


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
    return request


def make_handler(state, auth_token=None):
    class PollingHandler(BaseHTTPRequestHandler):
        def _is_authorized(self):
            return auth_token is None or self.headers.get("Authorization") == (
                "Bearer {0}".format(auth_token)
            )

        def do_GET(self):
            if self.path != "/poll":
                self._send_json(404, {"error": "Not found"})
                return
            if not self._is_authorized():
                self._send_json(401, {"error": "Unauthorized"})
                return
            self._send_json(200, {"command": state["command"]})

        def do_POST(self):
            if self.path != "/result":
                self._send_json(404, {"error": "Not found"})
                return
            if not self._is_authorized():
                self._send_json(401, {"error": "Unauthorized"})
                return
            try:
                content_length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
            except (TypeError, ValueError, UnicodeDecodeError, json.JSONDecodeError):
                self._send_json(400, {"error": "Invalid JSON"})
                return
            if payload.get("requestId") != state["command"]["requestId"]:
                self._send_json(409, {"error": "Unknown requestId"})
                return
            state["result"] = payload.get("result")
            self._send_json(200, {"ok": True})

        def do_OPTIONS(self):
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
            self.wfile.write(body)

        def log_message(self, format_string, *args):
            del format_string, args

    return PollingHandler


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


def ping_addin(auth_token, timeout_ms):
    state = {
        "command": {
            "action": "ping",
            "params": {},
            "requestId": "req-{0}".format(secrets.token_urlsafe(18)),
        },
        "result": None,
    }
    try:
        server = HTTPServer((HOST, PORT), make_handler(state, auth_token))
    except OSError as error:
        if error.errno == errno.EADDRINUSE:
            return False
        raise
    try:
        deadline = time.monotonic() + (timeout_ms / 1000.0)
        while state["result"] is None and time.monotonic() < deadline:
            server.timeout = min(0.25, max(0.0, deadline - time.monotonic()))
            server.handle_request()
    finally:
        server.server_close()
    result = state["result"]
    return (
        isinstance(result, dict)
        and result.get("success") is True
        and result.get("message") == "pong"
    )


def invoke(request):
    action_name = request["action"]
    action = load_action(action_name)
    if action.get("risk") != "read":
        raise RunnerError(
            "ACTION_NOT_READ_ONLY",
            "WPS Action is not read-only: {0}".format(action_name),
            False,
            action_name,
        )
    params = request.get("params", {})
    params_problem = validate_contract(params, action["parameters"], "params")
    if params_problem:
        raise RunnerError(
            "INVALID_PARAMS", params_problem, False, action_name
        )
    timeout_ms = request.get("timeout_ms", 30000)
    state = {
        "command": {
            "action": action_name,
            "params": params,
            "requestId": "req-{0}".format(secrets.token_urlsafe(18)),
        },
        "result": None,
    }

    try:
        server = HTTPServer((HOST, PORT), make_handler(state))
    except OSError as error:
        if error.errno != errno.EADDRINUSE:
            raise
        raise RunnerError(
            "PORT_IN_USE",
            "Loopback port {0} is already in use".format(PORT),
            True,
            action_name,
        )
    try:
        deadline = time.monotonic() + (timeout_ms / 1000.0)
        while state["result"] is None and time.monotonic() < deadline:
            server.timeout = min(0.25, max(0.0, deadline - time.monotonic()))
            server.handle_request()
    finally:
        server.server_close()

    if state["result"] is None:
        raise RunnerError(
            "ADDIN_NOT_READY",
            "WPS Add-in did not return a result before the timeout",
            True,
            action_name,
        )

    result = state["result"]
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
        data = {key: value for key, value in result.items() if key != "success"}
    result_problem = validate_contract(data, action["result"], "result")
    if result_problem:
        raise RunnerError(
            "INVALID_RESULT", result_problem, False, action_name
        )
    return {"ok": True, "action": action_name, "data": data}


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
        try:
            install_result = addon_installer.install()
        except addon_installer.AddinInstallError as error:
            raise RunnerError(
                error.code,
                error.message,
                error.retryable,
                operation="check",
            )
        if install_result["restart_required"]:
            check_result = {
                "status": "restart_required",
                "ready": False,
                "restart_required": True,
                "platform": install_result["platform"],
                "architecture": install_result["architecture"],
            }
        elif not addon_installer.wps_is_running(install_result["platform"]):
            check_result = {
                "status": "wps_not_running",
                "ready": False,
                "restart_required": False,
                "platform": install_result["platform"],
                "architecture": install_result["architecture"],
            }
        else:
            ready = ping_addin(
                addon_installer.auth_token(install_result["platform"]),
                options["timeout_ms"],
            )
            check_result = {
                "status": "ready" if ready else "addin_unavailable",
                "ready": ready,
                "restart_required": False,
                "platform": install_result["platform"],
                "architecture": install_result["architecture"],
            }
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
