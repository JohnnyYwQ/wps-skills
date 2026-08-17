"""Black-box tests for the WPS Skill Package Python Runner."""

import json
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
RUNNER = REPOSITORY_ROOT / "skills" / "wps-office" / "scripts" / "wps.py"
SKILL = REPOSITORY_ROOT / "skills" / "wps-office" / "SKILL.md"
POLL_URL = "http://127.0.0.1:58891/poll"
RESULT_URL = "http://127.0.0.1:58891/result"
UNKNOWN_URL = "http://127.0.0.1:58891/not-a-protocol-route"
REPRESENTATIVE_ACTIONS = json.loads(
    (REPOSITORY_ROOT / "tests/fixtures/representative-actions.json").read_text(
        encoding="utf-8"
    )
)
RUNNER_ENV = None
AUTH_TOKEN = None


def invoke_runner(request, runner=RUNNER):
    request_json = request if isinstance(request, str) else json.dumps(request)
    return subprocess.run(
        [sys.executable, str(runner), "invoke", request_json],
        cwd=REPOSITORY_ROOT,
        env=RUNNER_ENV,
        text=True,
        capture_output=True,
        timeout=5,
        check=False,
    )


def representative_action(action_name):
    return next(
        fixture
        for fixture in REPRESENTATIVE_ACTIONS
        if fixture["action"] == action_name
    )


def loopback_port_is_available():
    with socket.socket() as candidate:
        candidate.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            candidate.bind(("127.0.0.1", 58891))
        except OSError:
            return False
    return True


class PollingAddin:
    """Exercise the public loopback protocol as an external WPS Add-in client."""

    def __init__(self, auth_token=None):
        self.auth_token = AUTH_TOKEN if auth_token is None else auth_token
        self.action_request = None
        self.error = None
        self.action_received = threading.Event()
        self.finished = threading.Event()
        self.stopped = threading.Event()

    def start(self):
        thread = threading.Thread(target=self._run, daemon=True)
        thread.start()
        return thread

    def stop(self):
        self.stopped.set()

    def authorized_request(self, url, data=None, method="GET"):
        return Request(
            url,
            data=data,
            headers={
                "Authorization": "Bearer {0}".format(self.auth_token),
                "Content-Type": "application/json",
            },
            method=method,
        )

    def handle_action(self, action_request):
        raise NotImplementedError

    def send_partial_result(self, pause_seconds=0):
        with socket.create_connection(("127.0.0.1", 58891), timeout=1) as client:
            headers = (
                "POST /result HTTP/1.1\r\n"
                "Host: 127.0.0.1\r\n"
                "Authorization: Bearer {0}\r\n"
                "Content-Type: application/json\r\n"
                "Content-Length: 200\r\n"
                "Connection: close\r\n"
                "\r\n"
            ).format(self.auth_token)
            client.sendall(headers.encode("ascii") + b'{"requestId":')
            if pause_seconds:
                time.sleep(pause_seconds)

    def _run(self):
        try:
            deadline = time.monotonic() + 5
            while time.monotonic() < deadline and not self.stopped.is_set():
                try:
                    with urlopen(
                        self.authorized_request(POLL_URL), timeout=0.25
                    ) as response:
                        payload = json.load(response)
                except (OSError, URLError):
                    time.sleep(0.02)
                    continue
                if "actionRequest" not in payload:
                    time.sleep(0.02)
                    continue
                self.action_request = payload["actionRequest"]
                self.action_received.set()
                self.handle_action(self.action_request)
                return
            if not self.stopped.is_set():
                raise AssertionError(
                    "Runner did not publish a WPS Action before the deadline"
                )
        except BaseException as error:  # surfaced in the test process below
            self.error = error
        finally:
            self.finished.set()


class FakeAddin(PollingAddin):
    """Return a configured result through the public loopback protocol."""

    def __init__(self, result, auth_token=None, result_gate=None):
        super().__init__(auth_token)
        self.result = result
        self.result_gate = result_gate

    def handle_action(self, action_request):
        if self.result_gate is not None and not self.result_gate.wait(5):
            raise AssertionError("Test did not release the Add-in result")
        body = json.dumps(
            {"requestId": action_request["requestId"], "result": self.result}
        ).encode("utf-8")
        with urlopen(
            self.authorized_request(RESULT_URL, body, "POST"), timeout=1
        ) as response:
            json.load(response)


class InvalidResultPayloadAddin(PollingAddin):
    """Submit a configured invalid payload to the authenticated result endpoint."""

    def __init__(self, request_body=b"{bad json"):
        super().__init__()
        self.request_body = request_body
        self.error_response = None

    def handle_action(self, action_request):
        del action_request
        try:
            urlopen(
                self.authorized_request(RESULT_URL, self.request_body, "POST"),
                timeout=1,
            )
        except HTTPError as error:
            self.error_response = (error.code, json.load(error))
            return
        raise AssertionError("Invalid result payload was accepted")


class MismatchedRequestAddin(PollingAddin):
    """Submit a result correlated to a different WPS Action request."""

    def __init__(self):
        super().__init__()
        self.error_response = None

    def handle_action(self, action_request):
        del action_request
        body = json.dumps(
            {
                "requestId": "req-from-another-invocation",
                "result": {
                    "success": True,
                    "message": "wrong result",
                    "timestamp": 1723852800000,
                },
            }
        ).encode("utf-8")
        try:
            urlopen(
                self.authorized_request(RESULT_URL, body, "POST"), timeout=1
            )
        except HTTPError as error:
            self.error_response = (error.code, json.load(error))
            return
        raise AssertionError("Mismatched requestId was accepted")


class UnauthorizedResultAddin(PollingAddin):
    """Try an unauthenticated result before submitting the trusted result."""

    def __init__(self, result):
        super().__init__()
        self.result = result
        self.unauthorized_response = None

    def handle_action(self, action_request):
        body = json.dumps(
            {"requestId": action_request["requestId"], "result": self.result}
        ).encode("utf-8")
        request = Request(
            RESULT_URL,
            data=body,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            urlopen(request, timeout=1)
        except HTTPError as error:
            self.unauthorized_response = (error.code, json.load(error))
        else:
            raise AssertionError("Unauthenticated result was accepted")
        with urlopen(
            self.authorized_request(RESULT_URL, body, "POST"), timeout=1
        ) as response:
            json.load(response)


class DisconnectingAddin(PollingAddin):
    """Disconnect partway through an authenticated result request body."""

    def handle_action(self, action_request):
        del action_request
        self.send_partial_result()


class StallingAddin(PollingAddin):
    """Accept a WPS Action through polling but never return its result."""

    def handle_action(self, action_request):
        del action_request


class HangingResultAddin(PollingAddin):
    """Keep an authenticated result upload open beyond the Action deadline."""

    def handle_action(self, action_request):
        del action_request
        self.send_partial_result(pause_seconds=0.5)


class WpsRunnerBlackBoxTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        global AUTH_TOKEN, RUNNER_ENV
        cls.profile_directory = tempfile.TemporaryDirectory()
        profile = Path(cls.profile_directory.name)
        RUNNER_ENV = os.environ.copy()
        RUNNER_ENV.update(
            {
                "HOME": str(profile),
                "XDG_CONFIG_HOME": str(profile / ".config"),
                "XDG_DATA_HOME": str(profile / ".local" / "share"),
                "APPDATA": str(profile / "AppData" / "Roaming"),
                "WPS_SKILL_TEST_PLATFORM": "linux",
                "WPS_SKILL_TEST_ARCHITECTURE": "x86_64",
                "WPS_SKILL_TEST_WPS_RUNNING": "1",
            }
        )
        installed = subprocess.run(
            [sys.executable, str(RUNNER), "install"],
            cwd=REPOSITORY_ROOT,
            env=RUNNER_ENV,
            text=True,
            capture_output=True,
            timeout=5,
            check=False,
        )
        if installed.returncode != 0:
            raise AssertionError(installed.stdout + installed.stderr)
        config = json.loads(
            (profile / ".config/wps-office-skill/config.json").read_text(
                encoding="utf-8"
            )
        )
        AUTH_TOKEN = config["auth_token"]
        digest = json.loads(
            (
                profile
                / ".local/share/Kingsoft/wps/jsaddons/wps-office-skill_/.wps-skill-install.json"
            ).read_text(encoding="utf-8")
        )["source_digest"]
        readiness_addin = FakeAddin(
            {
                "success": True,
                "message": "pong",
                "timestamp": 1723852800000,
                "installDigest": digest,
            }
        )
        readiness_addin.start()
        checked = subprocess.run(
            [sys.executable, str(RUNNER), "check"],
            cwd=REPOSITORY_ROOT,
            env=RUNNER_ENV,
            text=True,
            capture_output=True,
            timeout=5,
            check=False,
        )
        if not readiness_addin.finished.wait(2):
            raise AssertionError(
                "Fake Add-in did not finish readiness setup: "
                + checked.stdout
                + checked.stderr
            )
        if readiness_addin.error:
            raise readiness_addin.error
        if checked.returncode != 0 or json.loads(checked.stdout)["data"]["status"] != "ready":
            raise AssertionError(checked.stdout + checked.stderr)

    @classmethod
    def tearDownClass(cls):
        global AUTH_TOKEN, RUNNER_ENV
        AUTH_TOKEN = None
        RUNNER_ENV = None
        cls.profile_directory.cleanup()

    def _invoke_with_fake_addin(self, request, result):
        addin = FakeAddin(result)
        addin.start()
        completed = invoke_runner(request)
        addin_finished = addin.finished.wait(2)
        if not addin_finished:
            addin.stop()
            addin.finished.wait(1)
        self.assertTrue(addin_finished, "Fake Add-in did not finish")
        if addin.error:
            raise addin.error
        return completed, addin

    def _assert_rejected_before_addin(self, request, expected_code):
        addin = FakeAddin({"success": True, "data": {}})
        addin.start()
        completed = invoke_runner(request)

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout)["error"]["code"], expected_code
        )
        self.assertFalse(addin.action_received.wait(0.1))
        addin.stop()
        self.assertTrue(addin.finished.wait(1))
        if addin.error:
            raise addin.error
        return completed

    def test_skill_documents_the_python_runner_process_contract(self):
        instructions = SKILL.read_text(encoding="utf-8")

        self.assertIn("scripts/wps.py invoke", instructions)
        self.assertIn('"action":"ping"', instructions)
        self.assertIn("stdout", instructions)
        self.assertIn("ACTION_BUSY", instructions)
        self.assertIn("ACTION_TIMEOUT", instructions)
        self.assertNotIn("MCP", instructions)

    def test_invalid_requests_return_structured_errors(self):
        cases = (
            ("{bad json", None, "Request must be valid JSON"),
            (
                json.dumps({"params": {}}),
                None,
                "Request field 'action' must be a string",
            ),
            (
                json.dumps({"action": "ping", "params": {}, "timeout_ms": "soon"}),
                "ping",
                "Request field 'timeout_ms' must be an integer between 1 and 300000",
            ),
            (
                json.dumps({"action": "ping", "params": {}, "confirmed": 1}),
                "ping",
                "Request field 'confirmed' must be a boolean",
            ),
        )

        for request_json, action, message in cases:
            with self.subTest(message=message):
                completed = invoke_runner(request_json)

                self.assertEqual(completed.returncode, 1)
                self.assertEqual(
                    json.loads(completed.stdout),
                    {
                        "ok": False,
                        "action": action,
                        "error": {
                            "code": "INVALID_REQUEST",
                            "message": message,
                            "retryable": False,
                        },
                    },
                )
                self.assertEqual(completed.stderr, "")

    def test_write_action_round_trips_without_a_confirmation_marker(self):
        fixture = representative_action("insertText")
        completed, addin = self._invoke_with_fake_addin(
            {
                "action": fixture["action"],
                "params": fixture["params"],
                "timeout_ms": 2000,
            },
            {"success": True, **fixture["result"]},
        )

        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": True,
                "action": "insertText",
                "data": {"position": "end", "textLength": 16},
            },
        )
        self.assertEqual(addin.action_request["action"], "insertText")
        self.assertEqual(addin.action_request["params"], fixture["params"])

    def test_powerpoint_advanced_action_round_trips_over_http(self):
        completed, addin = self._invoke_with_fake_addin(
            {
                "action": "insertPptChart",
                "params": {"slideIndex": 1, "type": "column"},
                "timeout_ms": 2000,
            },
            {"success": True, "data": {"name": "Chart 1"}},
        )

        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": True,
                "action": "insertPptChart",
                "data": {"name": "Chart 1"},
            },
        )
        self.assertEqual(addin.action_request["action"], "insertPptChart")
        self.assertEqual(
            addin.action_request["params"], {"slideIndex": 1, "type": "column"}
        )

    def test_powerpoint_image_replacement_requires_confirmation(self):
        self._assert_rejected_before_addin(
            {
                "action": "replacePptImage",
                "params": {
                    "filePath": "/tmp/replacement.png",
                    "slideIndex": 1,
                    "shapeIndex": 2,
                },
                "timeout_ms": 2000,
            },
            "CONFIRMATION_REQUIRED",
        )

    def test_powerpoint_external_slide_action_preserves_wps_failure_details(self):
        completed, _addin = self._invoke_with_fake_addin(
            {
                "action": "insertSlidesFromFile",
                "params": {"filePath": "/tmp/missing.pptx"},
                "timeout_ms": 2000,
            },
            {"success": False, "error": "source file not found"},
        )

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "insertSlidesFromFile",
                "error": {
                    "code": "WPS_ACTION_FAILED",
                    "message": "source file not found",
                    "retryable": False,
                },
            },
        )

    def test_destructive_action_without_confirmation_never_reaches_the_addin(self):
        fixture = representative_action("deleteSlide")
        addin = FakeAddin({"success": True, **fixture["result"]})
        addin.start()
        try:
            completed = invoke_runner(
                {
                    "action": fixture["action"],
                    "params": fixture["params"],
                    "timeout_ms": 2000,
                }
            )

            self.assertFalse(
                addin.action_received.wait(0.2),
                "Unconfirmed destructive Action reached the Fake Add-in",
            )
        finally:
            addin.stop()
            self.assertTrue(addin.finished.wait(1), "Fake Add-in did not stop")
        if addin.error:
            raise addin.error
        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "deleteSlide",
                "error": {
                    "code": "CONFIRMATION_REQUIRED",
                    "message": (
                        "Destructive WPS Action requires confirmed=true: deleteSlide"
                    ),
                    "retryable": False,
                },
            },
        )

    def test_confirmed_destructive_action_returns_a_structured_result(self):
        fixture = representative_action("deleteSlide")
        completed, addin = self._invoke_with_fake_addin(
            {
                "action": fixture["action"],
                "params": fixture["params"],
                "confirmed": True,
                "timeout_ms": 2000,
            },
            {"success": True, **fixture["result"]},
        )

        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": True,
                "action": "deleteSlide",
                "data": {"deleted": 2},
            },
        )
        self.assertEqual(addin.action_request["action"], "deleteSlide")
        self.assertNotIn("confirmed", addin.action_request)

    def test_unknown_or_missing_risk_is_rejected_before_contacting_the_addin(self):
        for risk in (None, "critical"):
            with self.subTest(risk=risk), tempfile.TemporaryDirectory() as directory:
                copied_skill = Path(directory) / "wps-office"
                shutil.copytree(RUNNER.parents[1], copied_skill)
                manifest_path = copied_skill / "references/action-manifest.json"
                manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
                ping = next(
                    action for action in manifest["actions"] if action["action"] == "ping"
                )
                if risk is None:
                    del ping["risk"]
                else:
                    ping["risk"] = risk
                manifest_path.write_text(json.dumps(manifest), encoding="utf-8")

                with socket.socket() as occupied_port:
                    occupied_port.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                    occupied_port.bind(("127.0.0.1", 58891))
                    occupied_port.listen(1)
                    completed = invoke_runner(
                        {"action": "ping", "params": {}, "timeout_ms": 50},
                        copied_skill / "scripts/wps.py",
                    )

                self.assertEqual(completed.returncode, 1)
                self.assertEqual(
                    json.loads(completed.stdout),
                    {
                        "ok": False,
                        "action": "ping",
                        "error": {
                            "code": "INVALID_ACTION_RISK",
                            "message": (
                                "WPS Action has an invalid or missing risk "
                                "classification: ping"
                            ),
                            "retryable": False,
                        },
                    },
                )
                self.assertEqual(completed.stderr, "")

    def test_invalid_params_are_rejected_before_contacting_the_addin(self):
        with socket.socket() as occupied_port:
            occupied_port.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            occupied_port.bind(("127.0.0.1", 58891))
            occupied_port.listen(1)

            completed = invoke_runner(
                {
                    "action": "ping",
                    "params": {"unexpected": True},
                    "timeout_ms": 50,
                }
            )

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "ping",
                "error": {
                    "code": "INVALID_PARAMS",
                    "message": "params contains unknown field 'unexpected'",
                    "retryable": False,
                },
            },
        )
        self.assertEqual(completed.stderr, "")

    def test_addin_result_must_match_the_manifest_contract(self):
        addin = FakeAddin(
            {"success": True, "message": "pong", "timestamp": "not-a-number"}
        )
        addin.start()

        completed = invoke_runner(
            {"action": "ping", "params": {}, "timeout_ms": 2000}
        )

        self.assertTrue(addin.finished.wait(1), "Fake Add-in did not finish")
        if addin.error:
            raise addin.error
        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "ping",
                "error": {
                    "code": "INVALID_RESULT",
                    "message": "result.timestamp must be a number",
                    "retryable": False,
                },
            },
        )
        self.assertEqual(completed.stderr, "")

    def test_read_action_round_trips_through_real_loopback_http(self):
        fixture = representative_action("getCellValue")
        completed, addin = self._invoke_with_fake_addin(
            {
                "action": fixture["action"],
                "params": fixture["params"],
                "timeout_ms": 2000,
            },
            {"success": True, **fixture["result"]},
        )

        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": True,
                "action": "getCellValue",
                "data": {"value": 42, "text": "42", "formula": ""},
            },
        )
        self.assertEqual(completed.stderr, "")
        self.assertEqual(addin.action_request["action"], "getCellValue")
        self.assertEqual(addin.action_request["params"], fixture["params"])

        self.assertTrue(loopback_port_is_available())

    def test_concurrent_action_and_check_return_busy_without_interrupting_owner(self):
        result_gate = threading.Event()
        addin = FakeAddin(
            {"success": True, "message": "pong", "timestamp": 1723852800000},
            result_gate=result_gate,
        )
        addin.start()
        first = subprocess.Popen(
            [
                sys.executable,
                str(RUNNER),
                "invoke",
                json.dumps(
                    {"action": "ping", "params": {}, "timeout_ms": 2000}
                ),
            ],
            cwd=REPOSITORY_ROOT,
            env=RUNNER_ENV,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        try:
            self.assertTrue(
                addin.action_received.wait(1), "First invocation was not published"
            )

            competing = invoke_runner(
                {"action": "ping", "params": {}, "timeout_ms": 2000}
            )

            self.assertEqual(competing.returncode, 1)
            self.assertEqual(
                json.loads(competing.stdout),
                {
                    "ok": False,
                    "action": "ping",
                    "error": {
                        "code": "ACTION_BUSY",
                        "message": "Another WPS Action is already running",
                        "retryable": True,
                    },
                },
            )
            self.assertEqual(competing.stderr, "")

            checked = subprocess.run(
                [sys.executable, str(RUNNER), "check"],
                cwd=REPOSITORY_ROOT,
                env=RUNNER_ENV,
                text=True,
                capture_output=True,
                timeout=3,
                check=False,
            )

            self.assertEqual(checked.returncode, 1)
            self.assertEqual(
                json.loads(checked.stdout),
                {
                    "ok": False,
                    "operation": "check",
                    "error": {
                        "code": "ACTION_BUSY",
                        "message": "Another WPS Action is already running",
                        "retryable": True,
                    },
                },
            )
            self.assertEqual(checked.stderr, "")
        finally:
            result_gate.set()
        first_stdout, first_stderr = first.communicate(timeout=3)
        self.assertTrue(addin.finished.wait(1), "Fake Add-in did not finish")
        if addin.error:
            raise addin.error
        self.assertEqual(first.returncode, 0, first_stderr)
        self.assertTrue(json.loads(first_stdout)["ok"])

    def test_excel_core_workflow_round_trips_through_real_loopback_http(self):
        steps = (
            (
                {"action": "getActiveWorkbook", "params": {}, "timeout_ms": 1000},
                {
                    "name": "Budget.xlsx",
                    "path": "/work/Budget.xlsx",
                    "sheetCount": 2,
                    "sheets": ["Raw", "Summary"],
                },
            ),
            (
                {
                    "action": "switchSheet",
                    "params": {"sheet": "Summary"},
                    "timeout_ms": 1000,
                },
                {"activeSheet": "Summary"},
            ),
            (
                {
                    "action": "setRangeData",
                    "params": {
                        "range": "A1:B2",
                        "data": [["Item", "Amount"], ["Hosting", 42]],
                    },
                    "timeout_ms": 1000,
                },
                {},
            ),
            (
                {
                    "action": "setFormula",
                    "params": {"range": "B3", "formula": "=SUM(B2:B2)"},
                    "timeout_ms": 1000,
                },
                {},
            ),
            (
                {
                    "action": "getRangeData",
                    "params": {"range": "A1:B3"},
                    "timeout_ms": 1000,
                },
                {"data": [["Item", "Amount"], ["Hosting", 42], [None, 42]]},
            ),
            (
                {"action": "save", "params": {}, "timeout_ms": 1000},
                {},
            ),
        )

        for request, data in steps:
            with self.subTest(action=request["action"]):
                completed, addin = self._invoke_with_fake_addin(
                    request, {"success": True, "data": data}
                )

                self.assertEqual(completed.returncode, 0, completed.stdout)
                self.assertEqual(
                    json.loads(completed.stdout),
                    {"ok": True, "action": request["action"], "data": data},
                )
                self.assertEqual(addin.action_request["action"], request["action"])
                self.assertEqual(addin.action_request["params"], request["params"])

    def test_empty_excel_result_is_preserved_as_structured_data(self):
        completed, _addin = self._invoke_with_fake_addin(
            {
                "action": "getRangeData",
                "params": {"range": "A1:B2"},
                "timeout_ms": 1000,
            },
            {"success": True, "data": {"data": []}},
        )

        self.assertEqual(completed.returncode, 0, completed.stdout)
        self.assertEqual(
            json.loads(completed.stdout),
            {"ok": True, "action": "getRangeData", "data": {"data": []}},
        )

    def test_invalid_excel_params_never_reach_the_addin(self):
        addin = FakeAddin({"success": True, "data": {"activeSheet": "Summary"}})
        addin.start()

        completed = invoke_runner(
            {"action": "switchSheet", "params": {}, "timeout_ms": 1000}
        )

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "switchSheet",
                "error": {
                    "code": "INVALID_PARAMS",
                    "message": "params is missing required field 'sheet'",
                    "retryable": False,
                },
            },
        )
        self.assertFalse(addin.action_received.wait(0.1))
        addin.stop()
        self.assertTrue(addin.finished.wait(1))
        if addin.error:
            raise addin.error

    def test_invalid_formula_never_reaches_the_addin(self):
        addin = FakeAddin({"success": True, "data": {}})
        addin.start()

        completed = invoke_runner(
            {
                "action": "setFormula",
                "params": {"range": "A1", "formula": "SUM(B1:B3)"},
                "timeout_ms": 1000,
            }
        )

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "setFormula",
                "error": {
                    "code": "INVALID_PARAMS",
                    "message": "params.formula must match pattern '^='",
                    "retryable": False,
                },
            },
        )
        self.assertFalse(addin.action_received.wait(0.1))
        addin.stop()
        self.assertTrue(addin.finished.wait(1))
        if addin.error:
            raise addin.error

    def test_excel_wps_failure_returns_a_structured_error(self):
        completed, _addin = self._invoke_with_fake_addin(
            {
                "action": "getCellValue",
                "params": {"sheet": "Summary", "row": 1, "col": 1},
                "timeout_ms": 1000,
            },
            {"success": False, "error": "Worksheet was not found"},
        )

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "getCellValue",
                "error": {
                    "code": "WPS_ACTION_FAILED",
                    "message": "Worksheet was not found",
                    "retryable": False,
                },
            },
        )

    def test_destructive_clean_data_requires_confirmation_before_the_addin(self):
        addin = FakeAddin(
            {
                "success": True,
                "data": {
                    "range": "A1:B10",
                    "operations": [],
                    "message": "cleanData completed",
                },
            }
        )
        addin.start()

        completed = invoke_runner(
            {
                "action": "cleanData",
                "params": {"range": "A1:B10", "operations": ["trim"]},
                "timeout_ms": 1000,
            }
        )

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(json.loads(completed.stdout)["error"]["code"], "CONFIRMATION_REQUIRED")
        self.assertFalse(addin.action_received.wait(0.1))
        addin.stop()
        self.assertTrue(addin.finished.wait(1))
        if addin.error:
            raise addin.error

    def test_clean_data_rejects_unknown_operations_before_the_addin(self):
        addin = FakeAddin({"success": True, "data": {}})
        addin.start()

        completed = invoke_runner(
            {
                "action": "cleanData",
                "params": {"range": "A1:B10", "operations": ["guess_dates"]},
                "confirmed": True,
                "timeout_ms": 1000,
            }
        )

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(json.loads(completed.stdout)["error"]["code"], "INVALID_PARAMS")
        self.assertFalse(addin.action_received.wait(0.1))
        addin.stop()
        self.assertTrue(addin.finished.wait(1))
        if addin.error:
            raise addin.error

    def test_clear_range_rejects_unknown_types_before_the_addin(self):
        addin = FakeAddin({"success": True, "data": {}})
        addin.start()

        completed = invoke_runner(
            {
                "action": "clearRange",
                "params": {"range": "A1:B10", "type": "content"},
                "confirmed": True,
                "timeout_ms": 1000,
            }
        )

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(json.loads(completed.stdout)["error"]["code"], "INVALID_PARAMS")
        self.assertFalse(addin.action_received.wait(0.1))
        addin.stop()
        self.assertTrue(addin.finished.wait(1))
        if addin.error:
            raise addin.error

    def test_cell_info_result_matches_the_manifest_contract(self):
        data = {
            "cell": "A1",
            "value": 42,
            "formula": "",
            "numberFormat": "0",
            "font": {"name": "Arial", "size": 11, "bold": False},
            "backgroundColor": 16777215,
        }
        completed, _addin = self._invoke_with_fake_addin(
            {
                "action": "getCellInfo",
                "params": {"cell": "A1"},
                "timeout_ms": 1000,
            },
            {"success": True, "data": data},
        )

        self.assertEqual(completed.returncode, 0, completed.stdout)
        self.assertEqual(
            json.loads(completed.stdout),
            {"ok": True, "action": "getCellInfo", "data": data},
        )

    def test_excel_analysis_and_presentation_workflow_round_trips(self):
        steps = (
            (
                {
                    "action": "setCellStyle",
                    "params": {
                        "range": "A1:B3",
                        "bold": True,
                        "backgroundColor": "#1F4E78",
                    },
                    "timeout_ms": 1000,
                },
                {"range": "A1:B3"},
            ),
            (
                {
                    "action": "createChart",
                    "params": {
                        "dataRange": "A1:B3",
                        "chartType": "column",
                        "title": "Quarterly revenue",
                    },
                    "timeout_ms": 1000,
                },
                {
                    "chartName": "Chart 1",
                    "chartIndex": 1,
                    "dataRange": "A1:B3",
                    "chartType": "column",
                    "position": {"left": 300, "top": 20, "width": 400, "height": 300},
                },
            ),
            (
                {
                    "action": "updateChart",
                    "params": {"chartName": "Chart 1", "showLegend": False},
                    "timeout_ms": 1000,
                },
                {"chartName": "Chart 1", "updatedProperties": ["showLegend"]},
            ),
            (
                {
                    "action": "createPivotTable",
                    "params": {
                        "sourceRange": "A1:C20",
                        "destinationCell": "E1",
                        "rowFields": ["Region"],
                        "valueFields": [{"field": "Revenue", "function": "sum"}],
                    },
                    "timeout_ms": 1000,
                },
                {
                    "pivotTableName": "PivotTable1",
                    "location": "E1",
                    "sheet": "Analysis",
                    "rowCount": 5,
                    "columnCount": 2,
                },
            ),
            (
                {
                    "action": "updatePivotTable",
                    "params": {
                        "pivotTableName": "PivotTable1",
                        "sheet": "Analysis",
                        "refresh": True,
                    },
                    "timeout_ms": 1000,
                },
                {
                    "pivotTableName": "PivotTable1",
                    "operations": [
                        {"operation": "refresh", "success": True, "message": "refreshed"}
                    ],
                },
            ),
            (
                {"action": "getExcelContext", "params": {}, "timeout_ms": 1000},
                {
                    "workbookName": "Revenue.xlsx",
                    "currentSheet": "Analysis",
                    "allSheets": ["Raw", "Analysis"],
                    "selectedCell": "E1",
                    "usedRange": "A1:F20",
                    "usedRangeAddress": "A1:F20",
                    "headers": [{"column": "A", "value": "Region"}],
                    "rowCount": 20,
                    "colCount": 6,
                },
            ),
        )

        for request, data in steps:
            with self.subTest(action=request["action"]):
                completed, addin = self._invoke_with_fake_addin(
                    request, {"success": True, "data": data}
                )

                self.assertEqual(completed.returncode, 0, completed.stdout)
                self.assertEqual(
                    json.loads(completed.stdout),
                    {"ok": True, "action": request["action"], "data": data},
                )
                self.assertEqual(addin.action_request["params"], request["params"])

    def test_invalid_zoom_never_reaches_the_addin(self):
        addin = FakeAddin({"success": True, "data": {"percent": 500}})
        addin.start()

        completed = invoke_runner(
            {
                "action": "setZoom",
                "params": {"percent": 500},
                "timeout_ms": 1000,
            }
        )

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout)["error"],
            {
                "code": "INVALID_PARAMS",
                "message": "params.percent must be at most 400",
                "retryable": False,
            },
        )
        self.assertFalse(addin.action_received.wait(0.1))
        addin.stop()
        self.assertTrue(addin.finished.wait(1))
        if addin.error:
            raise addin.error

    def test_advanced_overwrite_requires_confirmation_before_the_addin(self):
        addin = FakeAddin(
            {"success": True, "data": {"cell": "A1", "text": "replacement"}}
        )
        addin.start()

        completed = invoke_runner(
            {
                "action": "addCellComment",
                "params": {"cell": "A1", "text": "replacement"},
                "timeout_ms": 1000,
            }
        )

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout)["error"]["code"], "CONFIRMATION_REQUIRED"
        )
        self.assertFalse(addin.action_received.wait(0.1))
        addin.stop()
        self.assertTrue(addin.finished.wait(1))
        if addin.error:
            raise addin.error

    def test_word_document_workflow_round_trips_through_real_loopback_http(self):
        steps = (
            (
                {"action": "createDocument", "params": {}, "timeout_ms": 2000},
                {"name": "Document1", "path": ""},
            ),
            (
                {
                    "action": "insertText",
                    "params": {"text": "Quarterly report", "position": "end"},
                    "timeout_ms": 2000,
                },
                {"position": "end", "textLength": 16},
            ),
            (
                {
                    "action": "getDocumentParagraphs",
                    "params": {"startParagraph": 1, "endParagraph": 2},
                    "timeout_ms": 2000,
                },
                {
                    "paragraphs": [
                        {
                            "index": 1,
                            "text": "Quarterly report",
                            "style": "Normal",
                            "start": 0,
                            "end": 17,
                        }
                    ],
                    "totalCount": 1,
                    "returnedCount": 1,
                },
            ),
            (
                {
                    "action": "findInDocument",
                    "params": {"findText": "Quarterly"},
                    "timeout_ms": 2000,
                },
                {
                    "results": [
                        {
                            "text": "Quarterly",
                            "start": 0,
                            "end": 9,
                            "paragraphIndex": 1,
                            "context": "Quarterly report",
                        }
                    ],
                    "count": 1,
                    "findText": "Quarterly",
                },
            ),
            (
                {
                    "action": "enableTrackChanges",
                    "params": {"enable": True},
                    "timeout_ms": 2000,
                },
                {"trackChanges": True, "active": True},
            ),
            (
                {
                    "action": "replaceRange",
                    "params": {"startPos": 0, "endPos": 9, "text": "Monthly"},
                    "confirmed": True,
                    "timeout_ms": 2000,
                },
                {
                    "startPos": 0,
                    "originalEndPos": 9,
                    "endPos": 7,
                    "originalText": "Quarterly",
                    "newText": "Monthly",
                },
            ),
            (
                {
                    "action": "getTrackChangesStatus",
                    "params": {},
                    "timeout_ms": 2000,
                },
                {"trackChanges": True, "revisionCount": 0},
            ),
            (
                {"action": "save", "params": {}, "timeout_ms": 2000},
                {},
            ),
            (
                {
                    "action": "closeDocument",
                    "params": {"saveChanges": True},
                    "confirmed": True,
                    "timeout_ms": 2000,
                },
                {"closed": "Document1"},
            ),
        )

        for request, data in steps:
            with self.subTest(action=request["action"]):
                completed, addin = self._invoke_with_fake_addin(
                    request, {"success": True, "data": data}
                )

                self.assertEqual(completed.returncode, 0, completed.stdout)
                self.assertEqual(
                    json.loads(completed.stdout),
                    {"ok": True, "action": request["action"], "data": data},
                )
                self.assertEqual(addin.action_request["params"], request["params"])

    def test_powerpoint_core_authoring_workflow_round_trips_through_real_http(self):
        steps = (
            (
                {"action": "createPresentation", "params": {}, "timeout_ms": 2000},
                {"name": "Presentation1", "path": "", "slideCount": 0},
            ),
            (
                {
                    "action": "addSlide",
                    "params": {
                        "layout": "title_content",
                        "title": "Quarterly review",
                        "content": "Highlights",
                    },
                    "timeout_ms": 2000,
                },
                {"slideIndex": 1, "layout": "title_content"},
            ),
            (
                {
                    "action": "setSlideContent",
                    "params": {"slideIndex": 1, "content": "Revenue grew 18%."},
                    "timeout_ms": 2000,
                },
                {"slideIndex": 1, "content": "Revenue grew 18%."},
            ),
            (
                {
                    "action": "insertPptImage",
                    "params": {"slideIndex": 1, "path": "/tmp/chart.png"},
                    "timeout_ms": 2000,
                },
                {"name": "Picture 1", "path": "/tmp/chart.png", "slideIndex": 1},
            ),
            (
                {
                    "action": "insertPptTable",
                    "params": {"slideIndex": 1, "rows": 2, "cols": 2},
                    "timeout_ms": 2000,
                },
                {
                    "name": "Table 1",
                    "rows": 2,
                    "cols": 2,
                    "slideIndex": 1,
                },
            ),
            (
                {
                    "action": "getSlideInfo",
                    "params": {"slideIndex": 1},
                    "timeout_ms": 2000,
                },
                {
                    "slideIndex": 1,
                    "layout": 2,
                    "shapeCount": 4,
                    "shapes": [
                        {
                            "index": 1,
                            "name": "Title 1",
                            "type": 14,
                            "hasText": True,
                            "text": "Quarterly review",
                        }
                    ],
                },
            ),
            (
                {"action": "save", "params": {}, "timeout_ms": 2000},
                {},
            ),
        )

        for request, data in steps:
            with self.subTest(action=request["action"]):
                completed, addin = self._invoke_with_fake_addin(
                    request, {"success": True, "data": data}
                )

                self.assertEqual(completed.returncode, 0, completed.stdout)
                self.assertEqual(
                    json.loads(completed.stdout),
                    {"ok": True, "action": request["action"], "data": data},
                )
                self.assertEqual(addin.action_request["params"], request["params"])

    def test_destructive_powerpoint_edits_require_confirmation_before_addin(self):
        for request in (
            {
                "action": "deleteSlide",
                "params": {"slideIndex": 2},
                "timeout_ms": 1000,
            },
        ):
            with self.subTest(action=request["action"]):
                self._assert_rejected_before_addin(
                    request, "CONFIRMATION_REQUIRED"
                )

    def test_invalid_powerpoint_slide_index_never_reaches_the_addin(self):
        self._assert_rejected_before_addin(
            {
                "action": "getSlideInfo",
                "params": {"slideIndex": 0},
                "timeout_ms": 1000,
            },
            "INVALID_PARAMS",
        )

    def test_powerpoint_table_row_style_result_matches_its_contract(self):
        data = {"row": 1, "cols": 3}
        completed, addin = self._invoke_with_fake_addin(
            {
                "action": "setPptTableRowStyle",
                "params": {
                    "slideIndex": 1,
                    "tableName": "Table 1",
                    "row": 1,
                    "alignment": "center",
                },
                "timeout_ms": 1000,
            },
            {"success": True, "data": data},
        )

        self.assertEqual(completed.returncode, 0, completed.stdout)
        self.assertEqual(
            json.loads(completed.stdout),
            {"ok": True, "action": "setPptTableRowStyle", "data": data},
        )
        self.assertEqual(addin.action_request["params"]["alignment"], "center")

    def test_powerpoint_table_cell_accepts_the_value_alias(self):
        data = {"row": 1, "col": 1, "text": "Revenue"}
        completed, addin = self._invoke_with_fake_addin(
            {
                "action": "setPptTableCell",
                "params": {
                    "slideIndex": 1,
                    "tableName": "Table 1",
                    "row": 1,
                    "col": 1,
                    "value": "Revenue",
                },
                "timeout_ms": 1000,
            },
            {"success": True, "data": data},
        )

        self.assertEqual(completed.returncode, 0, completed.stdout)
        self.assertEqual(
            json.loads(completed.stdout),
            {"ok": True, "action": "setPptTableCell", "data": data},
        )
        self.assertEqual(addin.action_request["params"]["value"], "Revenue")

    def test_powerpoint_wps_failure_returns_a_structured_error(self):
        completed, _addin = self._invoke_with_fake_addin(
            {
                "action": "getSlideInfo",
                "params": {"slideIndex": 1},
                "timeout_ms": 1000,
            },
            {"success": False, "error": "No active presentation"},
        )

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "getSlideInfo",
                "error": {
                    "code": "WPS_ACTION_FAILED",
                    "message": "No active presentation",
                    "retryable": False,
                },
            },
        )

    def test_destructive_word_replacements_require_confirmation_before_the_addin(self):
        requests = (
            {
                "action": "replaceRange",
                "params": {"startPos": 0, "endPos": 9, "text": "Monthly"},
                "timeout_ms": 1000,
            },
            {
                "action": "replaceBookmarkContent",
                "params": {"name": "project_name", "text": "Apollo"},
                "timeout_ms": 1000,
            },
            {
                "action": "smartFillField",
                "params": {"keyword": "Project", "value": "Apollo"},
                "timeout_ms": 1000,
            },
        )

        for request in requests:
            with self.subTest(action=request["action"]):
                self._assert_rejected_before_addin(
                    request, "CONFIRMATION_REQUIRED"
                )

    def test_invalid_word_paragraph_range_never_reaches_the_addin(self):
        self._assert_rejected_before_addin(
            {
                "action": "getDocumentParagraphs",
                "params": {"startParagraph": 0},
                "timeout_ms": 1000,
            },
            "INVALID_PARAMS",
        )

    def test_word_wps_failure_returns_a_structured_error(self):
        completed, _addin = self._invoke_with_fake_addin(
            {
                "action": "findInDocument",
                "params": {"findText": "missing"},
                "timeout_ms": 1000,
            },
            {"success": False, "error": "No active Word document"},
        )

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "findInDocument",
                "error": {
                    "code": "WPS_ACTION_FAILED",
                    "message": "No active Word document",
                    "retryable": False,
                },
            },
        )

    def test_process_exit_releases_the_lock_even_when_the_lock_file_remains(self):
        result_gate = threading.Event()
        addin = FakeAddin(
            {"success": True, "message": "pong", "timestamp": 1723852800000},
            result_gate=result_gate,
        )
        addin.start()
        owner = subprocess.Popen(
            [
                sys.executable,
                str(RUNNER),
                "invoke",
                json.dumps(
                    {"action": "ping", "params": {}, "timeout_ms": 2000}
                ),
            ],
            cwd=REPOSITORY_ROOT,
            env=RUNNER_ENV,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        self.assertTrue(
            addin.action_received.wait(1), "Owner invocation was not published"
        )

        owner.terminate()
        owner.communicate(timeout=3)
        result_gate.set()
        self.assertTrue(addin.finished.wait(2), "Fake Add-in did not finish")
        lock_file = (
            Path(RUNNER_ENV["XDG_CONFIG_HOME"])
            / "wps-office-skill"
            / "action.lock"
        )
        self.assertTrue(lock_file.is_file())

        after_exit = invoke_runner(
            {"action": "ping", "params": {}, "timeout_ms": 20}
        )

        self.assertEqual(
            json.loads(after_exit.stdout)["error"]["code"], "ADDIN_NOT_READY"
        )

    def test_addin_failure_returns_structured_error_and_nonzero_exit(self):
        addin = FakeAddin(
            {"success": False, "error": "Fake Add-in refused the Action"}
        )
        addin.start()

        completed = invoke_runner(
            {"action": "ping", "params": {}, "timeout_ms": 2000}
        )

        self.assertTrue(addin.finished.wait(1), "Fake Add-in did not finish")
        if addin.error:
            raise addin.error
        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "ping",
                "error": {
                    "code": "WPS_ACTION_FAILED",
                    "message": "Fake Add-in refused the Action",
                    "retryable": False,
                },
            },
        )
        self.assertEqual(completed.stderr, "")

    def test_malformed_addin_json_returns_a_stable_error_and_releases_port(self):
        addin = InvalidResultPayloadAddin()
        addin.start()

        completed = invoke_runner(
            {"action": "ping", "params": {}, "timeout_ms": 2000}
        )

        self.assertTrue(addin.finished.wait(1), "Faulting Add-in did not finish")
        if addin.error:
            raise addin.error
        self.assertEqual(
            addin.error_response,
            (
                400,
                {
                    "ok": False,
                    "error": {
                        "code": "INVALID_ADDIN_JSON",
                        "message": "WPS Add-in result must be valid JSON",
                    },
                },
            ),
        )
        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "ping",
                "error": {
                    "code": "INVALID_ADDIN_JSON",
                    "message": "WPS Add-in result must be valid JSON",
                    "retryable": False,
                },
            },
        )
        self.assertEqual(completed.stderr, "")
        self.assertTrue(loopback_port_is_available())

    def test_non_object_addin_json_returns_a_stable_protocol_error(self):
        addin = InvalidResultPayloadAddin(b"[]")
        addin.start()

        completed = invoke_runner(
            {"action": "ping", "params": {}, "timeout_ms": 2000}
        )

        self.assertTrue(addin.finished.wait(1), "Faulting Add-in did not finish")
        if addin.error:
            raise addin.error
        expected_error = {
            "code": "INVALID_ADDIN_RESPONSE",
            "message": "WPS Add-in result payload must be a JSON object",
        }
        self.assertEqual(
            addin.error_response,
            (400, {"ok": False, "error": expected_error}),
        )
        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "ping",
                "error": dict(expected_error, retryable=False),
            },
        )
        self.assertEqual(completed.stderr, "")

    def test_mismatched_request_id_is_rejected_and_never_used_as_the_result(self):
        addin = MismatchedRequestAddin()
        addin.start()

        completed = invoke_runner(
            {"action": "ping", "params": {}, "timeout_ms": 2000}
        )

        self.assertTrue(addin.finished.wait(1), "Faulting Add-in did not finish")
        if addin.error:
            raise addin.error
        expected_error = {
            "code": "REQUEST_ID_MISMATCH",
            "message": "WPS Add-in result does not match the pending WPS Action",
        }
        self.assertEqual(
            addin.error_response,
            (409, {"ok": False, "error": expected_error}),
        )
        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "ping",
                "error": dict(expected_error, retryable=False),
            },
        )
        self.assertEqual(completed.stderr, "")

    def test_disconnected_result_upload_has_a_stable_error_and_cleans_up(self):
        addin = DisconnectingAddin()
        addin.start()

        completed = invoke_runner(
            {"action": "ping", "params": {}, "timeout_ms": 2000}
        )

        self.assertTrue(addin.finished.wait(1), "Faulting Add-in did not finish")
        if addin.error:
            raise addin.error
        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "ping",
                "error": {
                    "code": "ADDIN_DISCONNECTED",
                    "message": "WPS Add-in disconnected while sending its result",
                    "retryable": False,
                },
            },
        )
        self.assertEqual(completed.stderr, "")
        self.assertNotIn(AUTH_TOKEN, completed.stdout + completed.stderr)
        self.assertTrue(loopback_port_is_available())

        after_failure = invoke_runner(
            {"action": "ping", "params": {}, "timeout_ms": 20}
        )
        self.assertEqual(
            json.loads(after_failure.stdout)["error"]["code"], "ADDIN_NOT_READY"
        )

    def test_missing_addin_returns_retryable_timeout_error(self):
        completed = invoke_runner({"action": "ping", "params": {}, "timeout_ms": 50})

        self.assertEqual(completed.returncode, 1)
        self.assertTrue(completed.stdout, completed.stderr)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "ping",
                "error": {
                    "code": "ADDIN_NOT_READY",
                    "message": "WPS Add-in did not return a result before the timeout",
                    "retryable": True,
                },
            },
        )
        self.assertEqual(completed.stderr, "")

        self.assertTrue(loopback_port_is_available())

    def test_stalled_action_is_distinct_from_an_addin_that_never_polled(self):
        addin = StallingAddin()
        addin.start()

        completed = invoke_runner(
            {"action": "ping", "params": {}, "timeout_ms": 200}
        )

        self.assertTrue(addin.finished.wait(1), "Stalling Add-in did not poll")
        if addin.error:
            raise addin.error
        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "ping",
                "error": {
                    "code": "ACTION_TIMEOUT",
                    "message": "WPS Action did not return a result before the timeout",
                    "retryable": True,
                },
            },
        )
        self.assertEqual(completed.stderr, "")

    def test_action_deadline_bounds_a_stalled_result_upload(self):
        addin = HangingResultAddin()
        addin.start()
        started = time.monotonic()

        completed = invoke_runner(
            {"action": "ping", "params": {}, "timeout_ms": 200}
        )
        elapsed = time.monotonic() - started

        self.assertTrue(addin.finished.wait(1), "Hanging Add-in did not finish")
        if addin.error:
            raise addin.error
        self.assertLess(elapsed, 0.45)
        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "ping",
                "error": {
                    "code": "ACTION_TIMEOUT",
                    "message": "WPS Action did not return a result before the timeout",
                    "retryable": True,
                },
            },
        )
        self.assertEqual(completed.stderr, "")

    def test_action_polling_rejects_a_different_install_credential(self):
        addin = FakeAddin(
            {"success": True, "message": "pong", "timestamp": 1723852800000},
            auth_token="wrong-credential-that-is-long-enough",
        )
        addin.start()

        completed = invoke_runner(
            {"action": "ping", "params": {}, "timeout_ms": 100}
        )
        addin.stop()

        self.assertTrue(addin.finished.wait(1), "Fake Add-in did not stop")
        if addin.error:
            raise addin.error
        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout)["error"]["code"], "ADDIN_NOT_READY"
        )
        self.assertIsNone(addin.action_request)

    def test_unauthorized_poll_is_structured_and_does_not_consume_the_action(self):
        runner = subprocess.Popen(
            [
                sys.executable,
                str(RUNNER),
                "invoke",
                json.dumps(
                    {"action": "ping", "params": {}, "timeout_ms": 2000}
                ),
            ],
            cwd=REPOSITORY_ROOT,
            env=RUNNER_ENV,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        unauthorized_response = None
        deadline = time.monotonic() + 2
        while time.monotonic() < deadline:
            try:
                request = Request(
                    POLL_URL,
                    headers={"Authorization": "Bearer invalid-credential"},
                )
                urlopen(request, timeout=0.25)
            except HTTPError as error:
                unauthorized_response = (error.code, json.load(error))
                break
            except (OSError, URLError):
                time.sleep(0.02)
        self.assertEqual(
            unauthorized_response,
            (
                401,
                {
                    "ok": False,
                    "error": {
                        "code": "AUTHENTICATION_FAILED",
                        "message": "Loopback request authentication failed",
                    },
                },
            ),
        )

        addin = FakeAddin(
            {"success": True, "message": "pong", "timestamp": 1723852800000}
        )
        addin.start()
        stdout, stderr = runner.communicate(timeout=3)

        self.assertTrue(addin.finished.wait(1), "Fake Add-in did not finish")
        if addin.error:
            raise addin.error
        self.assertEqual(runner.returncode, 0, stderr)
        self.assertTrue(json.loads(stdout)["ok"])
        self.assertNotIn(AUTH_TOKEN, stdout + stderr)

    def test_all_protected_routes_authenticate_and_preflight_exposes_no_action(self):
        runner = subprocess.Popen(
            [
                sys.executable,
                str(RUNNER),
                "invoke",
                json.dumps(
                    {"action": "ping", "params": {}, "timeout_ms": 2000}
                ),
            ],
            cwd=REPOSITORY_ROOT,
            env=RUNNER_ENV,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        preflight_payload = None
        deadline = time.monotonic() + 2
        while time.monotonic() < deadline:
            try:
                request = Request(POLL_URL, method="OPTIONS")
                with urlopen(request, timeout=0.25) as response:
                    preflight_payload = json.load(response)
                break
            except (OSError, URLError):
                time.sleep(0.02)
        self.assertEqual(preflight_payload, {})

        try:
            urlopen(Request(UNKNOWN_URL), timeout=0.25)
        except HTTPError as error:
            unknown_response = (error.code, json.load(error))
        else:
            self.fail("Unauthenticated unknown route was accepted")
        expected_auth_error = {
            "ok": False,
            "error": {
                "code": "AUTHENTICATION_FAILED",
                "message": "Loopback request authentication failed",
            },
        }
        self.assertEqual(unknown_response, (401, expected_auth_error))

        addin = UnauthorizedResultAddin(
            {"success": True, "message": "pong", "timestamp": 1723852800000}
        )
        addin.start()
        stdout, stderr = runner.communicate(timeout=3)

        self.assertTrue(addin.finished.wait(1), "Fake Add-in did not finish")
        if addin.error:
            raise addin.error
        self.assertEqual(addin.unauthorized_response, (401, expected_auth_error))
        self.assertEqual(runner.returncode, 0, stderr)
        self.assertTrue(json.loads(stdout)["ok"])

    def test_unknown_action_is_rejected_before_contacting_the_addin(self):
        with socket.socket() as occupied_port:
            occupied_port.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            occupied_port.bind(("127.0.0.1", 58891))
            occupied_port.listen(1)

            completed = invoke_runner(
                {"action": "notARealAction", "params": {}, "timeout_ms": 50}
            )

        self.assertEqual(completed.returncode, 1)
        self.assertTrue(completed.stdout, completed.stderr)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "notARealAction",
                "error": {
                    "code": "UNKNOWN_ACTION",
                    "message": "Unknown WPS Action: notARealAction",
                    "retryable": False,
                },
            },
        )
        self.assertEqual(completed.stderr, "")

    def test_occupied_loopback_port_returns_structured_error(self):
        with socket.socket() as occupied_port:
            occupied_port.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            occupied_port.bind(("127.0.0.1", 58891))
            occupied_port.listen(1)

            completed = invoke_runner(
                {"action": "ping", "params": {}, "timeout_ms": 50}
            )

        self.assertEqual(completed.returncode, 1)
        self.assertTrue(completed.stdout, completed.stderr)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "ping",
                "error": {
                    "code": "PORT_IN_USE",
                    "message": "Loopback port 58891 is already in use",
                    "retryable": True,
                },
            },
        )
        self.assertEqual(completed.stderr, "")


if __name__ == "__main__":
    unittest.main()
