"""Black-box tests for the WPS Skill Package Python Runner."""

import json
import socket
import subprocess
import sys
import threading
import time
import unittest
from pathlib import Path
from urllib.error import URLError
from urllib.request import Request, urlopen


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
RUNNER = REPOSITORY_ROOT / "skills" / "wps-office" / "scripts" / "wps.py"
SKILL = REPOSITORY_ROOT / "skills" / "wps-office" / "SKILL.md"
POLL_URL = "http://127.0.0.1:58891/poll"
RESULT_URL = "http://127.0.0.1:58891/result"


def invoke_runner(request):
    request_json = request if isinstance(request, str) else json.dumps(request)
    return subprocess.run(
        [sys.executable, str(RUNNER), "invoke", request_json],
        cwd=REPOSITORY_ROOT,
        text=True,
        capture_output=True,
        timeout=5,
        check=False,
    )


class FakeAddin:
    """Use the public loopback protocol exactly as a WPS Add-in does."""

    def __init__(self, result):
        self.result = result
        self.command = None
        self.error = None
        self.finished = threading.Event()

    def start(self):
        thread = threading.Thread(target=self._run, daemon=True)
        thread.start()
        return thread

    def _run(self):
        try:
            deadline = time.monotonic() + 5
            while time.monotonic() < deadline:
                try:
                    with urlopen(POLL_URL, timeout=0.25) as response:
                        payload = json.load(response)
                except (OSError, URLError):
                    time.sleep(0.02)
                    continue

                if "command" not in payload:
                    time.sleep(0.02)
                    continue

                self.command = payload["command"]
                body = json.dumps(
                    {
                        "requestId": self.command["requestId"],
                        "result": self.result,
                    }
                ).encode("utf-8")
                request = Request(
                    RESULT_URL,
                    data=body,
                    headers={"Content-Type": "application/json"},
                    method="POST",
                )
                with urlopen(request, timeout=1) as response:
                    json.load(response)
                return
            raise AssertionError("Runner did not publish a command before the deadline")
        except BaseException as error:  # surfaced in the test process below
            self.error = error
        finally:
            self.finished.set()


class WpsRunnerBlackBoxTests(unittest.TestCase):
    def test_skill_documents_the_python_runner_process_contract(self):
        instructions = SKILL.read_text(encoding="utf-8")

        self.assertIn("scripts/wps.py invoke", instructions)
        self.assertIn('"action":"ping"', instructions)
        self.assertIn("stdout", instructions)
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

    def test_non_read_action_is_rejected_before_contacting_the_addin(self):
        with socket.socket() as occupied_port:
            occupied_port.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            occupied_port.bind(("127.0.0.1", 58891))
            occupied_port.listen(1)

            completed = invoke_runner(
                {
                    "action": "openFile",
                    "params": {"path": "/tmp/example.docx"},
                    "timeout_ms": 50,
                }
            )

        self.assertEqual(completed.returncode, 1)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": False,
                "action": "openFile",
                "error": {
                    "code": "ACTION_NOT_READ_ONLY",
                    "message": "WPS Action is not read-only: openFile",
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
        addin = FakeAddin(
            {"success": True, "message": "pong", "timestamp": 1723852800000}
        )
        addin.start()

        completed = invoke_runner(
            {"action": "ping", "params": {}, "timeout_ms": 2000}
        )

        self.assertTrue(addin.finished.wait(1), "Fake Add-in did not finish")
        if addin.error:
            raise addin.error
        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertEqual(
            json.loads(completed.stdout),
            {
                "ok": True,
                "action": "ping",
                "data": {"message": "pong", "timestamp": 1723852800000},
            },
        )
        self.assertEqual(completed.stderr, "")
        self.assertEqual(addin.command["action"], "ping")
        self.assertEqual(addin.command["params"], {})

        with socket.socket() as released_port:
            released_port.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            released_port.bind(("127.0.0.1", 58891))

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

        with socket.socket() as released_port:
            released_port.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            released_port.bind(("127.0.0.1", 58891))

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
