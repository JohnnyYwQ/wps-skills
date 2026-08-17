"""Black-box tests for the WPS Skill Package Python Runner."""

import json
import os
import socket
import subprocess
import sys
import tempfile
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
RUNNER_ENV = None
AUTH_TOKEN = None


def invoke_runner(request):
    request_json = request if isinstance(request, str) else json.dumps(request)
    return subprocess.run(
        [sys.executable, str(RUNNER), "invoke", request_json],
        cwd=REPOSITORY_ROOT,
        env=RUNNER_ENV,
        text=True,
        capture_output=True,
        timeout=5,
        check=False,
    )


class FakeAddin:
    """Use the public loopback protocol exactly as a WPS Add-in does."""

    def __init__(self, result, auth_token=None):
        self.result = result
        self.auth_token = auth_token or AUTH_TOKEN
        self.action_request = None
        self.error = None
        self.finished = threading.Event()
        self.stopped = threading.Event()

    def start(self):
        thread = threading.Thread(target=self._run, daemon=True)
        thread.start()
        return thread

    def stop(self):
        self.stopped.set()

    def _run(self):
        try:
            deadline = time.monotonic() + 5
            while time.monotonic() < deadline and not self.stopped.is_set():
                try:
                    poll_request = Request(
                        POLL_URL,
                        headers={
                            "Authorization": "Bearer {0}".format(self.auth_token)
                        },
                    )
                    with urlopen(poll_request, timeout=0.25) as response:
                        payload = json.load(response)
                except (OSError, URLError):
                    time.sleep(0.02)
                    continue

                if "actionRequest" not in payload:
                    time.sleep(0.02)
                    continue

                self.action_request = payload["actionRequest"]
                body = json.dumps(
                    {
                        "requestId": self.action_request["requestId"],
                        "result": self.result,
                    }
                ).encode("utf-8")
                request = Request(
                    RESULT_URL,
                    data=body,
                    headers={"Content-Type": "application/json"},
                    method="POST",
                )
                request.add_header(
                    "Authorization", "Bearer {0}".format(self.auth_token)
                )
                with urlopen(request, timeout=1) as response:
                    json.load(response)
                return
            if not self.stopped.is_set():
                raise AssertionError(
                    "Runner did not publish a WPS Action before the deadline"
                )
        except BaseException as error:  # surfaced in the test process below
            self.error = error
        finally:
            self.finished.set()


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
        if not readiness_addin.finished.wait(1):
            raise AssertionError("Fake Add-in did not finish readiness setup")
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
        self.assertEqual(addin.action_request["action"], "ping")
        self.assertEqual(addin.action_request["params"], {})

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
