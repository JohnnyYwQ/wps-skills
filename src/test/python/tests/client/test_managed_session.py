"""Cross-process CLI usage, response recovery, and owned client lifecycle."""

import json
import os
from pathlib import Path
import signal
import subprocess
import sys
import tempfile
import time
import unittest
from unittest.mock import patch

from wps_skills.client import managed_session as managed


class ManagedSessionTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.environment = patch.dict(os.environ, {
            "WPS_SKILLS_SESSION_DIR": str(self.root / "sessions"),
            "PYTHONPATH": str(Path(managed.__file__).resolve().parents[2]),
        })
        self.environment.start()
        self.handles = []
        self.events = self.root / "actions.txt"

    def tearDown(self):
        for handle in self.handles:
            managed.close(handle, "word", timeout=3)
        self.environment.stop()
        self.temporary.cleanup()

    def start(self, mode="normal", **kwargs):
        result = managed.start("word", host_command=[sys.executable, str(Path(__file__).with_name("session_host_fixture.py")), mode, str(self.events)], **kwargs)
        # Host fixture is deliberately absent from production distributions.
        self.handles.append(result["handle"])
        return result

    def cli(self, *args, input_text=None):
        result = subprocess.run([sys.executable, "-m", "wps_skills.cli.call", "--app", "word", *args],
                                input=input_text, capture_output=True, text=True, encoding="utf-8", timeout=15)
        self.assertTrue(result.stdout, result.stderr)
        return result.returncode, json.loads(result.stdout)

    def test_separate_commands_keep_session_and_unicode_and_do_not_replay_steps(self):
        ready = self.start()
        self.assertEqual("ready", ready["state"], ready)
        handle = ready["handle"]
        parameters = self.root / "parameters.json"
        parameters.write_text(json.dumps({"text": "项目周报 😀", "path": "C:\\文件\\new.xlsx"}, ensure_ascii=False), encoding="utf-8-sig")
        args = ("--call", handle, "--step", "1", "--action", "create", "--params-file", str(parameters))
        code, first = self.cli(*args)
        self.assertEqual(0, code)
        self.assertEqual("项目周报 😀", first["response"]["data"]["echo"]["text"])
        code, duplicate = self.cli(*args)
        self.assertEqual(first, duplicate)
        code, conflict = self.cli("--call", handle, "--step", "1", "--action", "different")
        self.assertEqual(4, code)
        self.assertEqual("STEP_CONFLICT", conflict["error"]["code"])
        code, second = self.cli("--call", handle, "--step", "2", "--action", "inspect", "--params-stdin", input_text='{"text":"中文"}')
        self.assertEqual(0, code)
        self.assertEqual(first["response"]["sessionId"], second["response"]["sessionId"])
        code, closed = self.cli("--close", handle)
        self.assertEqual(0, code)
        self.assertEqual("succeeded", closed["sessionOutcome"]["outcome"])
        self.assertEqual(["create", "inspect"], self.events.read_text().splitlines())
        code, old = self.cli("--status", handle, "--step", "1")
        self.assertEqual(first, old)

    def test_failure_allows_a_deliberately_selected_read_and_cleanup_is_separate(self):
        handle = self.start("cleanup_failed")["handle"]
        code, failed = self.cli("--call", handle, "--step", "1", "--action", "fail")
        self.assertEqual(2, code)
        self.assertEqual("CONTENT_VERIFICATION_FAILED", failed["response"]["error"]["code"])
        self.assertTrue(failed["canExecute"])
        code, inspected = self.cli("--call", handle, "--step", "2", "--action", "inspect")
        self.assertEqual(0, code)
        code, closed = self.cli("--close", handle)
        self.assertEqual(4, code)
        self.assertEqual("succeeded", closed["response"]["outcome"])
        self.assertEqual("failed", closed["sessionOutcome"]["outcome"])
        self.assertIsNotNone(closed["cleanupError"])

    def test_wait_timeout_cannot_queue_next_action_and_same_step_retrieves_receipt(self):
        handle = self.start()["handle"]
        args = ("--call", handle, "--step", "1", "--action", "slow")
        code, pending = self.cli(*args, "--timeout", "0.05")
        self.assertEqual(4, code)
        self.assertTrue(pending["mayHaveEffect"])
        code, rejected = self.cli("--call", handle, "--step", "2", "--action", "inspect")
        self.assertEqual("STEP_NOT_AVAILABLE", rejected["error"]["code"])
        code, done = self.cli(*args)
        self.assertEqual(0, code)
        self.assertEqual("succeeded", done["response"]["outcome"])
        self.assertEqual(["slow"], self.events.read_text().splitlines())

    def test_terminal_unknown_never_restarts_and_retains_action_error(self):
        handle = self.start()["handle"]
        code, failed = self.cli("--call", handle, "--step", "1", "--action", "terminal")
        self.assertEqual(2, code)
        self.assertEqual("unknown", failed["response"]["outcome"])
        self.assertFalse(failed["canExecute"])
        code, rejected = self.cli("--call", handle, "--step", "2", "--action", "create")
        self.assertEqual(4, code)
        self.assertEqual(["terminal"], self.events.read_text().splitlines())

    def test_channel_loss_does_not_report_the_previous_action_as_current(self):
        handle = self.start()["handle"]
        code, first = self.cli("--call", handle, "--step", "1", "--action", "create")
        self.assertEqual(0, code)
        code, lost = self.cli("--call", handle, "--step", "2", "--action", "drop")
        self.assertEqual(4, code)
        self.assertTrue(lost["mayHaveEffect"])
        self.assertIsNone(lost["response"])
        self.assertIsNone(lost["sessionOutcome"])

    def test_concurrent_duplicate_commands_dispatch_only_once(self):
        handle = self.start()["handle"]
        command = [sys.executable, "-m", "wps_skills.cli.call", "--app", "word",
                   "--call", handle, "--step", "1", "--action", "slow"]
        processes = [subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                      text=True, encoding="utf-8") for _ in range(2)]
        results = []
        for process in processes:
            stdout, stderr = process.communicate(timeout=10)
            self.assertEqual(0, process.returncode, stderr)
            results.append(json.loads(stdout))
        self.assertEqual(results[0], results[1])
        self.assertEqual(["slow"], self.events.read_text().splitlines())

    def test_startup_failure_is_reported_without_a_usable_session(self):
        result = self.start("startup")
        self.assertEqual("closed", result["state"])
        self.assertFalse(result["canExecute"])
        self.assertEqual("SESSION_CLIENT_FAILED", result["error"]["code"])
        self.assertIn("fixture startup failed", result["stderr"])

    def test_abandoned_client_expires_without_status_extending_lifetime(self):
        handle = self.start(idle_timeout=0.3)["handle"]
        deadline = time.monotonic() + 5
        while time.monotonic() < deadline:
            result = managed.status(handle, "word")
            if result["state"] == "closed":
                break
            time.sleep(0.05)
        self.assertEqual("closed", result["state"])
        self.assertEqual("idle_timeout", result["closeReason"])
        self.assertEqual("succeeded", result["sessionOutcome"]["outcome"])
        self.assertTrue(self.events.with_suffix(".closed").exists())

    def test_bad_json_and_wrong_application_never_dispatch(self):
        handle = self.start()["handle"]
        for raw in ('[]', '{"x":1,"x":2}', '{"x":NaN}', '{"x":1e999}', '{bad'):
            code, result = self.cli("--call", handle, "--step", "1", "--action", "create", "--params-stdin", input_text=raw)
            self.assertEqual(4, code)
            self.assertEqual("CLI_REQUEST_FAILED", result["error"]["code"])
        with self.assertRaises(ValueError):
            managed.status(handle, "excel")
        self.assertFalse(self.events.exists())

    def test_worker_death_disconnects_host_and_stale_handle_cannot_restart(self):
        ready = self.start()
        if os.name == "nt":
            subprocess.run(["taskkill", "/PID", str(ready["clientPid"]), "/F"],
                           check=True, capture_output=True, timeout=10)
        else:
            os.kill(ready["clientPid"], signal.SIGKILL)
        deadline = time.monotonic() + 5
        while not self.events.with_suffix(".closed").exists() and time.monotonic() < deadline:
            time.sleep(0.05)
        self.assertTrue(self.events.with_suffix(".closed").exists())
        heartbeat = self.root / "sessions" / ready["handle"] / "heartbeat"
        old = time.time() - 20
        os.utime(heartbeat, (old, old))
        state = managed.status(ready["handle"], "word")
        self.assertEqual("unavailable", state["state"])
        result = managed.call(ready["handle"], "word", step=1, action="create", params={})
        self.assertEqual("STEP_NOT_AVAILABLE", result["error"]["code"])
        self.assertFalse(self.events.exists())
