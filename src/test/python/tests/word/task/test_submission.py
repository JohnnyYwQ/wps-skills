"""Word Task: code-owned identity, Action verification and input lifecycle."""

import copy
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from wps_skills.cli.call import main
from wps_skills.client import task_client, task_request, task_store
from wps_skills.client.json_io import _write
from wps_skills.client.task_file import ConsumableTaskFile
from tests.word.task.test_plan_execution import RecordingExecutor, action, step, ref


class WordSubmissionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / "input.json"
        env = patch.dict(os.environ, {"WPS_SKILLS_TASK_DIR": str(self.root / "receipts")})
        env.start()
        self.addCleanup(env.stop)
        self.request = {
            "app": "word", "document": action("createDocument"),
            "steps": [step("write", "writeContent", {
                "anchor": {"kind": "documentEnd"},
                "blocks": [{"kind": "paragraph", "runs": [{"text": "中文、English 与🙂"}]}],
            })],
            "completion": [action("saveAs", {
                "outputPath": "C:/work/output.docx", "overwritePolicy": "renameIfExists",
            })],
        }

    def write_input(self, request=None):
        self.source.write_text(json.dumps(request or self.request, ensure_ascii=False), encoding="utf-8-sig")

    def cli(self, *args, client=None):
        output, errors = io.StringIO(), io.StringIO()
        self.client = client or RecordingExecutor()
        with patch("wps_skills.word.windows.task_factory.build_task", return_value=self.client):
            code = main(["--app", "word", *args], output_stream=output, error_stream=errors)
        return code, json.loads(output.getvalue())

    def submit(self, client=None):
        return self.cli("--task-file", str(self.source), client=client)

    def test_submission_needs_no_id_checks_or_cleanup_flags(self):
        self.write_input()
        code, result = self.submit()
        self.assertEqual(0, code, result)
        self.assertEqual(["createDocument", "writeContent", "saveAs"], [n for n, _ in self.client.calls])
        self.assertEqual("succeeded", result["outcome"])
        self.assertEqual("succeeded", result["cleanup"]["outcome"])
        self.assertFalse(self.source.exists())
        self.assertEqual("removed", result.pop("taskFile")["state"])
        receipt = Path(result["recordPath"])
        self.assertEqual(self.request, json.loads(receipt.with_name("request.json").read_text()))
        self.assertEqual(result, json.loads(receipt.read_text()))
        self.assertNotIn("sessionOutcome", result)
        self.assertNotIn("ready", result)
        self.assertNotIn("sessionId", json.dumps(result))
        for record in task_request.result_steps(result):
            self.assertNotIn("checks", record)
            self.assertNotIn("sessionId", record["response"])

    def test_submission_compiles_once_and_reuses_the_plan_for_receipts(self):
        from wps_skills.word.task import plan
        self.write_input()
        with patch.object(plan, "compile_request", wraps=plan.compile_request) as compile_request:
            code, result = self.submit()
        self.assertEqual(0, code, result)
        self.assertEqual(1, compile_request.call_count)
        self.assertEqual(["createDocument", "writeContent", "saveAs"], [n for n, _ in self.client.calls])
        self.assertEqual("succeeded", result["completion"]["save"]["state"])

    def test_lost_response_can_be_queried_by_consumed_input_path_without_replay(self):
        self.write_input()
        _, first = self.submit()
        first.pop("taskFile")
        for arguments in (("--task-status-file", str(self.source)),
                          ("--task-file", str(self.source)),
                          ("--task-status", first["taskId"])):
            with self.subTest(arguments=arguments):
                code, result = self.cli(*arguments)
                self.assertEqual(0, code, result)
                self.assertEqual(first, result)
                self.assertEqual([], self.client.calls)

    def test_restored_same_request_only_reads_receipt_and_consumes_input(self):
        self.write_input()
        _, first = self.submit()
        self.write_input()
        _, second = self.submit()
        self.assertEqual(first, second)
        self.assertEqual([], self.client.calls)
        self.assertFalse(self.source.exists())

    def test_changed_admitted_input_is_retained_and_cannot_start_a_new_task(self):
        self.write_input()
        _, first = self.submit()
        original = Path(first["recordPath"]).read_bytes()
        self.request["completion"][0]["params"]["outputPath"] = "C:/work/different.docx"
        self.write_input()
        code, rejected = self.submit()
        self.assertEqual(4, code)
        self.assertEqual("TASK_INPUT_CONFLICT", rejected["stop"]["error"]["code"])
        self.assertTrue(self.source.is_file())
        self.assertEqual([], self.client.calls)
        self.assertEqual(original, Path(first["recordPath"]).read_bytes())

    def test_same_basename_in_different_task_directories_has_distinct_identity(self):
        self.write_input()
        _, first = self.submit()
        self.source = self.root / "next" / "input.json"
        self.source.parent.mkdir()
        self.write_input()
        _, second = self.submit()
        self.assertNotEqual(first["taskId"], second["taskId"])
        self.assertEqual(3, len(self.client.calls))
        self.assertEqual(task_store.input_identity(self.source),
                         task_store.input_identity(self.source.parent / ".." / "next" / "input.json"))

    def test_manual_identity_and_checks_in_every_section_are_rejected_before_admission(self):
        variants = [dict(self.request, version=v) for v in (1, 2, 3, True)]
        manual = copy.deepcopy(self.request)
        manual["taskId"] = "reused-id"
        variants.append(manual)
        for section in ("document", "steps", "completion"):
            request = copy.deepcopy(self.request)
            target = request[section] if section == "document" else request[section][0]
            target["checks"] = [{"actual": ref("task_save", "response", "data"),
                                 "op": "equals", "expected": "saved"}]
            variants.append(request)
        for request in variants:
            with self.subTest(request=request):
                self.write_input(request)
                code, result = self.submit()
                self.assertEqual(4, code)
                self.assertEqual("rejected", result["state"])
                self.assertIsNone(result["recordPath"])
                self.assertEqual([], self.client.calls)
                self.assertTrue(self.source.exists())

    def test_action_failure_and_uncertainty_stop_save_without_agent_checks(self):
        for kind in ("fail", "drop"):
            with self.subTest(kind=kind):
                self.source = self.root / (kind + ".json")
                self.write_input()
                code, result = self.submit(RecordingExecutor(**{kind: "writeContent"}))
                self.assertEqual(2, code, result)
                self.assertEqual("failed" if kind == "fail" else "unknown", result["outcome"])
                self.assertEqual("not_executed", result["completion"]["save"]["state"])
                self.assertEqual(["createDocument", "writeContent"], [n for n, _ in self.client.calls])
                self.assertFalse(self.source.exists())

    def test_cleanup_failure_does_not_erase_success_or_allow_replay(self):
        client = RecordingExecutor()
        def close():
            return {"outcome": "failed", "error": {"code": "CLEANUP_FAILED"}}
        client.close = close
        self.write_input()
        code, result = self.submit(client)
        self.assertEqual(4, code)
        self.assertEqual("succeeded", result["outcome"])
        self.assertEqual("failed", result["cleanup"]["outcome"])
        _, repeated = self.submit()
        self.assertEqual(result["taskId"], repeated["taskId"])
        self.assertEqual([], self.client.calls)

    def test_input_deletion_failure_is_reported_without_repeating_work(self):
        self.write_input()
        original = Path.unlink
        def unlink(path, *args, **kwargs):
            if path == self.source:
                raise PermissionError("input is locked")
            return original(path, *args, **kwargs)
        with patch.object(Path, "unlink", unlink):
            code, result = self.submit()
        self.assertEqual(0, code, result)
        self.assertEqual("cleanup_failed", result["taskFile"]["reason"])
        self.assertTrue(self.source.exists())
        _, repeated = self.submit()
        self.assertEqual([], self.client.calls)
        self.assertEqual(result["taskId"], repeated["taskId"])
        self.assertFalse(self.source.exists())

    def test_crash_between_request_and_initial_receipt_never_executes_on_resubmit(self):
        self.write_input()
        path = task_store.directory("word", task_store.input_identity(self.source))
        path.mkdir(parents=True)
        _write(path / "request.json", self.request)
        code, result = self.submit()
        self.assertEqual(4, code)
        self.assertEqual("interrupted", result["state"])
        self.assertEqual("TASK_OWNER_LOST", result["stop"]["error"]["code"])
        self.assertEqual([], self.client.calls)
        self.assertTrue(self.source.exists())

    def test_concurrent_submission_observes_live_owner_instead_of_dispatching(self):
        self.write_input()
        path = task_store.directory("word", task_store.input_identity(self.source))
        path.mkdir(parents=True)
        lock = task_store.acquire(path)
        try:
            _write(path / "request.json", self.request)
            task_store.publish(path, task_store.initial(self.request, path))
            code, result = self.submit()
            self.assertEqual(4, code)
            self.assertEqual("running", result["state"])
            self.assertEqual([], self.client.calls)
            self.assertFalse(self.source.exists())
        finally:
            lock.close()
        _, lost = self.cli("--task-status-file", str(self.source))
        self.assertEqual("interrupted", lost["state"])
        self.assertEqual("unknown", lost["cleanup"]["outcome"])

    def test_ranges_still_flow_from_prior_action_responses(self):
        self.request["document"] = action("openDocument", {"path": "C:/work/input.docx"})
        self.request["steps"] = [step("inspect", "inspectDocument", {
            "scope": {"kind": "document"},
            "limits": {"maxTextCharacters": 4096, "maxParagraphs": 128, "maxRuns": 512},
        }), step("replace", "replaceContent", {
            "target": {"kind": "range", "range": ref("inspect", "data", "paragraphs", 0, "range")},
            "replacement": {"kind": "blocks", "blocks": [{"kind": "paragraph", "runs": [{"text": "更新"}]}]},
        })]
        self.write_input()
        code, result = self.submit()
        self.assertEqual(0, code, result)
        self.assertEqual("revision-1", self.client.calls[2][1]["target"]["range"]["revision"])

    def test_cli_utf8_file_uses_task_resources_without_session_stack(self):
        self.request["completion"] = []
        self.write_input()
        events = self.root / "events.txt"
        run = subprocess.run([
            sys.executable, str(Path(__file__).with_name("cli_fixture.py")), str(events), "normal",
            "--app", "word", "--task-file", str(self.source),
        ], capture_output=True, env=dict(os.environ, PYTHONIOENCODING="ascii"), timeout=10)
        self.assertEqual(0, run.returncode, run.stdout + run.stderr)
        result = json.loads(run.stdout)
        self.assertEqual("中文、English 与🙂", result["steps"][0]["response"]["data"]["echo"]["blocks"][0]["runs"][0]["text"])
        self.assertNotIn("sessionId", result["steps"][0]["response"])
        self.assertNotIn("ready", result)
        self.assertTrue(events.with_suffix(".closed").exists())
        self.assertEqual("succeeded", result["cleanup"]["outcome"])
        self.assertFalse(self.source.exists())

    def test_formal_task_owner_loss_retains_effects_and_never_replays(self):
        self.request["completion"] = []
        for mode in ("owner_loss", "cleanup_loss"):
            with self.subTest(mode=mode):
                self.source = self.root / (mode + ".json")
                self.write_input()
                events = self.root / (mode + ".events")
                run = subprocess.run([
                    sys.executable, str(Path(__file__).with_name("cli_fixture.py")), str(events), mode,
                    "--app", "word", "--task-file", str(self.source),
                ], capture_output=True, timeout=10)
                self.assertEqual(9, run.returncode, run.stdout + run.stderr)
                self.assertFalse(self.source.exists())
                before = events.read_bytes()
                code, result = self.cli("--task-file", str(self.source))
                self.assertEqual(4, code)
                self.assertEqual("interrupted", result["state"])
                self.assertEqual("unknown", result["cleanup"]["outcome"])
                self.assertEqual("unknown" if mode == "owner_loss" else "succeeded", result["outcome"])
                self.assertEqual("unknown" if mode == "owner_loss" else "succeeded", result["steps"][0]["state"])
                self.assertEqual([], self.client.calls)
                self.assertEqual(before, events.read_bytes())

    def test_historical_receipts_are_query_only_and_missing_response_is_unknown(self):
        for version in (1, 2, 3):
            path = task_store.directory("word", "historical-" + str(version))
            path.mkdir(parents=True)
            historical = dict(self.request, version=version)
            _write(path / "request.json", historical)
            code, result = self.cli("--task-status", path.name)
            self.assertEqual(4, code)
            self.assertEqual("unknown", result["outcome"])
            self.assertEqual([], self.client.calls)
            receipt = {"type": "task.response", "version": version, "taskId": path.name,
                       "app": "word", "state": "completed", "outcome": "succeeded", "steps": []}
            _write(path / "response.json", receipt)
            code, result = self.cli("--task-status", path.name)
            self.assertEqual(0, code)
            self.assertEqual(receipt, result)
            self.assertEqual(historical, json.loads((path / "request.json").read_text()))

    def test_request_without_input_locator_is_rejected_not_assigned_fresh_identity(self):
        result = task_client.execute(self.request, "word")
        self.assertEqual("rejected", result["state"])
        self.assertIsNone(result["taskId"])


if __name__ == "__main__":
    unittest.main()
