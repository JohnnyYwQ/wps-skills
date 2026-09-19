import copy
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile

REPO = Path(__file__).resolve().parents[5]
HERE = REPO / "src/test/python/diagnostics/startup_communication"
RESOURCES = REPO / "src/test/resources/diagnostics/startup-communication"
sys.path.insert(0, str(HERE))
import child
import common
import debug
import worker
sys.path.insert(0, str(REPO / "src/main/python"))


class ExperimentTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.config = common.read_json(RESOURCES / "quick.json")

    def bundle(self):
        bundle = self.root / "bundle"
        bundle.mkdir()
        common.write_json(bundle / "experiment.json", self.config)
        common.write_json(bundle / "manifest.json", {"files": common.file_hashes(bundle)})

    def test_invalid_deadline_cannot_disable_watchdog(self):
        for value in (True, 0, -1, float("nan"), float("inf")):
            config = copy.deepcopy(self.config)
            config["limits"]["caseSeconds"] = value
            with self.subTest(value=value), self.assertRaises(ValueError):
                common.validate_config(config)

    def test_bundle_tamper_rejected(self):
        self.bundle()
        common.verify_bundle(self.root / "bundle")
        (self.root / "bundle/experiment.json").write_text("{}")
        with self.assertRaisesRegex(ValueError, "hash mismatch"):
            common.verify_bundle(self.root / "bundle")

    def test_first_unexpected_failure_stops_entire_run_and_never_replays(self):
        self.bundle()
        def failure(command, directory, timeout, journal):
            common.write_json(directory / "result.json", {"state": "failed", "error": "injected"})
            return 1, False
        with patch.object(worker, "environment", return_value={}), patch.object(worker, "supervise", side_effect=failure) as execute:
            self.assertEqual(worker.run(self.root, "quick"), 1)
            self.assertEqual(execute.call_count, 1)
            with self.assertRaises(FileExistsError):
                worker.run(self.root, "quick")
            self.assertEqual(execute.call_count, 1)
        status = common.read_json(self.root / "windows/status.json")
        self.assertEqual(status["state"], "failed")
        self.assertEqual(status["remaining"], "not_executed")
        self.assertEqual(status["cases"][0]["result"]["error"], "injected")

    def test_long_progress_is_bounded_while_complete_case_history_is_retained(self):
        self.config['cases']=['pipe']
        self.config['profiles']['quick']['maxRounds']=75
        self.bundle()
        def success(command,directory,timeout,journal):
            common.write_json(directory/'result.json',{'state':'passed'})
            return 0,False
        with patch.object(worker,'environment',return_value={}),patch.object(worker,'supervise',side_effect=success):
            self.assertEqual(worker.run(self.root,'quick'),0)
        status=common.read_json(self.root/'windows/status.json')
        self.assertLessEqual(len(status['cases']),50)
        self.assertEqual(status['observedPassed'],75)
        history=[json.loads(line) for line in (self.root/'windows/cases.jsonl').read_text().splitlines()]
        self.assertEqual(len(history),75)
        self.assertEqual(history[0]['name'],'r00001-pipe')
        self.assertEqual(history[-1]['name'],'r00075-pipe')

    def test_requested_stop_finishes_current_case_without_starting_next(self):
        self.bundle()
        def stop_after_case(command, directory, timeout, journal):
            common.write_json(directory / "result.json", {"state": "passed"})
            (self.root / "stop.request").write_text("stop")
            return 0, False
        with patch.object(worker, "environment", return_value={}), patch.object(worker, "supervise", side_effect=stop_after_case) as execute:
            self.assertEqual(worker.run(self.root, "quick"), 0)
            self.assertEqual(execute.call_count, 1)
        status = common.read_json(self.root / "windows/status.json")
        self.assertEqual(status["state"], "stopped")
        self.assertEqual(status["observedPassed"], 1)

    def test_watchdog_overrides_a_stale_success_result(self):
        self.bundle()
        def timeout_case(command, directory, timeout, journal):
            common.write_json(directory / "result.json", {"state": "passed"})
            return -9, True
        with patch.object(worker, "environment", return_value={}), patch.object(worker, "supervise", side_effect=timeout_case):
            worker.run(self.root, "quick")
        status = common.read_json(self.root / "windows/status.json")
        self.assertEqual(status["observedPassed"], 0)
        self.assertEqual(status["cases"][0]["result"]["state"], "unknown")

    def test_supervisor_bounds_a_hung_owned_child(self):
        code, expired = worker.supervise([sys.executable, "-c", "import time; time.sleep(30)"], self.root, 0.1,
                                         common.Journal(self.root / "events.jsonl"))
        self.assertTrue(expired)
        self.assertNotEqual(code, 0)

    def test_pending_visual_review_is_not_reported_as_completed_review(self):
        self.config["profiles"]["quick"]["maxRounds"] = 1
        self.bundle()
        def success(command, directory, timeout, journal):
            common.write_json(directory / "result.json", {"state": "passed", "visualReview": "pending"})
            return 0, False
        with patch.object(worker, "environment", return_value={}), patch.object(worker, "supervise", side_effect=success):
            self.assertEqual(worker.run(self.root, "quick"), 0)
        self.assertEqual(common.read_json(self.root / "windows/status.json")["visualReview"], "pending")

    def test_outer_wrapper_finishing_does_not_invalidate_supervisor_evidence(self):
        self.config["profiles"]["quick"]["maxRounds"] = 1
        self.bundle()
        evidence = self.root / "windows"
        evidence.mkdir()
        common.write_json(evidence / "launcher.json", {"state": "running"})
        def success(command, directory, timeout, journal):
            common.write_json(directory / "result.json", {"state": "passed"})
            return 0, False
        with patch.object(worker, "environment", return_value={}), patch.object(worker, "supervise", side_effect=success):
            worker.run(self.root, "quick")
        common.write_json(evidence / "launcher.json", {"state": "completed"})
        hashes = common.read_json(self.root / "evidence.sha256.json")
        self.assertNotIn("launcher.json", hashes)
        self.assertIn("status.json", hashes)
        for name, expected in hashes.items():
            self.assertEqual(common.digest(evidence / name), expected)

    def test_host_phase_logs_do_not_overwrite_each_other(self):
        debug.command([sys.executable, "-c", "print('one')"], self.root, "remote.status")
        debug.command([sys.executable, "-c", "print('two')"], self.root, "remote.status")
        outputs = list((self.root / "host-logs").glob("*.stdout.txt"))
        self.assertEqual(len(outputs), 2)
        self.assertEqual({p.read_text().strip() for p in outputs}, {"one", "two"})

    def test_remote_script_is_sent_as_stdin_not_an_unbounded_command_line(self):
        from unittest.mock import Mock
        script = "# 中文\n" * 5000 + "Write-Output 'ok'"
        with patch.object(debug, "command", return_value="ok") as command:
            self.assertEqual(debug.remote("win", script, self.root, "test"), "ok")
        args, kwargs = command.call_args
        self.assertLess(len(args[0][-1]), 1024)
        self.assertTrue(kwargs["input_text"].endswith(script))

    def test_evidence_failure_cannot_prevent_owned_probe_cleanup(self):
        from unittest.mock import Mock
        from contextlib import nullcontext
        launcher = Mock()
        launcher.close.return_value = []
        launcher._stderr_tails = {}
        journal = Mock()
        journal.event.side_effect = PermissionError("evidence unavailable")
        transport = Mock(side_effect=PermissionError("evidence unavailable"))
        with patch("wps_skills.windows.owned_process.WindowsOwnedProcessLauncher", return_value=launcher), \
                patch("wps_skills.windows.bridge_runtime.windows_powershell_executable", return_value="powershell"), \
                patch("wps_skills.core.timing.submission", return_value=nullcontext()):
            with self.assertRaises(PermissionError):
                child.pipe_case(self.root, self.root, self.config, journal, transport)
        launcher.close.assert_called_once()

    def test_package_has_provenance_and_does_not_modify_production_source(self):
        import bundle
        source = bundle.REPO / "src/main/resources/wps_skills/windows/bridge_loop.ps1"
        before = common.digest(source)
        destination = self.root / "bundle"
        bundle.build(destination, RESOURCES / "quick.json")
        self.assertEqual(common.digest(source), before)
        manifest = common.verify_bundle(destination)
        self.assertIn("host-source/debug.py", manifest["files"])
        info = common.read_json(destination / "instrumentation.json")
        self.assertEqual(info["changes"][0]["beforeSha256"], before)
        self.assertNotEqual(info["changes"][0]["afterSha256"], before)

    def test_independent_artifact_inspection_detects_false_success(self):
        path = self.root / "test.xlsx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="test" r:id="rId1"/></sheets></workbook>')
            archive.writestr("xl/_rels/workbook.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>')
            archive.writestr("xl/worksheets/sheet1.xml", '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row><c r="A1" t="inlineStr"><is><t>ok</t></is></c><c r="B1"><v>7</v></c></row></sheetData></worksheet>')
        self.assertEqual(child.verify_artifact(path, {"A1": "ok", "B1": 7})["cells"]["B1"], 7)
        with self.assertRaises(AssertionError):
            child.verify_artifact(path, {"A1": "wrong", "B1": 7})

    def test_snapshot_pointer_survives_run_directory_relocation(self):
        snapshot_id = 'snapshot-' + 'a' * 32
        destination = self.root / 'snapshots' / snapshot_id
        destination.mkdir(parents=True)
        self.assertEqual(destination, debug.snapshot_directory(self.root, {'path': '/old/run/snapshots/' + snapshot_id}))
        self.assertEqual(destination, debug.snapshot_directory(self.root, {'path': 'snapshots/' + snapshot_id}))
        with self.assertRaises(ValueError):
            debug.snapshot_directory(self.root, {'path': '/outside/other'})

    def test_windows_snapshot_separators_and_path_escape(self):
        archive = self.root / "snapshot.zip"
        target = self.root / "extract"
        with zipfile.ZipFile(archive, "w") as output:
            output.writestr("windows\\", "")
            output.writestr("windows\\status.json", "{}")
        debug.extract_snapshot(archive, target)
        self.assertTrue((target / "windows/status.json").is_file())
        with zipfile.ZipFile(archive, "w") as output:
            output.writestr("../escape", "bad")
        with self.assertRaises(ValueError):
            debug.extract_snapshot(archive, target)

    def test_submission_uncertainty_does_not_allow_second_start(self):
        common.write_json(self.root / "run.json", {"state": "trigger_pending"})
        from argparse import Namespace
        with self.assertRaisesRegex(ValueError, "already submitted"):
            debug.start(Namespace(run=self.root))


if __name__ == "__main__":
    unittest.main()
