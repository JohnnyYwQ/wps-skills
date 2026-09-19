import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from wps_skills.cli.build_word_skill import build_word_skill
from wps_skills.cli.build_skill import MAIN
from wps_skills.core.action_runtime import ActionAddress
from wps_skills.word.contracts import WORD_PRODUCTION_CONTRACT_SET
from wps_skills.word.actions.registry import WORD_ACTIONS


class WordSkillBuildTests(unittest.TestCase):
    def test_source_action_definitions_match_production_contracts(self):
        source = MAIN / "resources/skills/wps-word/references/actions.json"
        definitions = json.loads(source.read_text(encoding="utf-8"))
        resolved = WORD_PRODUCTION_CONTRACT_SET.batch_resolve([
            ActionAddress(app="word", action=name)
            for name in WORD_ACTIONS
        ])
        self.assertEqual("complete", resolved["status"])
        self.assertEqual(
            {item["address"]["action"]: item["contract"] for item in resolved["items"]},
            definitions,
        )

    def test_relocated_skill_discovers_and_resolves_without_repository_or_windows(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            built = build_word_skill(root / "first" / "wps-word")
            installed = root / "installed" / "wps-word"
            installed.parent.mkdir()
            built.rename(installed)
            environment = dict(os.environ, PYTHONPATH="", PYTHONNOUSERSITE="1")
            entry = installed / "scripts" / "word.py"
            help_result = subprocess.run([sys.executable, str(entry), "--help"], cwd=root,
                                         env=environment, capture_output=True, text=True, check=True)
            for legacy_option in ("--session", "--start", "--call", "--close", "--task-stdin",
                                  "--consume-task-file", "--debug-close-created-document"):
                self.assertNotIn(legacy_option, help_result.stdout)
            result = subprocess.run([sys.executable, str(entry), "--app", "word", "--index"], cwd=root, env=environment, capture_output=True, text=True, check=True)
            index = json.loads(result.stdout)
            self.assertEqual([entry.to_wire() for entry in WORD_PRODUCTION_CONTRACT_SET.action_index()], index["actions"])
            self.assertFalse((installed / "references/actions.json").exists())
            definitions = {name: action.contract.to_wire("word") for name, action in WORD_ACTIONS.items()}
            self.assertEqual(list(WORD_ACTIONS), list(definitions))
            result = subprocess.run([sys.executable, str(entry), "--app", "word", "--resolve", *definitions], cwd=root, env=environment, capture_output=True, text=True, check=True)
            resolved = json.loads(result.stdout)
            self.assertEqual("complete", resolved["status"])
            self.assertEqual(
                {item["address"]["action"]: item["contract"] for item in resolved["items"]},
                definitions,
            )
            # The Agent-facing package rejects legacy plans rather than letting
            # their manual IDs/checks bypass the new submission boundary.
            request = {"version": 2, "taskId": "relocated-v2", "app": "word",
                       "document": {"address": {"app": "word", "action": "createDocument"}, "params": {}}, "steps": [],
                       "completion": [{"address": {"app": "word", "action": "saveAs"},
                                       "params": {"outputPath": "relative.docx", "overwritePolicy": "failIfExists"}}]}
            result = subprocess.run(
                [sys.executable, str(entry), "--app", "word", "--task-stdin"],
                input=json.dumps(request), cwd=root,
                env=dict(environment, WPS_SKILLS_TASK_DIR=str(root / "receipts")),
                capture_output=True, text=True,
            )
            self.assertEqual(2, result.returncode, result.stdout + result.stderr)
            self.assertIn("error:", result.stderr)
            request.pop("version")
            request.pop("taskId")
            task_file = root / "临时请求🙂.json"
            task_file.write_text(json.dumps(request, ensure_ascii=False), encoding="utf-8")
            result = subprocess.run(
                [sys.executable, str(entry), "--app", "word", "--task-file", str(task_file)], cwd=root,
                env=dict(environment, WPS_SKILLS_TASK_DIR=str(root / "receipts")),
                capture_output=True, text=True,
            )
            self.assertEqual(4, result.returncode, result.stdout + result.stderr)
            rejection = json.loads(result.stdout)
            self.assertEqual("task_save", rejection["stop"]["stepId"])
            self.assertEqual("retained", rejection["taskFile"]["state"])
            self.assertTrue(task_file.exists())
            self.assertIsNone(rejection["recordPath"])
            if os.name != "nt":
                # Exercise the relocated production submission/receipt path up
                # to its explicit Windows startup boundary, without WPS.
                request["completion"] = []
                task_file.write_text(json.dumps(request), encoding="utf-8")
                result = subprocess.run(
                    [sys.executable, str(entry), "--app", "word", "--task-file", str(task_file)], cwd=root,
                    env=dict(environment, WPS_SKILLS_TASK_DIR=str(root / "receipts")), capture_output=True, text=True,
                )
                self.assertEqual(2, result.returncode, result.stdout + result.stderr)
                response = json.loads(result.stdout)
                self.assertEqual("startup", response["stop"]["phase"])
                self.assertEqual("removed", response.pop("taskFile")["state"])
                self.assertNotIn("version", response)
                query = subprocess.run(
                    [sys.executable, str(entry), "--app", "word", "--task-status-file", str(task_file)], cwd=root,
                    env=dict(environment, WPS_SKILLS_TASK_DIR=str(root / "receipts")), capture_output=True, text=True,
                )
                self.assertEqual(response, json.loads(query.stdout))
                self.assertFalse(task_file.exists())
            # The packaged Runtime must also resolve the actual PowerShell resources.
            code = "import sys; sys.path.insert(0, sys.argv[1]); import word; import importlib, pkgutil, wps_skills; [importlib.import_module(m.name) for m in pkgutil.walk_packages(wps_skills.__path__, 'wps_skills.')]; from wps_skills.word.windows.task_factory import BRIDGE_SCRIPT; assert BRIDGE_SCRIPT.is_file(); assert BRIDGE_SCRIPT.with_name('word_actions.ps1').is_file(); assert BRIDGE_SCRIPT.with_name('word_story_xml.ps1').is_file(); assert not hasattr(word, 'open_session'); assert 'wps_skills.core.action_session' not in sys.modules"
            subprocess.run([sys.executable, "-c", code, str(entry.parent)], cwd=root, env=environment, check=True)
            runtime = installed / "runtime/src/main/python/wps_skills"
            for retired in ("core/action_session.py", "client/session_client.py", "client/managed_session.py",
                            "client/legacy_task_client.py", "client/legacy_task_request.py",
                            "cli/call.py", "cli/legacy_call.py", "host"):
                self.assertFalse((runtime / retired).exists(), retired)
            manifest = json.loads((installed / "runtime/files.sha256.json").read_text())
            for name, digest in manifest.items():
                self.assertEqual(digest, hashlib.sha256((installed / name).read_bytes()).hexdigest())
            self.assertTrue((installed / "SKILL.md").is_file())
            self.assertFalse((installed / "runtime/src/test").exists())
            self.assertFalse((installed / "runtime/src/main/resources/skills").exists())

            # Historical receipts remain readable without shipping the legacy
            # execution stack; a missing receipt must never trigger replay.
            history = root / "receipts/word/historical"
            history.mkdir(parents=True)
            (history / "request.json").write_text(json.dumps({
                "app": "word", "version": 2, "taskId": "historical", "steps": [],
            }))
            query = subprocess.run(
                [sys.executable, str(entry), "--app", "word", "--task-status", "historical"],
                cwd=root, env=dict(environment, WPS_SKILLS_TASK_DIR=str(root / "receipts")),
                capture_output=True, text=True,
            )
            self.assertEqual(4, query.returncode, query.stderr)
            self.assertEqual("HISTORICAL_TASK_RECEIPT_UNAVAILABLE",
                             json.loads(query.stdout)["stop"]["error"]["code"])

    def test_existing_destination_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as temporary:
            destination = Path(temporary) / "wps-word"
            destination.mkdir()
            marker = destination / "user.txt"
            marker.write_text("keep")
            with self.assertRaises(FileExistsError):
                build_word_skill(destination)
            self.assertEqual("keep", marker.read_text())


if __name__ == "__main__":
    unittest.main()
