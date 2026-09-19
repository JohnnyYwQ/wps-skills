"""Observable Word Task lifecycle, using real request/contract validation."""

import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from wps_skills.client import task_client, task_request, task_store


def step(identifier, action, params, checks=None):
    result = {"id": identifier, "address": {"app": "word", "action": action}, "params": params}
    if checks is not None:
        result["checks"] = checks
    return result


def action(name, params=None, checks=None):
    value = {"address": {"app": "word", "action": name}, "params": {} if params is None else params}
    if checks is not None:
        value["checks"] = checks
    return value


def pdf(path="C:/work/output.pdf"):
    return action("exportPdf", {"outputPath": path, "overwritePolicy": "failIfExists"})


def ref(identifier, *path):
    return {"$ref": {"step": identifier, "path": list(path)}}


class RecordingExecutor:
    def __init__(self, *, modified=False, fail=None, drop=None):
        self.modified, self.fail, self.drop = modified, fail, drop
        self.calls = []
        self.closed = False
        self.can_execute = True

    def execute(self, address, params):
        name = address["action"]
        self.calls.append((name, params))
        if name == self.drop:
            raise RuntimeError("native response lost")
        failed = name == self.fail
        self.last_response = {
            "outcome": "failed" if failed else "succeeded",
            "data": {"documentState": {"persistenceState": "modified" if self.modified else "saved"},
                     "text": "项目周报 😀", "truncated": False,
                     "paragraphs": [{"range": {"start": 0, "end": 8, "revision": "revision-1"}}]},
            "error": {"code": "TEST_FAILURE", "message": "failed"} if failed else None,
        }
        return self.last_response

    def close(self):
        self.closed = True
        return {"outcome": "succeeded", "error": None}


class WordTaskTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        env = patch.dict(os.environ, {"WPS_SKILLS_TASK_DIR": str(self.root / "tasks")})
        env.start()
        self.addCleanup(env.stop)
        self.source = self.root / "input.json"
        self.request = {
            "app": "word",
            "document": action("openDocument", {"path": "C:/work/input.docx"}),
            "steps": [step("inspect", "inspectDocument", {
                "scope": {"kind": "document"},
                "limits": {"maxTextCharacters": 4096, "maxParagraphs": 128, "maxRuns": 512},
            })],
            "completion": [action("save")],
        }

    def execute(self, client=None):
        self.client = client or RecordingExecutor()
        return task_client.execute(self.request, "word", executor_factory=lambda **kwargs: self.client, source_path=self.source)

    def names(self):
        return [name for name, params in self.client.calls]

    def test_existing_document_is_acquired_once_and_save_is_always_last(self):
        result = self.execute()
        self.assertEqual("completed", result["state"], result)
        self.assertEqual(["openDocument", "inspectDocument", "save"], self.names())
        self.assertEqual(["inspect"], [s["id"] for s in result["steps"]])
        self.assertEqual("succeeded", result["document"]["state"])
        self.assertEqual("succeeded", result["completion"]["save"]["state"])
        self.assertIsNone(result["completion"]["pdf"])
        self.assertEqual(result, task_store.status_file("word", self.source))
        with patch("wps_skills.word.windows.task_factory.build_task", side_effect=AssertionError("no replay")):
            self.assertEqual(result, task_client.execute(self.request, "word", source_path=self.source))

    def test_new_document_save_as_then_pdf_and_no_content_steps(self):
        self.request.update(document=action("createDocument"), steps=[], completion=[
            pdf(), action("saveAs", {"outputPath": "C:/work/output.docx", "overwritePolicy": "failIfExists"}),
        ])
        result = self.execute()
        self.assertEqual("succeeded", result["outcome"], result)
        self.assertEqual(["createDocument", "saveAs", "exportPdf"], self.names())
        self.assertEqual("failIfExists", self.client.calls[1][1]["overwritePolicy"])
        self.assertEqual("succeeded", result["completion"]["pdf"]["state"])

    def test_read_only_allows_existing_changes_and_never_saves(self):
        self.request["completion"] = []
        result = self.execute(RecordingExecutor(modified=True))
        self.assertEqual("succeeded", result["outcome"])
        self.assertEqual(["openDocument", "inspectDocument"], self.names())

    def test_pdf_only_does_not_save_word_even_with_existing_changes(self):
        self.request["completion"] = [pdf()]
        result = self.execute(RecordingExecutor(modified=True))
        self.assertEqual("succeeded", result["outcome"])
        self.assertIsNone(result["completion"]["save"])
        self.assertEqual(["openDocument", "inspectDocument", "exportPdf"], self.names())

    def test_pdf_failure_preserves_successful_word_save(self):
        self.request["completion"].append(pdf())
        result = self.execute(RecordingExecutor(fail="exportPdf"))
        self.assertEqual("failed", result["outcome"])
        self.assertEqual("succeeded", result["completion"]["save"]["state"])
        self.assertEqual("failed", result["completion"]["pdf"]["state"])
        self.assertEqual("task_pdf", result["stop"]["stepId"])

    def test_existing_changes_stop_before_content_and_save_unless_authorized(self):
        result = self.execute(RecordingExecutor(modified=True))
        self.assertEqual("TASK_EXISTING_CHANGES_CONFIRMATION_REQUIRED", result["stop"]["error"]["code"])
        self.assertEqual(["openDocument"], self.names())
        self.assertEqual("not_executed", result["completion"]["save"]["state"])
        self.source = self.root / "authorized.json"
        self.request["includeExistingChanges"] = True
        self.assertEqual("succeeded", self.execute(RecordingExecutor(modified=True))["outcome"])
        self.assertEqual(["openDocument", "inspectDocument", "save"], self.names())




    def test_failed_acquisition_or_content_never_runs_completion(self):
        for action, expected in (("openDocument", ["openDocument"]),
                                 ("inspectDocument", ["openDocument", "inspectDocument"])):
            self.source = self.root / (action + ".json")
            result = self.execute(RecordingExecutor(fail=action))
            self.assertEqual("failed", result["outcome"])
            self.assertEqual(expected, self.names())
            self.assertEqual("not_executed", result["completion"]["save"]["state"])

    def test_failed_or_uncertain_save_prevents_pdf_and_is_never_replayed(self):
        self.request["completion"].append(pdf())
        for client, outcome in ((RecordingExecutor(fail="save"), "failed"),
                                (RecordingExecutor(drop="save"), "unknown")):
            self.source = self.root / (outcome + ".json")
            result = self.execute(client)
            self.assertEqual(outcome, result["completion"]["save"]["state"])
            self.assertEqual("not_executed", result["completion"]["pdf"]["state"])
            self.assertEqual(["openDocument", "inspectDocument", "save"], self.names())
            self.assertEqual(result, self.execute(client))
            self.assertEqual(3, len(client.calls))

    def test_dynamic_ranges_still_resolve_inside_content_steps(self):
        self.request["steps"].append(step("replace", "replaceContent", {
            "target": {"kind": "range", "range": ref("inspect", "data", "paragraphs", 0, "range")},
            "replacement": {"kind": "blocks", "blocks": [
                {"kind": "paragraph", "runs": [{"text": "更新"}]},
            ]},
        }))
        self.assertEqual("succeeded", self.execute()["outcome"])
        self.assertEqual({"start": 0, "end": 8, "revision": "revision-1"},
                         self.client.calls[2][1]["target"]["range"])

    def test_invalid_lifecycle_and_parameters_reject_before_start(self):
        variants = [
            {"document": action("createDocument")},  # save has no path on new documents
            {"document": {"kind": "new"}},  # obsolete alias envelope is rejected
            {"completion": {"mode": "save"}},
            {"completion": [action("saveAs")]},
            {"completion": [action("save", {"outputPath": "C:/extra.docx"})]},
            {"includeExistingChanges": "yes"},
            {"completion": [pdf("relative.pdf")]},
            {"document": action("openDocument", {"path": "relative.docx"})},
            {"document": action("save")},
            {"completion": [action("createDocument")]},
            {"completion": [action("save"), action("saveAs", {"outputPath": "C:/a.docx", "overwritePolicy": "failIfExists"})]},
            {"completion": [pdf(), pdf()]},
            {"completion": [action("exportPdf", {"outputPath": "C:/a.pdf"})]},
            {"completion": [{"address": {"app": "excel", "action": "save"}, "params": {}}]},
            {"steps": [step("save", "save", {})]},
            {"steps": [step("task_document", "inspectDocument", {})]},
            {"steps": [step("bad", "writeContent", {})]},
            {"steps": [step("bad", "notRegistered", {})]},
            {"completion": [action("saveAs", {"outputPath": ref("inspect", "data", "text"), "overwritePolicy": "failIfExists"})]},
        ]
        for i, changes in enumerate(variants):
            request = dict(self.request, **changes)
            with self.subTest(changes=changes), patch("wps_skills.word.windows.task_factory.build_task", side_effect=AssertionError("no startup")):
                self.assertEqual("rejected", task_client.execute(request, "word", source_path=self.root / (str(i) + ".json"))["state"])

    def test_known_invalid_fields_beside_references_reject_before_acquisition(self):
        self.request["steps"].append(step("replace", "replaceContent", {
            "target": {"kind": "range", "range": ref("inspect", "data", "paragraphs", 0, "range")},
            "replacement": {"kind": "blocks", "blocks": [{"kind": "paragraph", "runs": [{"text": "new"}]}]},
            "unexpected": True,
        }))
        result = self.execute()
        self.assertEqual("rejected", result["state"])
        self.assertEqual("replace", result["stop"]["stepId"])
        self.assertEqual([], self.names())

    def test_interruption_keeps_prior_effects_and_releases_resources(self):
        client = RecordingExecutor()
        original = client.execute
        def interrupted(address, params):
            if address["action"] == "inspectDocument":
                raise KeyboardInterrupt()
            return original(address, params)
        client.execute = interrupted
        result = self.execute(client)
        self.assertEqual("TASK_INTERRUPTED", result["stop"]["error"]["code"])
        self.assertEqual("succeeded", result["document"]["state"])
        self.assertEqual("unknown", result["steps"][0]["state"])
        self.assertEqual("not_executed", result["completion"]["save"]["state"])
        self.assertTrue(client.closed)

    def test_owner_loss_marks_lifecycle_in_flight_unknown(self):
        from wps_skills.word.task.plan import compile_request
        compile_request(self.request)
        path = task_store.directory("word", task_store.input_identity(self.source))
        path.mkdir(parents=True)
        result = task_store.initial(self.request, path)
        result["completion"]["save"]["state"] = "running"
        task_store.publish(path, result)
        snapshot = task_store.snapshot(path, owner_alive=False)
        self.assertEqual("unknown", snapshot["completion"]["save"]["state"])
        self.assertEqual("task_save", snapshot["stop"]["stepId"])



if __name__ == "__main__":
    unittest.main()
