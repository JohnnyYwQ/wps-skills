import json
from pathlib import Path
import sys
import tempfile
import unittest

REPO = Path(__file__).resolve().parents[5]
HERE = REPO / "src/test/python/diagnostics/startup_communication"
RESOURCES = REPO / "src/test/resources/diagnostics/startup-communication"
sys.path.insert(0, str(HERE))
from common import read_json, validate_config, write_json
from com_report import summarize


class ComReportTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        (self.root / "com").mkdir()

    def events(self, records):
        (self.root / "com/com-1.jsonl").write_text("\n".join(json.dumps(r) for r in records), encoding="utf-8")

    def event(self, stage, state="succeeded", **kwargs):
        return dict(stage=stage, event="stage.finished", outcome=state, **kwargs)

    def success_artifacts(self):
        write_json(self.root / "response.json", {"steps": [{"id": "write", "state": "succeeded"}, {"id": "after", "state": "succeeded"}], "stop": None})
        write_json(self.root / "artifact-verification.json", {"cells": {"A1": "hello", "B1": 1}})

    def test_existing_connection_does_not_claim_activation_or_cold_start(self):
        self.events([self.event(n) for n in ("environment", "com.registration", "com.attach", "document.bind_create")])
        self.success_artifacts()
        value = summarize(self.root)
        self.assertTrue(value["diagnosticsComplete"])
        self.assertEqual(value["connection"], "existing")
        self.assertEqual(value["stages"]["com.activate"]["state"], "not_executed")
        self.assertFalse(value["coverage"]["activation"])
        self.assertFalse(value["coverage"]["coldStartProven"])

    def test_successful_fallback_keeps_original_attach_error(self):
        error = {"hresult": "0x800401E3", "type": "COMException", "message": "unavailable"}
        self.events([self.event("environment"), self.event("com.registration"),
                     self.event("com.attach", "not_found", errors=[error]), self.event("com.activate"), self.event("document.bind_create")])
        self.success_artifacts()
        value = summarize(self.root)
        self.assertTrue(value["diagnosticsComplete"])
        self.assertIsNone(value["failureStage"])
        self.assertEqual(value["connection"], "activation")
        self.assertEqual(value["stages"]["com.attach"]["detail"]["errors"], [error])

    def test_activation_failure_and_inflight_binding_are_distinguishable(self):
        self.events([self.event("environment"), self.event("com.registration"),
                     self.event("com.attach", "failed"), self.event("com.activate", "failed"),
                     {"stage": "document.bind_create", "event": "stage.started"}])
        value = summarize(self.root)
        self.assertEqual(value["failureStage"], "com.activate")
        self.assertEqual(value["stages"]["document.bind_create"]["state"], "unknown")
        self.assertFalse(value["diagnosticsComplete"])

    def test_missing_registration_stops_before_connection(self):
        self.events([self.event("environment"), self.event("com.registration", "failed", data={"reason": "registration_unavailable"})])
        value = summarize(self.root)
        self.assertEqual(value["failureStage"], "com.registration")
        self.assertEqual(value["stages"]["com.attach"]["state"], "not_executed")

    def test_truncated_diagnostic_record_cannot_be_a_complete_pass(self):
        self.events([self.event(n) for n in ("environment", "com.registration", "com.attach", "document.bind_create")])
        with (self.root / "com/com-1.jsonl").open("a") as stream:
            stream.write('\n{"stage":')
        self.success_artifacts()
        value = summarize(self.root)
        self.assertFalse(value["diagnosticsComplete"])
        self.assertEqual(len(value["malformedRecords"]), 1)

    def test_failed_independent_artifact_check_is_not_unexecuted(self):
        (self.root / "events.jsonl").write_text(json.dumps({"name": "wps.artifact_verify", "event": "phase.finished", "outcome": "failed"}))
        value = summarize(self.root)
        self.assertEqual(value["stages"]["artifact.verify"]["state"], "failed")

    def test_open_case_requires_earlier_fixture_and_uses_its_own_binding(self):
        config = read_json(RESOURCES / "com.json")
        validate_config(config)
        config["cases"] = ["wps-open", "wps"]
        with self.assertRaisesRegex(ValueError, "earlier"):
            validate_config(config)
        write_json(self.root / "input.json", {"document": {"address": {"action": "openWorkbook"}}})
        self.events([self.event(n) for n in ("environment", "com.registration", "com.attach", "document.bind_open")])
        self.success_artifacts()
        self.assertTrue(summarize(self.root)["diagnosticsComplete"])


if __name__ == "__main__":
    unittest.main()
