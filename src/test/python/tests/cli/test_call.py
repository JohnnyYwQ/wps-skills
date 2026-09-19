import io
import json
import unittest
from unittest.mock import patch

from wps_skills.cli import call
from wps_skills.core.trace_journal import JsonlTraceJournal


class CallDispatchTests(unittest.TestCase):
    def test_discovery_is_side_effect_free_and_only_exposes_production_word(self):
        output = io.StringIO()
        with patch("wps_skills.word.windows.task_factory.build_task", side_effect=AssertionError("must not start a Task")), patch.object(JsonlTraceJournal, "default", side_effect=AssertionError("must not create traces")):
            self.assertEqual(0, call.main(["--app", "word", "--index"], output_stream=output))
        index = json.loads(output.getvalue())
        self.assertEqual("word", index["app"])
        self.assertEqual(14, len(index["actions"]))
        self.assertIn("saveAs", [item["action"] for item in index["actions"]])

    def test_batch_resolution_reports_unavailable_actions_without_hiding_resolved_items(self):
        output = io.StringIO()
        code = call.main(["--app", "word", "--resolve", "writeContent", "unsupportedAction"], output_stream=output)
        result = json.loads(output.getvalue())
        self.assertEqual(2, code)
        self.assertEqual("partial", result["status"])
        self.assertEqual("resolved", result["items"][0]["status"])
        self.assertEqual("UNKNOWN_ACTION", result["items"][1]["error"]["code"])
        self.assertIn("parameters", result["items"][0]["contract"])

    def test_word_old_entrances_reject_before_resource_creation(self):
        variants = (["--session"], ["--start"], ["--call", "handle"], ["--status", "handle"],
                    ["--close", "handle"], ["--task-stdin"],
                    ["--task-file", "input.json", "--consume-task-file"],
                    ["--session", "--debug-close-created-document"])
        for args in variants:
            with self.subTest(args=args), patch("wps_skills.word.windows.task_factory.build_task", side_effect=AssertionError("no startup")):
                with self.assertRaises(SystemExit) as error:
                    call.main(["--app", "word", *args])
                self.assertEqual(2, error.exception.code)


if __name__ == "__main__":
    unittest.main()
