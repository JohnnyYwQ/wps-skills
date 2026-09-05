import io
import json
import unittest
from unittest.mock import patch

from wps_skills.cli import call


class CallSessionHostTests(unittest.TestCase):
    def test_discovery_is_side_effect_free_and_only_exposes_production_word(self):
        output = io.StringIO()
        with patch.object(call, "_production_session_factory", side_effect=AssertionError("must not start a Session")), patch.object(call.JsonlTraceJournal, "default", side_effect=AssertionError("must not create traces")):
            self.assertEqual(0, call.main(["--app", "word", "--index"], output_stream=output))
        index = json.loads(output.getvalue())
        self.assertEqual("word", index["app"])
        self.assertEqual(13, len(index["actions"]))
        self.assertNotIn("saveAs", [item["action"] for item in index["actions"]])

    def test_batch_resolution_reports_unavailable_actions_without_hiding_resolved_items(self):
        output = io.StringIO()
        code = call.main(["--app", "word", "--resolve", "writeContent", "saveAs"], output_stream=output)
        result = json.loads(output.getvalue())
        self.assertEqual(2, code)
        self.assertEqual("partial", result["status"])
        self.assertEqual("resolved", result["items"][0]["status"])
        self.assertEqual("UNKNOWN_ACTION", result["items"][1]["error"]["code"])
        self.assertIn("parameters", result["items"][0]["contract"])

    def test_unavailable_application_discovery_has_no_word_fallback(self):
        output, errors = io.StringIO(), io.StringIO()
        with patch("wps_skills.ppt.contracts.PPT_PRODUCTION_CONTRACT_SET", None):
            code = call.main(["--app", "ppt", "--index"], output_stream=output, error_stream=errors)
        self.assertEqual(4, code)
        self.assertEqual("", output.getvalue())
        self.assertIn("WPS_DISCOVERY_UNAVAILABLE", errors.getvalue())

    def test_excel_discovery_and_resolution_use_its_production_set_without_launching(self):
        output = io.StringIO()
        with patch.object(call, "_production_session_factory", side_effect=AssertionError("discovery must not launch")):
            code = call.main(["--app", "excel", "--index"], output_stream=output)
        self.assertEqual(0, code)
        index = json.loads(output.getvalue())
        self.assertEqual("excel", index["app"])
        self.assertEqual(31, len(index["actions"]))
        self.assertNotIn("writeContent", [entry["action"] for entry in index["actions"]])
        output = io.StringIO()
        code = call.main(["--app", "excel", "--resolve", "readRange", "createWorkbook"], output_stream=output)
        self.assertEqual(2, code)
        self.assertEqual("partial", json.loads(output.getvalue())["status"])

    def test_excel_assembly_has_no_word_fallback(self):
        sentinel = object()
        with patch("wps_skills.excel.windows.session.build_session", return_value=sentinel) as build:
            self.assertIs(sentinel, call._production_session_factory(application="excel", session_id="excel-1"))
        build.assert_called_once_with(session_id="excel-1")

    def test_ppt_discovery_and_assembly_are_application_scoped(self):
        output = io.StringIO()
        with patch.object(call, "_production_session_factory", side_effect=AssertionError("must not launch")):
            self.assertEqual(0, call.main(["--app", "ppt", "--index"], output_stream=output))
        index = json.loads(output.getvalue())
        self.assertEqual("ppt", index["app"])
        self.assertEqual(33, len(index["actions"]))
        self.assertNotIn("createPresentation", [entry["action"] for entry in index["actions"]])
        sentinel = object()
        with patch("wps_skills.ppt.windows.session.build_session", return_value=sentinel) as build:
            self.assertIs(sentinel, call._production_session_factory(application="ppt", session_id="ppt-1"))
        build.assert_called_once_with(session_id="ppt-1")

    def test_apps_without_a_production_slice_fail_before_protocol_output(self):
        for application in ("ppt",):
            with self.subTest(application=application):
                output = io.StringIO()
                errors = io.StringIO()

                with patch("wps_skills.ppt.contracts.PPT_PRODUCTION_CONTRACT_SET", None):
                    exit_code = call.main(
                        ["--session", "--app", application],
                        input_stream=io.StringIO(""),
                        output_stream=output,
                        error_stream=errors,
                    )

                self.assertEqual(4, exit_code)
                self.assertEqual("", output.getvalue())
                self.assertIn("WPS_SESSION_STARTUP_FAILED", errors.getvalue())
                self.assertIn(f"app={application}", errors.getvalue())

    def test_debug_cleanup_is_an_explicit_word_session_control(self):
        parsed = call._parser().parse_args([
            "--session",
            "--app",
            "word",
            "--debug-close-created-document",
        ])
        self.assertTrue(parsed.debug_close_created_document)

        sentinel = object()
        with patch(
            "wps_skills.word.windows.session.build_session",
            return_value=sentinel,
        ) as build:
            result = call._production_session_factory(
                application="word",
                session_id="session-1",
                debug_close_created_document=True,
            )

        self.assertIs(sentinel, result)
        build.assert_called_once_with(
            session_id="session-1",
            debug_close_created_document=True,
        )


if __name__ == "__main__":
    unittest.main()
