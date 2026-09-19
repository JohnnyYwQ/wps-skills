import copy
import unittest

from wps_skills.core.action_runtime import ContractValidationError, ControllerContext
from wps_skills.windows.bridge_types import BackendActionFailure
from wps_skills.word.persistence.output_names import MAX_OUTPUT_ATTEMPTS, output_candidate
from wps_skills.word.actions.registry import WORD_ACTIONS, WORD_PRODUCTION_CONTRACT_SET


def conflict(code="OUTPUT_ALREADY_EXISTS", outcome="failed", binding="unchanged"):
    return BackendActionFailure(outcome=outcome, code=code, message=code, binding_disposition=binding)


class OutputBackend:
    def __init__(self, failures=()):
        self.failures = list(failures)
        self.calls = []

    def invoke(self, document, operation, context):
        self.calls.append((document, operation, context))
        if self.failures:
            raise self.failures.pop(0)
        state = {"persistenceState": "saved", "readOnly": False}
        result = {
            "revisionBefore": "revision-1", "revisionAfter": "revision-1",
            "artifact": {"path": operation.arguments["outputPath"],
                         "format": "docx" if operation.name == "save_as_artifact" else "pdf", "sizeBytes": 100},
            "replacedExisting": False,
        }
        if operation.name == "save_as_artifact":
            result["documentState"] = state
        else:
            result.update(documentStateBefore=state, documentStateAfter=dict(state))
        return result


class OutputRetryTests(unittest.TestCase):
    def setUp(self):
        self.document = object()
        self.context = ControllerContext(request_id="r1", trace_id="t1", deadline_at=100)

    def run_action(self, name, backend, policy="renameIfExists"):
        params = {"outputPath": "C:/work/报告." + ("docx" if name == "saveAs" else "pdf"),
                  "overwritePolicy": policy}
        WORD_PRODUCTION_CONTRACT_SET.validate_params(name, params)
        result = WORD_ACTIONS[name].handler(backend, self.document, params, self.context)
        if result.outcome == "succeeded":
            WORD_PRODUCTION_CONTRACT_SET.validate_result(name, result.data, params=params)
        return result, params

    def test_name_conflicts_retry_inside_one_action_on_same_document_and_deadline(self):
        for name in ("saveAs", "exportPdf"):
            with self.subTest(action=name):
                backend = OutputBackend([conflict(), conflict("OUTPUT_IN_USE")])
                result, params = self.run_action(name, backend)
                self.assertEqual("succeeded", result.outcome)
                self.assertEqual([output_candidate(params["outputPath"], n) for n in range(3)],
                                 [call[1].arguments["outputPath"] for call in backend.calls])
                self.assertTrue(all(call[0] is self.document and call[2] is self.context for call in backend.calls))
                self.assertTrue(all(call[1].arguments["overwritePolicy"] == "failIfExists" for call in backend.calls))
                self.assertEqual({"requestedPath": params["outputPath"], "attempts": 3, "renamed": True},
                                 dict(result.data["outputResolution"]))
                self.assertEqual("renameIfExists", params["overwritePolicy"])

    def test_available_original_path_saves_once_and_reports_no_rename(self):
        result, params = self.run_action("saveAs", OutputBackend())
        self.assertEqual(params["outputPath"], result.data["artifact"]["path"])
        self.assertEqual(1, result.data["outputResolution"]["attempts"])
        self.assertFalse(result.data["outputResolution"]["renamed"])

    def test_strict_and_replace_policies_do_not_enable_retry(self):
        for name, policy in (("saveAs", "failIfExists"), ("exportPdf", "failIfExists"),
                             ("exportPdf", "replaceExisting")):
            backend = OutputBackend([conflict()])
            result, _ = self.run_action(name, backend, policy)
            self.assertEqual("failed", result.outcome)
            self.assertEqual(1, len(backend.calls))

    def test_unknown_binding_loss_permission_and_other_errors_never_retry(self):
        failures = [conflict(outcome="unknown"), conflict(binding="unprovable"),
                    conflict(binding="lost"), conflict("OUTPUT_ACCESS_DENIED"),
                    conflict("OUTPUT_MATCHES_BOUND_DOCUMENT"), conflict("OUTPUT_PARENT_NOT_FOUND"),
                    conflict("OUTPUT_WRITE_FAILED"), conflict("DOCUMENT_LEASE_CONFLICT"),
                    conflict("DOCUMENT_QUARANTINED")]
        for name in ("saveAs", "exportPdf"):
            for failure in failures:
                with self.subTest(action=name, error=failure.code, outcome=failure.outcome):
                    backend = OutputBackend([conflict(), failure])
                    result, _ = self.run_action(name, backend)
                    self.assertEqual(failure.outcome, result.outcome)
                    self.assertEqual(failure.code, result.error.code)
                    self.assertEqual(2, len(backend.calls))

    def test_exhaustion_is_bounded_and_leaves_binding_usable(self):
        backend = OutputBackend([conflict() for _ in range(MAX_OUTPUT_ATTEMPTS)])
        result, _ = self.run_action("saveAs", backend)
        self.assertEqual(MAX_OUTPUT_ATTEMPTS, len(backend.calls))
        self.assertEqual("OUTPUT_NAME_EXHAUSTED", result.error.code)
        self.assertEqual("unchanged", result.binding_disposition)
        self.assertEqual("usable", result.controller_state)

    def test_candidate_keeps_directory_unicode_extension_and_separator_style(self):
        for path, expected in [
            (r"C:\目录\报告.docx", r"C:\目录\报告 (1).docx"),
            (r"\\server\share\报告.docx", r"\\server\share\报告 (1).docx"),
            ("/work/报告.pdf", "/work/报告 (1).pdf"),
        ]:
            self.assertEqual(expected, output_candidate(path, 1))

    def test_contract_rejects_other_directories_overwrite_and_unproved_renames(self):
        for name in ("saveAs", "exportPdf"):
            result, params = self.run_action(name, OutputBackend([conflict()]))
            # ControllerResult freezes nested mappings; obtain plain wire data.
            def plain(value):
                if hasattr(value, "items"):
                    return {k: plain(v) for k, v in value.items()}
                return value
            original = plain(result.data)
            variants = []
            for path in ("C:/other/报告 (1)." + original["artifact"]["format"], params["outputPath"]):
                data = copy.deepcopy(original); data["artifact"]["path"] = path; variants.append(data)
            data = copy.deepcopy(original); data["replacedExisting"] = True; variants.append(data)
            data = copy.deepcopy(original); data.pop("outputResolution"); variants.append(data)
            data = copy.deepcopy(original); data["outputResolution"]["attempts"] = MAX_OUTPUT_ATTEMPTS + 1; variants.append(data)
            data = copy.deepcopy(original); data["outputResolution"]["renamed"] = False; variants.append(data)
            for data in variants:
                with self.assertRaises(ContractValidationError):
                    WORD_PRODUCTION_CONTRACT_SET.validate_result(name, data, params=params)


if __name__ == "__main__":
    unittest.main()
