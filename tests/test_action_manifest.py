"""Behavior tests for the WPS Action baseline validator CLI."""

from __future__ import annotations

from contextlib import contextmanager
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import unittest


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
VALIDATOR = REPOSITORY_ROOT / "scripts" / "validate_action_manifest.py"
REPRESENTATIVE_FIXTURES = (
    REPOSITORY_ROOT / "tests" / "fixtures" / "representative-actions.json"
)
UNIFIED_ADDIN = REPOSITORY_ROOT / "skills/wps-office/assets/wps-addin/main.js"
EXCEL_REFERENCE = REPOSITORY_ROOT / "skills/wps-office/references/excel.md"


class ActionManifestValidatorTests(unittest.TestCase):
    def test_repository_examples_cover_each_application_and_risk_contract(self) -> None:
        manifest = json.loads(
            (REPOSITORY_ROOT / "skills/wps-office/references/action-manifest.json").read_text(
                encoding="utf-8"
            )
        )
        fixtures = json.loads(REPRESENTATIVE_FIXTURES.read_text(encoding="utf-8"))
        actions = {entry["action"]: entry for entry in manifest["actions"]}

        self.assertEqual(
            {fixture["application"] for fixture in fixtures},
            {"common", "excel", "word", "powerpoint"},
        )
        for fixture in fixtures:
            action = actions[fixture["action"]]
            self.assertEqual(action["application"], fixture["application"])
            self.assertEqual(action["risk"], fixture["risk"])
            self.assertCountEqual(action["parameters"]["required"], fixture["required"])
            self.assertIn(
                {
                    "name": action["examples"][0]["name"],
                    "params": fixture["params"],
                    "result": fixture["result"],
                },
                action["examples"],
            )

    def test_valid_baseline_passes(self) -> None:
        with self._minimal_baseline() as root:

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 0, completed.stderr)
        self.assertIn("1 WPS Actions", completed.stdout)

    def test_unknown_manifest_field_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["risque"] = "read"
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("unknown field 'risque'", completed.stderr)

    def test_reference_group_cannot_name_an_unknown_action(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["reference_groups"] = {
                "common_core": {
                    "reference": "common.md",
                    "actions": ["ping", "notAnAction"],
                }
            }
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("references unknown WPS Action 'notAnAction'", completed.stderr)

    def test_invalid_risk_classification_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["risk"] = "unsafe"
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("risk must be one of", completed.stderr)

    def test_invalid_application_owner_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["application"] = "spreadsheet"
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("application must be one of", completed.stderr)

    def test_invalid_action_identifier_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["action"] = "wps_ping"
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("must be an original lowerCamelCase identifier", completed.stderr)

    def test_missing_required_metadata_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            del manifest["actions"][0]["result"]
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("missing required field 'result'", completed.stderr)

    def test_parameter_property_without_type_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["parameters"]["properties"] = {
                "token": {"description": "Authentication token."}
            }
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("parameter 'token' requires a valid JSON type", completed.stderr)

    def test_unknown_nested_schema_field_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["parameters"]["properties"] = {
                "options": {
                    "type": "object",
                    "properties": {"mode": {"type": "string", "typo": True}},
                    "required": [],
                    "additionalProperties": False,
                }
            }
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("unknown schema field 'typo'", completed.stderr)

    def test_array_items_schema_is_validated_recursively(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["result"]["properties"] = {
                "items": {"type": "array", "items": {"type": "integer"}}
            }
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("requires a valid JSON type", completed.stderr)

    def test_string_pattern_is_a_valid_contract_constraint(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["parameters"] = {
                "type": "object",
                "properties": {
                    "formula": {"type": "string", "pattern": "^="},
                },
                "required": ["formula"],
                "additionalProperties": False,
            }
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 0, completed.stderr)

    def test_result_array_requires_an_items_contract(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["result"]["properties"] = {
                "items": {"type": "array"}
            }
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("array result field requires an items schema", completed.stderr)

    def test_result_object_requires_a_properties_contract(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["result"]["properties"] = {
                "details": {"type": "object"}
            }
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("object result field requires a properties schema", completed.stderr)

    def test_parameter_array_requires_an_items_contract(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["parameters"]["properties"] = {
                "items": {"type": "array"}
            }
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("array parameter requires an items schema", completed.stderr)

    def test_parameter_object_requires_a_properties_contract(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["parameters"]["properties"] = {
                "options": {"type": "object"}
            }
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("object parameter requires a properties schema", completed.stderr)

    def test_example_must_satisfy_any_of_requirement(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["parameters"].update(
                {
                    "properties": {
                        "path": {"type": "string"},
                        "filePath": {"type": "string"},
                    },
                    "anyOf": [{"required": ["path"]}, {"required": ["filePath"]}],
                }
            )
            manifest["actions"][0]["examples"] = [
                {"name": "Missing path", "params": {}, "result": {}}
            ]
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("must satisfy at least one anyOf branch", completed.stderr)

    def test_example_must_match_parameter_contract(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["parameters"] = {
                "type": "object",
                "properties": {"token": {"type": "string"}},
                "required": ["token"],
                "additionalProperties": False,
            }
            manifest["actions"][0]["examples"] = [
                {"name": "Missing token", "params": {}, "result": {}}
            ]
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn(
            "example 'Missing token' params: missing required property 'token'",
            completed.stderr,
        )

    def test_duplicate_action_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"].append(dict(manifest["actions"][0]))
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("duplicate WPS Action 'ping'", completed.stderr)

    def test_unmapped_bridge_action_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            self._write_text(
                root / "wps-claude-assistant/main.js",
                "switch (action) {\n  case 'ping': return true;\n"
                "  case 'hiddenCapability': return true;\n}\n",
            )

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn(
            "JavaScript dispatch has unmanifested WPS Action 'hiddenCapability'",
            completed.stderr,
        )

    def test_unmapped_powershell_action_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            self._write_text(
                root / "wps-office-mcp/scripts/wps-com.ps1",
                'switch ($Action) {\n    "ping" { @{ success = $true } }\n'
                '    "hiddenCapability" { @{ success = $true } }\n}\n',
            )

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn(
            "PowerShell dispatch has unmanifested WPS Action 'hiddenCapability'",
            completed.stderr,
        )

    def test_unexplained_bridge_omission_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            extra = dict(manifest["actions"][0])
            extra["action"] = "powershellOnly"
            manifest["actions"].append(extra)
            self._write_json(manifest_path, manifest)
            self._write_text(
                root / "wps-office-mcp/scripts/wps-com.ps1",
                'switch ($Action) {\n    "ping" { @{ success = $true } }\n'
                '    "powershellOnly" { @{ success = $true } }\n}\n',
            )

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn(
            "WPS Action 'powershellOnly' is missing from JavaScript without an explanation",
            completed.stderr,
        )

    def test_explained_bridge_omission_passes(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            extra = dict(manifest["actions"][0])
            extra["action"] = "powershellOnly"
            manifest["actions"].append(extra)
            self._write_json(manifest_path, manifest)
            self._write_text(
                root / "wps-office-mcp/scripts/wps-com.ps1",
                'switch ($Action) {\n    "ping" { @{ success = $true } }\n'
                '    "powershellOnly" { @{ success = $true } }\n}\n',
            )
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["bridge_exceptions"]["javascript_missing"] = [
                {
                    "action": "powershellOnly",
                    "reason": "The legacy JavaScript add-in has not ported this action yet.",
                }
            ]
            self._write_json(mapping_path, mapping)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 0, completed.stderr)

    def test_bridge_exception_without_reason_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["bridge_exceptions"]["javascript_missing"] = [
                {"action": "ping", "reason": ""}
            ]
            self._write_json(mapping_path, mapping)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("bridge exception for 'ping' requires a reason", completed.stderr)

    def test_unexplained_duplicate_dispatch_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            self._write_text(
                root / "wps-office-mcp/scripts/wps-com.ps1",
                'switch ($Action) {\n    "ping" { @{ success = $true } }\n'
                '    "ping" { @{ success = $true } }\n}\n',
            )

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn(
            "PowerShell dispatch duplicates WPS Action 'ping' without an explanation",
            completed.stderr,
        )

    def test_unmapped_legacy_tool_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            self._write_text(
                root / "wps-office-mcp/src/tools/common/general.ts",
                "const first = { name: 'wps_common_ping' };\n"
                "const second = { name: 'wps_common_hidden' };\n",
            )

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("legacy WPS Tool 'wps_common_hidden' is not mapped", completed.stderr)

    def test_mapping_to_unknown_action_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["legacy_tools"][0]["actions"] = ["notAnAction"]
            self._write_json(mapping_path, mapping)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("maps to unknown WPS Action 'notAnAction'", completed.stderr)

    def test_cross_application_mapping_requires_conflict_status(self) -> None:
        with self._minimal_baseline() as root:
            self._write_text(
                root / "wps-office-mcp/src/tools/common/general.ts",
                "const definition = { name: 'wps_excel_ping' };\n",
            )
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["legacy_tools"][0]["tool"] = "wps_excel_ping"
            self._write_json(mapping_path, mapping)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("owned by excel maps to common WPS Action 'ping'", completed.stderr)

    def test_mapped_tool_without_an_action_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["legacy_tools"][0]["actions"] = []
            self._write_json(mapping_path, mapping)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("status 'mapped' requires exactly one WPS Action", completed.stderr)

    @contextmanager
    def _minimal_baseline(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._write_minimal_baseline(root)
            yield root

    def _write_minimal_baseline(self, root: Path) -> None:
        self._write_json(
            root / "skills/wps-office/references/action-manifest.json",
            {
                "schema_version": 1,
                "actions": [
                    {
                        "action": "ping",
                        "application": "common",
                        "description": "Check whether the Platform Bridge responds.",
                        "parameters": {
                            "type": "object",
                            "properties": {},
                            "required": [],
                            "additionalProperties": False,
                        },
                        "result": {
                            "type": "object",
                            "properties": {},
                            "required": [],
                            "additionalProperties": False,
                            "description": "Connection status returned by WPS.",
                        },
                        "prerequisites": [],
                        "risk": "read",
                    }
                ],
            },
        )
        self._write_json(
            root / "doc/migration/legacy-tool-action-map.json",
            {
                "schema_version": 1,
                "legacy_tools": [
                    {
                        "tool": "wps_common_ping",
                        "actions": ["ping"],
                        "status": "mapped",
                    }
                ],
                "bridge_exceptions": {
                    "javascript_missing": [],
                    "powershell_missing": [],
                    "powershell_duplicate_dispatches": [],
                    "contract_conflicts": [],
                },
            },
        )
        self._write_text(
            root / "wps-claude-assistant/main.js",
            "switch (action) {\n  case 'ping': return true;\n}\n",
        )
        self._write_text(
            root / "wps-office-mcp/scripts/wps-com.ps1",
            'switch ($Action) {\n    "ping" { @{ success = $true } }\n}\n',
        )
        self._write_text(
            root / "wps-office-mcp/src/tools/common/general.ts",
            "const definition = { name: 'wps_common_ping' };\n",
        )
        self._write_text(
            root / "wps-office-mcp/src/server/mcp-server.ts",
            "// No built-in tools in this fixture.\n",
        )

    @staticmethod
    def _run_validator(root: Path) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [sys.executable, str(VALIDATOR), "--root", str(root)],
            check=False,
            capture_output=True,
            text=True,
        )

    @staticmethod
    def _write_json(path: Path, value: object) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(value), encoding="utf-8")

    @staticmethod
    def _write_text(path: Path, value: str) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(value, encoding="utf-8")


class ExcelCoreContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        manifest = json.loads(
            (
                REPOSITORY_ROOT
                / "skills/wps-office/references/action-manifest.json"
            ).read_text(encoding="utf-8")
        )
        cls.actions = {entry["action"]: entry for entry in manifest["actions"]}
        cls.core_actions = set(manifest["reference_groups"]["excel_core"]["actions"])
        cls.addin_source = UNIFIED_ADDIN.read_text(encoding="utf-8")
        cls.reference_source = (
            EXCEL_REFERENCE.read_text(encoding="utf-8")
            if EXCEL_REFERENCE.is_file()
            else ""
        )

    def test_every_core_action_has_a_manifest_contract_and_addin_dispatch(self) -> None:
        dispatch_actions = set(
            re.findall(
                r"^\s*case\s+['\"]([A-Za-z][A-Za-z0-9]*)['\"]\s*:",
                self.addin_source,
                re.MULTILINE,
            )
        )
        self.assertEqual(self.core_actions - set(self.actions), set())
        self.assertEqual(self.core_actions - dispatch_actions, set())

        for action_name in sorted(self.core_actions):
            with self.subTest(action=action_name):
                action = self.actions[action_name]
                self.assertEqual(action["application"], "excel")
                self.assertEqual(action["parameters"]["type"], "object")
                self.assertEqual(action["result"]["type"], "object")
                self.assertFalse(action["parameters"]["additionalProperties"])
                self.assertFalse(action["result"]["additionalProperties"])
                self.assertIn(action["risk"], {"read", "write", "destructive"})

                self.assertLessEqual(
                    set(action["parameters"]["required"]),
                    set(action["parameters"]["properties"]),
                )
                self.assertLessEqual(
                    set(action["result"]["required"]),
                    set(action["result"]["properties"]),
                )

    def test_core_actions_require_inputs_needed_by_the_addin(self) -> None:
        required_inputs = {
            "switchSheet": {"sheet"},
            "deleteRows": {"row"},
            "findInSheet": {"searchText"},
            "replaceInSheet": {"searchText", "replaceText"},
        }

        for action_name, expected in required_inputs.items():
            with self.subTest(action=action_name):
                actual = set(self.actions[action_name]["parameters"]["required"])
                self.assertEqual(actual, expected)

    def test_row_actions_return_numeric_row_positions(self) -> None:
        result_positions = {
            "insertRows": "insertedAt",
            "deleteRows": "deletedFrom",
        }

        for action_name, field in result_positions.items():
            with self.subTest(action=action_name):
                field_schema = self.actions[action_name]["result"]["properties"][field]
                self.assertEqual(field_schema["type"], "number")

    def test_cell_info_and_sheet_positions_match_addin_values(self) -> None:
        cell_info = self.actions["getCellInfo"]["result"]["properties"]

        self.assertEqual(cell_info["backgroundColor"]["type"], "number")
        self.assertEqual(cell_info["font"]["type"], "object")
        self.assertEqual(
            cell_info["value"]["type"],
            ["string", "number", "boolean", "null"],
        )
        self.assertEqual(
            self.actions["copySheet"]["parameters"]["properties"]["before"]["type"],
            "string",
        )
        self.assertEqual(
            self.actions["moveSheet"]["parameters"]["properties"]["before"]["type"],
            "string",
        )

    def test_clean_data_rejects_unknown_operations_before_the_addin(self) -> None:
        operation_items = self.actions["cleanData"]["parameters"]["properties"][
            "operations"
        ]["items"]

        self.assertEqual(
            operation_items["enum"],
            ["trim", "remove_duplicates", "unify_date", "remove_empty_rows"],
        )
        self.assertEqual(self.actions["cleanData"]["risk"], "destructive")

    def test_clear_range_rejects_unknown_clear_types(self) -> None:
        clear_type = self.actions["clearRange"]["parameters"]["properties"]["type"]

        self.assertEqual(
            clear_type["enum"],
            ["all", "contents", "formats", "comments"],
        )

    def test_find_and_replace_pass_match_case_in_the_wps_match_case_position(self) -> None:
        self.assertIn(
            "searchRange.Find(params.searchText, null, -4163, 2, 1, 1, !!params.matchCase)",
            self.addin_source,
        )
        self.assertIn(
            "searchRange.Replace(params.searchText, params.replaceText, 2, 1, !!params.matchCase)",
            self.addin_source,
        )

    def test_excel_reference_uses_only_the_skill_package_runtime(self) -> None:
        self.assertTrue(EXCEL_REFERENCE.is_file())
        self.assertIn("scripts/wps.py invoke", self.reference_source)
        self.assertNotIn("MCP", self.reference_source)
        self.assertNotIn("Node.js", self.reference_source)
        self.assertNotIn("PowerShell", self.reference_source)

    def test_excel_reference_documents_the_complete_core_workflow(self) -> None:
        workflow_actions = (
            "getActiveWorkbook",
            "switchSheet",
            "setRangeData",
            "setFormula",
            "getRangeData",
            "save",
        )

        positions = [
            self.reference_source.find(f"`{action}`") for action in workflow_actions
        ]
        self.assertTrue(all(position >= 0 for position in positions))
        self.assertEqual(positions, sorted(positions))

    def test_excel_reference_catalog_is_derived_from_the_manifest_group(self) -> None:
        start = self.reference_source.index("<!-- excel-core-actions:start -->")
        end = self.reference_source.index("<!-- excel-core-actions:end -->")
        documented = set(
            re.findall(r"`([a-z][A-Za-z0-9]*)`", self.reference_source[start:end])
        )

        self.assertEqual(documented, self.core_actions)


if __name__ == "__main__":
    unittest.main()
