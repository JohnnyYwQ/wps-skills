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
LEGACY_MAC_MAIN = REPOSITORY_ROOT / "wps-claude-assistant/main.js"
LEGACY_MAC_WORD_HANDLER = (
    REPOSITORY_ROOT / "wps-claude-assistant/handlers/word-handler.js"
)
WINDOWS_BRIDGE = REPOSITORY_ROOT / "wps-office-mcp/scripts/wps-com.ps1"
EXCEL_REFERENCE = REPOSITORY_ROOT / "skills/wps-office/references/excel.md"
WORD_REFERENCE = REPOSITORY_ROOT / "skills/wps-office/references/word.md"
POWERPOINT_REFERENCE = (
    REPOSITORY_ROOT / "skills/wps-office/references/powerpoint.md"
)


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

    def test_numeric_bounds_are_valid_contract_constraints(self) -> None:
        with self._minimal_baseline() as root:
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["parameters"] = {
                "type": "object",
                "properties": {
                    "percent": {
                        "type": "number",
                        "minimum": 10,
                        "maximum": 400,
                    },
                },
                "required": ["percent"],
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

    def test_cross_application_mapping_is_rejected(self) -> None:
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

    def test_conflict_status_is_not_an_available_mapping_escape_hatch(self) -> None:
        with self._minimal_baseline() as root:
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["legacy_tools"][0].update(
                {
                    "actions": [],
                    "status": "conflict",
                    "reason": "Temporary conflict exceptions are not permitted.",
                }
            )
            self._write_json(mapping_path, mapping)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("has invalid status 'conflict'", completed.stderr)

    def test_stale_contract_conflict_exceptions_are_rejected(self) -> None:
        with self._minimal_baseline() as root:
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["bridge_exceptions"]["contract_conflicts"] = []
            self._write_json(mapping_path, mapping)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("unknown bridge exception category 'contract_conflicts'", completed.stderr)

    def test_mapped_tool_without_an_action_is_rejected(self) -> None:
        with self._minimal_baseline() as root:
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["legacy_tools"][0]["actions"] = []
            self._write_json(mapping_path, mapping)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("status 'mapped' requires exactly one WPS Action", completed.stderr)

    def test_retired_action_must_not_be_exposed_by_the_manifest(self) -> None:
        with self._minimal_baseline() as root:
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["retired_actions"] = [
                {
                    "action": "ping",
                    "decision": "ADR-0014",
                    "reason": "A reviewed retirement example.",
                    "replacements": [],
                }
            ]
            self._write_json(mapping_path, mapping)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("retired WPS Action 'ping' is still exposed", completed.stderr)

    def test_retired_action_requires_a_reviewed_decision_and_reason(self) -> None:
        with self._minimal_baseline() as root:
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["retired_actions"] = [
                {
                    "action": "oldAction",
                    "decision": "",
                    "reason": "",
                    "replacements": [],
                }
            ]
            self._write_json(mapping_path, mapping)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("retired WPS Action 'oldAction' requires a reviewed decision", completed.stderr)

    def test_retired_legacy_tool_requires_a_reviewed_decision(self) -> None:
        with self._minimal_baseline() as root:
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["legacy_tools"][0].update(
                {"actions": [], "status": "retired", "reason": "Retired example."}
            )
            self._write_json(mapping_path, mapping)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("status 'retired' requires a reviewed decision", completed.stderr)

    def test_retired_contract_parameters_must_not_remain_in_the_manifest(self) -> None:
        with self._minimal_baseline() as root:
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["retired_contracts"] = [
                {
                    "action": "ping",
                    "decision": "ADR-0014",
                    "reason": "A reviewed contract retirement example.",
                    "retired_parameters": ["token"],
                }
            ]
            manifest_path = root / "skills/wps-office/references/action-manifest.json"
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["actions"][0]["parameters"]["properties"]["token"] = {
                "type": "string"
            }
            self._write_json(mapping_path, mapping)
            self._write_json(manifest_path, manifest)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("retired parameter 'token' is still exposed", completed.stderr)

    def test_current_ledger_schema_rejects_an_unreviewed_retirement_boundary(self) -> None:
        with self._minimal_baseline() as root:
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["schema_version"] = 2
            mapping["retired_actions"] = [
                {
                    "action": "oldAction",
                    "decision": "ADR-0014",
                    "reason": "An arbitrary retirement is not accepted.",
                    "replacements": [],
                }
            ]
            self._write_json(mapping_path, mapping)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("not the reviewed ADR-0014 boundary", completed.stderr)

    def test_retirement_entries_cannot_bypass_the_current_ledger_schema(self) -> None:
        with self._minimal_baseline() as root:
            mapping_path = root / "doc/migration/legacy-tool-action-map.json"
            mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
            mapping["retired_actions"] = [
                {
                    "action": "oldAction",
                    "decision": "ADR-0014",
                    "reason": "An arbitrary retirement is not accepted.",
                    "replacements": [],
                }
            ]
            self._write_json(mapping_path, mapping)

            completed = self._run_validator(root)

        self.assertEqual(completed.returncode, 1)
        self.assertIn("retirement entries require", completed.stderr)

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
                "retired_actions": [],
                "retired_contracts": [],
                "bridge_exceptions": {
                    "javascript_missing": [],
                    "powershell_missing": [],
                    "powershell_duplicate_dispatches": [],
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


class RetirementContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.manifest = json.loads(
            (REPOSITORY_ROOT / "skills/wps-office/references/action-manifest.json").read_text(
                encoding="utf-8"
            )
        )
        cls.actions = {entry["action"]: entry for entry in cls.manifest["actions"]}
        cls.ledger = json.loads(
            (REPOSITORY_ROOT / "doc/migration/legacy-tool-action-map.json").read_text(
                encoding="utf-8"
            )
        )
        cls.tools = {entry["tool"]: entry for entry in cls.ledger["legacy_tools"]}

    def test_retired_actions_are_replaced_by_canonical_workflows(self) -> None:
        retirements = {entry["action"]: entry for entry in self.ledger["retired_actions"]}

        self.assertNotIn("openFile", self.actions)
        self.assertEqual(
            retirements["openFile"]["replacements"],
            ["openDocument", "openWorkbook", "openPresentation"],
        )
        self.assertNotIn("convertFormat", self.actions)
        self.assertEqual(
            retirements["convertFormat"]["replacements"],
            ["convertToPDF", "save", "saveAs"],
        )
        for retirement in retirements.values():
            self.assertEqual(retirement["decision"], "ADR-0014")
            self.assertTrue(retirement["reason"].strip())
            self.assertTrue(set(retirement["replacements"]).issubset(self.actions))

    def test_accepted_legacy_retirements_are_auditable(self) -> None:
        expected_retired_tools = {
            "wps_cache_data",
            "wps_clear_cache",
            "wps_execute_method",
            "wps_get_cached_data",
            "wps_list_cache",
            "wps_word_proofread_basic",
            "wps_convert_format",
            "wps_excel_find_replace",
        }
        retired_tools = {
            tool_name
            for tool_name, entry in self.tools.items()
            if entry["status"] == "retired"
        }

        self.assertEqual(retired_tools, expected_retired_tools)
        for tool_name in expected_retired_tools:
            with self.subTest(tool=tool_name):
                entry = self.tools[tool_name]
                self.assertEqual(entry["actions"], [])
                self.assertEqual(entry["decision"], "ADR-0014")
                self.assertTrue(entry["reason"].strip())

    def test_application_references_use_canonical_open_and_excel_replace_actions(self) -> None:
        references = {
            "excel": EXCEL_REFERENCE.read_text(encoding="utf-8"),
            "word": WORD_REFERENCE.read_text(encoding="utf-8"),
            "powerpoint": POWERPOINT_REFERENCE.read_text(encoding="utf-8"),
        }

        self.assertIn("`openWorkbook`", references["excel"])
        self.assertIn("`openDocument`", references["word"])
        self.assertIn("`openPresentation`", references["powerpoint"])
        self.assertIn("`replaceInSheet`", references["excel"])
        self.assertNotIn("`findReplace`", references["excel"])
        for source in references.values():
            self.assertNotIn("`openFile`", source)
            self.assertNotIn("`convertFormat`", source)

    def test_add_arrow_uses_only_the_packaged_coordinate_contract(self) -> None:
        retirement = self.ledger["retired_contracts"][0]
        parameters = self.actions["addArrow"]["parameters"]["properties"]
        powershell_source = WINDOWS_BRIDGE.read_text(encoding="utf-8")
        arrow_dispatch = powershell_source.split('    "addArrow" {', 1)[1].split(
            '\n    "', 1
        )[0]

        self.assertEqual(retirement["action"], "addArrow")
        self.assertEqual(retirement["decision"], "ADR-0014")
        self.assertEqual(
            retirement["retired_parameters"], ["height", "left", "top", "width"]
        )
        self.assertTrue({"startX", "startY", "endX", "endY"}.issubset(parameters))
        self.assertFalse(set(retirement["retired_parameters"]) & set(parameters))
        for parameter in ("startX", "startY", "endX", "endY"):
            self.assertIn(f"$p.{parameter}", arrow_dispatch)
        for retired_parameter in retirement["retired_parameters"]:
            self.assertNotIn(f"$p.{retired_parameter}", arrow_dispatch)

    def test_corrected_cross_application_mappings_are_canonical(self) -> None:
        mappings = {entry["tool"]: entry for entry in self.ledger["legacy_tools"]}
        excel_source = (
            REPOSITORY_ROOT / "wps-office-mcp/src/tools/excel/data.ts"
        ).read_text(encoding="utf-8")
        powerpoint_source = (
            REPOSITORY_ROOT / "wps-office-mcp/src/tools/ppt/presentation.ts"
        ).read_text(encoding="utf-8")

        self.assertEqual(
            mappings["wps_excel_add_comment"],
            {
                "tool": "wps_excel_add_comment",
                "actions": ["addCellComment"],
                "status": "mapped",
            },
        )
        self.assertEqual(
            mappings["wps_ppt_insert_slide_image"],
            {
                "tool": "wps_ppt_insert_slide_image",
                "actions": ["insertPptImage"],
                "status": "mapped",
            },
        )
        self.assertFalse(
            any(entry["status"] == "conflict" for entry in mappings.values())
        )
        self.assertEqual(
            set(self.actions["addArrow"]["parameters"]["required"]),
            {"startX", "startY", "endX", "endY"},
        )
        self.assertIn("'addCellComment'", excel_source)
        self.assertNotIn("'addComment'", excel_source)
        self.assertIn("'insertPptImage'", powerpoint_source)
        self.assertNotIn("'insertImage'", powerpoint_source)


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


class ExcelAdvancedContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.manifest = json.loads(
            (
                REPOSITORY_ROOT
                / "skills/wps-office/references/action-manifest.json"
            ).read_text(encoding="utf-8")
        )
        cls.actions = {entry["action"]: entry for entry in cls.manifest["actions"]}
        cls.excel_actions = {
            entry["action"]
            for entry in cls.manifest["actions"]
            if entry["application"] == "excel"
        }
        cls.core_actions = set(
            cls.manifest["reference_groups"]["excel_core"]["actions"]
        )
        advanced_group = cls.manifest["reference_groups"].get(
            "excel_advanced", {"actions": []}
        )
        cls.advanced_group = advanced_group
        cls.advanced_actions = set(advanced_group["actions"])
        cls.addin_source = UNIFIED_ADDIN.read_text(encoding="utf-8")
        cls.reference_source = EXCEL_REFERENCE.read_text(encoding="utf-8")

    def test_core_and_advanced_groups_partition_every_excel_action(self) -> None:
        self.assertEqual(self.core_actions & self.advanced_actions, set())
        self.assertEqual(self.core_actions | self.advanced_actions, self.excel_actions)
        self.assertEqual(self.advanced_group.get("reference"), "excel.md")

    def test_every_advanced_action_has_a_contract_and_addin_dispatch(self) -> None:
        dispatch_actions = set(
            re.findall(
                r"^\s*case\s+['\"]([A-Za-z][A-Za-z0-9]*)['\"]\s*:",
                self.addin_source,
                re.MULTILINE,
            )
        )

        self.assertEqual(self.advanced_actions - dispatch_actions, set())
        for action_name in sorted(self.advanced_actions):
            with self.subTest(action=action_name):
                action = self.actions[action_name]
                self.assertEqual(action["application"], "excel")
                self.assertEqual(action["parameters"]["type"], "object")
                self.assertEqual(action["result"]["type"], "object")
                self.assertFalse(action["parameters"]["additionalProperties"])
                self.assertFalse(action["result"]["additionalProperties"])
                self.assertIn(action["risk"], {"read", "write", "destructive"})

    def test_advanced_actions_require_inputs_used_unconditionally(self) -> None:
        required_inputs = {
            "addCellComment": {"cell", "text"},
            "addDataValidation": {"range", "validationType"},
            "clearFormats": {"range"},
            "consolidate": {"destination", "sources"},
            "copyFormat": {"source", "target"},
            "copyRange": {"range"},
            "createPivotTable": {"sourceRange", "destinationCell"},
            "deleteNamedRange": {"name"},
            "groupColumns": {"startColumn", "endColumn"},
            "getConditionalFormats": {"range"},
            "getDataValidations": {"range"},
            "removeConditionalFormat": {"range"},
            "removeDataValidation": {"range"},
            "setCellFormat": {"range", "numberFormat"},
            "transpose": {"sourceRange"},
            "wrapText": {"range"},
        }

        for action_name, expected in required_inputs.items():
            with self.subTest(action=action_name):
                actual = set(self.actions[action_name]["parameters"]["required"])
                self.assertEqual(actual, expected)

    def test_actions_with_alternative_identifiers_declare_any_of_contracts(self) -> None:
        expected_alternatives = {
            "hideColumns": [{"required": ["column"]}, {"required": ["columns"]}],
            "hideRows": [{"required": ["row"]}, {"required": ["rows"]}],
            "showColumns": [{"required": ["column"]}, {"required": ["columns"]}],
            "showRows": [{"required": ["row"]}, {"required": ["rows"]}],
            "subtotal": [
                {"required": ["totalColumn"]},
                {"required": ["totalColumns"]},
            ],
            "transpose": [{"required": ["destinationCell"]}, {"required": ["targetCell"]}],
            "updateChart": [{"required": ["chartIndex"]}, {"required": ["chartName"]}],
            "updatePivotTable": [
                {"required": ["pivotTableName"]},
                {"required": ["pivotTableCell"]},
            ],
        }

        for action_name, expected in expected_alternatives.items():
            with self.subTest(action=action_name):
                self.assertEqual(
                    self.actions[action_name]["parameters"].get("anyOf"), expected
                )

    def test_advanced_result_types_match_the_addin_values(self) -> None:
        self.assertEqual(
            self.actions["wrapText"]["result"]["properties"]["wrapText"]["type"],
            "boolean",
        )
        self.assertEqual(
            self.actions["getDataValidations"]["result"]["properties"]["type"]["type"],
            "number",
        )
        for action_name, result_field in (
            ("hideColumns", "hiddenColumns"),
            ("hideRows", "hiddenRows"),
            ("showColumns", "shownColumns"),
            ("showRows", "shownRows"),
        ):
            with self.subTest(action=action_name):
                result = self.actions[action_name]["result"]["properties"][result_field]
                self.assertEqual(result["type"], "array")
                self.assertEqual(result["items"]["type"], "number")

    def test_context_actions_share_the_unified_structured_result(self) -> None:
        expected_fields = {
            "workbookName",
            "currentSheet",
            "allSheets",
            "selectedCell",
            "usedRange",
            "usedRangeAddress",
            "headers",
            "rowCount",
            "colCount",
        }

        for action_name in ("getContext", "getExcelContext"):
            with self.subTest(action=action_name):
                result = self.actions[action_name]["result"]
                self.assertEqual(set(result["properties"]), expected_fields)
                self.assertEqual(set(result["required"]), expected_fields)

    def test_set_zoom_rejects_values_outside_the_wps_range(self) -> None:
        percent = self.actions["setZoom"]["parameters"]["properties"]["percent"]

        self.assertEqual(percent["minimum"], 10)
        self.assertEqual(percent["maximum"], 400)
        self.assertEqual(
            self.actions["setZoom"]["result"]["required"], ["percent"]
        )

    def test_overwriting_advanced_actions_are_destructive(self) -> None:
        for action_name in ("addCellComment", "pasteRange"):
            with self.subTest(action=action_name):
                self.assertEqual(self.actions[action_name]["risk"], "destructive")

    def test_pivot_contracts_preserve_the_destination_sheet_and_field_names(self) -> None:
        create_pivot = self.actions["createPivotTable"]
        self.assertIn("sheet", create_pivot["result"]["required"])
        self.assertIn("sheet", self.actions["updatePivotTable"]["parameters"]["properties"])

        for action_name, property_name in (
            ("createPivotTable", "valueFields"),
            ("updatePivotTable", "addValueFields"),
            ("updatePivotTable", "updateValueFields"),
        ):
            with self.subTest(action=action_name, property=property_name):
                item_contract = self.actions[action_name]["parameters"]["properties"][
                    property_name
                ]["items"]
                self.assertEqual(item_contract["required"], ["field"])

    def test_excel_reference_catalog_is_derived_from_the_advanced_group(self) -> None:
        start = self.reference_source.index("<!-- excel-advanced-actions:start -->")
        end = self.reference_source.index("<!-- excel-advanced-actions:end -->")
        documented = set(
            re.findall(r"`([a-z][A-Za-z0-9]*)`", self.reference_source[start:end])
        )

        self.assertEqual(documented, self.advanced_actions)


class WordContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.manifest = json.loads(
            (
                REPOSITORY_ROOT
                / "skills/wps-office/references/action-manifest.json"
            ).read_text(encoding="utf-8")
        )
        cls.actions = {entry["action"]: entry for entry in cls.manifest["actions"]}
        cls.word_actions = {
            entry["action"]
            for entry in cls.manifest["actions"]
            if entry["application"] == "word"
        }
        cls.word_group = cls.manifest["reference_groups"].get(
            "word_workflows", {"actions": []}
        )
        cls.group_actions = set(cls.word_group["actions"])
        cls.addin_source = UNIFIED_ADDIN.read_text(encoding="utf-8")
        cls.legacy_mac_main_source = LEGACY_MAC_MAIN.read_text(encoding="utf-8")
        cls.legacy_mac_word_handler_source = LEGACY_MAC_WORD_HANDLER.read_text(
            encoding="utf-8"
        )
        cls.reference_source = (
            WORD_REFERENCE.read_text(encoding="utf-8")
            if WORD_REFERENCE.is_file()
            else ""
        )

    def test_word_group_covers_every_word_action_and_addin_dispatch(self) -> None:
        dispatch_actions = set(
            re.findall(
                r"^\s*case\s+['\"]([A-Za-z][A-Za-z0-9]*)['\"]\s*:",
                self.addin_source,
                re.MULTILINE,
            )
        )

        self.assertEqual(self.group_actions, self.word_actions)
        self.assertEqual(self.word_group.get("reference"), "word.md")
        self.assertEqual(self.word_actions - dispatch_actions, set())

        for action_name in sorted(self.word_actions):
            with self.subTest(action=action_name):
                action = self.actions[action_name]
                self.assertEqual(action["parameters"]["type"], "object")
                self.assertEqual(action["result"]["type"], "object")
                self.assertFalse(action["parameters"]["additionalProperties"])
                self.assertFalse(action["result"]["additionalProperties"])
                self.assertIn(action["risk"], {"read", "write", "destructive"})

    def test_word_inputs_match_the_addin_requirements(self) -> None:
        required_inputs = {
            "findInDocument": {"findText"},
            "findReplace": {"findText", "replaceText"},
            "insertBookmark": {"name"},
            "insertHeader": {"text"},
            "insertFooter": {"text"},
            "insertTable": {"rows", "cols"},
            "insertText": {"text"},
            "openDocument": {"path"},
            "replaceBookmarkContent": {"name", "text"},
            "replaceRange": {"startPos", "endPos", "text"},
            "setLineSpacing": {"lineSpacing"},
            "setTextColor": {"color"},
            "smartFillField": {"keyword", "value"},
        }

        for action_name, expected in required_inputs.items():
            with self.subTest(action=action_name):
                self.assertEqual(
                    set(self.actions[action_name]["parameters"]["required"]),
                    expected,
                )

        self.assertEqual(
            self.actions["switchDocument"]["parameters"].get("anyOf"),
            [{"required": ["name"]}, {"required": ["index"]}],
        )
        self.assertEqual(
            self.actions["insertHyperlink"]["parameters"].get("anyOf"),
            [{"required": ["url"]}, {"required": ["address"]}],
        )

    def test_word_range_and_revision_results_use_runtime_types(self) -> None:
        replace_result = self.actions["replaceRange"]["result"]["properties"]
        for field in ("startPos", "originalEndPos", "endPos"):
            self.assertEqual(replace_result[field]["type"], "number")

        paragraph_item = self.actions["getDocumentParagraphs"]["result"][
            "properties"
        ]["paragraphs"]["items"]
        self.assertEqual(
            set(paragraph_item["required"]),
            {"index", "text", "style", "start", "end"},
        )
        self.assertEqual(
            self.actions["getTrackChangesStatus"]["result"]["required"],
            ["trackChanges", "revisionCount"],
        )
        self.assertEqual(
            self.actions["setFont"]["result"]["properties"]["settings"]["type"],
            "object",
        )

    def test_destructive_word_actions_are_gated(self) -> None:
        for action_name in (
            "closeDocument",
            "findReplace",
            "replaceBookmarkContent",
            "replaceRange",
            "smartFillField",
        ):
            with self.subTest(action=action_name):
                self.assertEqual(self.actions[action_name]["risk"], "destructive")

    def test_legacy_mac_bridge_dispatches_template_actions(self) -> None:
        template_actions = {
            "getDocumentParagraphs": "getDocumentParagraphs",
            "findInDocument": "findInDocument",
            "smartFillField": "smartFillField",
            "replaceBookmarkContent": "replaceBookmarkContent",
        }
        dispatch_actions = set(
            re.findall(
                r"^\s*case\s+['\"]([A-Za-z][A-Za-z0-9]*)['\"]\s*:",
                self.legacy_mac_main_source,
                re.MULTILINE,
            )
        )

        for action, function_name in template_actions.items():
            with self.subTest(action=action):
                main_function_name = f"handle{function_name[0].upper()}{function_name[1:]}"
                self.assertIn(action, dispatch_actions)
                self.assertRegex(
                    self.legacy_mac_main_source,
                    rf"function {main_function_name}\s*\(",
                )
                self.assertRegex(
                    self.legacy_mac_word_handler_source,
                    rf"function {function_name}\s*\(",
                )
                self.assertIn(
                    f"{function_name}: {function_name}",
                    self.legacy_mac_word_handler_source,
                )

    def test_find_in_document_does_not_wrap_and_repeat_matches(self) -> None:
        bridge_sources = {
            "unified add-in": self.addin_source,
            "legacy mac main": self.legacy_mac_main_source,
            "legacy mac handler": self.legacy_mac_word_handler_source,
        }
        for bridge_name, source in bridge_sources.items():
            with self.subTest(bridge=bridge_name):
                start = source.index(
                    "function handleFindInDocument"
                    if bridge_name != "legacy mac handler"
                    else "function findInDocument"
                )
                end = source.find("\nfunction ", start + 1)
                block = source[start:] if end < 0 else source[start:end]
                self.assertRegex(
                    block,
                    r"true,\s*\n\s*0,\s*\n\s*false,",
                    "Find.Execute must use wdFindStop (0), not wdFindContinue (1)",
                )

        windows_block = WINDOWS_BRIDGE.read_text(encoding="utf-8")
        start = windows_block.index('"findInDocument" {')
        end = windows_block.index('"smartFillField" {', start)
        find_block = windows_block[start:end]
        self.assertNotIn(
            "$true, 1, $false, \"\", 0",
            find_block,
            "PowerShell Find.Execute must not wrap to the start of the document",
        )
        self.assertIn("$true, 0, $false, \"\", 0", find_block)

    def test_windows_bridge_is_bom_free(self) -> None:
        self.assertTrue(
            WINDOWS_BRIDGE.read_bytes().startswith(b"# Input:"),
            "PowerShell bridge must start with its comment, not a UTF-8 BOM",
        )

    def test_word_reference_documents_workflow_and_complete_catalog(self) -> None:
        self.assertTrue(WORD_REFERENCE.is_file())
        self.assertIn("scripts/wps.py invoke", self.reference_source)
        self.assertNotIn("MCP", self.reference_source)
        self.assertNotIn("Node.js", self.reference_source)
        self.assertNotIn("PowerShell", self.reference_source)

        workflow_actions = (
            "createDocument",
            "insertText",
            "getDocumentParagraphs",
            "findInDocument",
            "replaceRange",
            "getTrackChangesStatus",
            "save",
        )
        positions = [
            self.reference_source.find(f"`{action}`") for action in workflow_actions
        ]
        self.assertTrue(all(position >= 0 for position in positions))
        self.assertEqual(positions, sorted(positions))

        start = self.reference_source.index("<!-- word-actions:start -->")
        end = self.reference_source.index("<!-- word-actions:end -->")
        documented = set(
            re.findall(r"`([a-z][A-Za-z0-9]*)`", self.reference_source[start:end])
        )
        self.assertEqual(documented, self.word_actions)


class PowerPointCoreContractTests(unittest.TestCase):
    CORE_ACTIONS = {
        "addSlide",
        "addTextBox",
        "closePresentation",
        "createPresentation",
        "deletePptImage",
        "deleteSlide",
        "deleteTextBox",
        "duplicateSlide",
        "exportSlideAsImage",
        "getActivePresentation",
        "getOpenPresentations",
        "getPptTableCell",
        "getSlideCount",
        "getSlideInfo",
        "getSlideNotes",
        "getSlideTitle",
        "getTextBoxes",
        "insertPptImage",
        "insertPptTable",
        "moveSlide",
        "openPresentation",
        "setBackgroundColor",
        "setBackgroundImage",
        "setImageStyle",
        "setPptTableCell",
        "setPptTableCellStyle",
        "setPptTableRowStyle",
        "setPptTableStyle",
        "setSlideBackground",
        "setSlideContent",
        "setSlideLayout",
        "setSlideNotes",
        "setSlideSize",
        "setSlideSubtitle",
        "setSlideTitle",
        "setTextBoxStyle",
        "setTextBoxText",
        "switchPresentation",
        "switchSlide",
    }

    @classmethod
    def setUpClass(cls) -> None:
        cls.manifest = json.loads(
            (
                REPOSITORY_ROOT
                / "skills/wps-office/references/action-manifest.json"
            ).read_text(encoding="utf-8")
        )
        cls.actions = {entry["action"]: entry for entry in cls.manifest["actions"]}
        cls.powerpoint_group = cls.manifest["reference_groups"].get(
            "powerpoint_core", {"actions": []}
        )
        cls.group_actions = set(cls.powerpoint_group["actions"])
        cls.addin_source = UNIFIED_ADDIN.read_text(encoding="utf-8")
        cls.reference_source = (
            POWERPOINT_REFERENCE.read_text(encoding="utf-8")
            if POWERPOINT_REFERENCE.is_file()
            else ""
        )

    def test_core_group_matches_scope_and_addin_dispatch(self) -> None:
        dispatch_actions = set(
            re.findall(
                r"^\s*case\s+['\"]([A-Za-z][A-Za-z0-9]*)['\"]\s*:",
                self.addin_source,
                re.MULTILINE,
            )
        )

        self.assertEqual(self.group_actions, self.CORE_ACTIONS)
        self.assertEqual(self.powerpoint_group.get("reference"), "powerpoint.md")
        self.assertEqual(self.CORE_ACTIONS - dispatch_actions, set())

        for action_name in sorted(self.CORE_ACTIONS):
            with self.subTest(action=action_name):
                action = self.actions[action_name]
                self.assertEqual(action["application"], "powerpoint")
                self.assertEqual(action["parameters"]["type"], "object")
                self.assertEqual(action["result"]["type"], "object")
                self.assertFalse(action["parameters"]["additionalProperties"])
                self.assertFalse(action["result"]["additionalProperties"])
                self.assertIn(action["risk"], {"read", "write", "destructive"})

    def test_core_inputs_match_the_addin_requirements(self) -> None:
        required_inputs = {
            "openPresentation": {"path"},
            "setSlideContent": {"slideIndex", "content"},
            "setSlideLayout": {"slideIndex", "layout"},
            "setSlideNotes": {"slideIndex", "notes"},
            "setSlideSize": {"width", "height"},
            "setSlideSubtitle": {"slideIndex", "subtitle"},
            "setSlideTitle": {"slideIndex", "title"},
            "insertPptImage": {"slideIndex", "path"},
            "insertPptTable": {"slideIndex", "rows", "cols"},
        }

        for action_name, expected in required_inputs.items():
            with self.subTest(action=action_name):
                self.assertEqual(
                    set(self.actions[action_name]["parameters"]["required"]),
                    expected,
                )

        self.assertEqual(
            self.actions["addSlide"]["parameters"]["properties"]["layout"]["enum"],
            ["title", "title_content", "blank", "two_column", "comparison"],
        )
        self.assertEqual(
            self.actions["setSlideLayout"]["parameters"]["properties"]["layout"][
                "enum"
            ],
            [
                "title",
                "title_content",
                "blank",
                "two_column",
                "comparison",
                "title_only",
            ],
        )

        for action_name in (
            "deletePptImage",
            "deleteTextBox",
            "getPptTableCell",
            "setImageStyle",
            "setPptTableCellStyle",
            "setPptTableStyle",
            "setTextBoxStyle",
            "setTextBoxText",
        ):
            with self.subTest(action=action_name):
                self.assertEqual(
                    self.actions[action_name]["parameters"].get("anyOf"),
                    [{"required": ["name"]}, {"required": ["shapeIndex"]}]
                    if action_name
                    in {
                        "deletePptImage",
                        "deleteTextBox",
                        "setImageStyle",
                        "setTextBoxStyle",
                        "setTextBoxText",
                    }
                    else [
                        {"required": ["tableName"]},
                        {"required": ["tableIndex"]},
                    ],
                )

        self.assertEqual(
            self.actions["setPptTableCell"]["parameters"].get("anyOf"),
            [
                {"required": ["tableName", "text"]},
                {"required": ["tableName", "value"]},
                {"required": ["tableIndex", "text"]},
                {"required": ["tableIndex", "value"]},
            ],
        )

    def test_context_results_describe_observable_presentation_state(self) -> None:
        presentation_result = self.actions["getActivePresentation"]["result"]
        self.assertEqual(
            presentation_result["required"], ["name", "path", "slideCount"]
        )

        slide_result = self.actions["getSlideInfo"]["result"]
        self.assertEqual(
            set(slide_result["required"]),
            {"slideIndex", "layout", "shapeCount", "shapes"},
        )
        shape_item = slide_result["properties"]["shapes"]["items"]
        self.assertEqual(
            set(shape_item["required"]), {"index", "name", "type", "hasText", "text"}
        )

        table_result = self.actions["getPptTableCell"]["result"]
        self.assertEqual(
            set(table_result["required"]), {"row", "col", "value"}
        )

        row_style = self.actions["setPptTableRowStyle"]
        self.assertEqual(
            row_style["parameters"]["properties"]["alignment"]["enum"],
            ["left", "center", "right"],
        )
        self.assertEqual(set(row_style["result"]["required"]), {"row", "cols"})

    def test_destructive_core_actions_are_gated(self) -> None:
        for action_name in (
            "closePresentation",
            "deletePptImage",
            "deleteSlide",
            "deleteTextBox",
            "exportSlideAsImage",
        ):
            with self.subTest(action=action_name):
                self.assertEqual(self.actions[action_name]["risk"], "destructive")

    def test_powerpoint_reference_documents_workflow_and_core_catalog(self) -> None:
        self.assertTrue(POWERPOINT_REFERENCE.is_file())
        self.assertIn("scripts/wps.py invoke", self.reference_source)
        self.assertNotIn("MCP", self.reference_source)
        self.assertNotIn("Node.js", self.reference_source)
        self.assertNotIn("PowerShell", self.reference_source)

        workflow_actions = (
            "createPresentation",
            "addSlide",
            "setSlideContent",
            "insertPptImage",
            "insertPptTable",
            "getSlideInfo",
            "save",
        )
        positions = [
            self.reference_source.find(f"`{action}`") for action in workflow_actions
        ]
        self.assertTrue(all(position >= 0 for position in positions))
        self.assertEqual(positions, sorted(positions))

        start = self.reference_source.index("<!-- powerpoint-core-actions:start -->")
        end = self.reference_source.index("<!-- powerpoint-core-actions:end -->")
        documented = set(
            re.findall(r"`([a-z][A-Za-z0-9]*)`", self.reference_source[start:end])
        )
        self.assertEqual(documented, self.CORE_ACTIONS)


class PowerPointAdvancedContractTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.manifest = json.loads(
            (
                REPOSITORY_ROOT
                / "skills/wps-office/references/action-manifest.json"
            ).read_text(encoding="utf-8")
        )
        cls.actions = {entry["action"]: entry for entry in cls.manifest["actions"]}
        cls.powerpoint_actions = {
            entry["action"]
            for entry in cls.manifest["actions"]
            if entry["application"] == "powerpoint"
        }
        cls.core_actions = set(
            cls.manifest["reference_groups"]["powerpoint_core"]["actions"]
        )
        cls.advanced_group = cls.manifest["reference_groups"].get(
            "powerpoint_advanced", {"actions": []}
        )
        cls.advanced_actions = set(cls.advanced_group["actions"])
        cls.addin_source = UNIFIED_ADDIN.read_text(encoding="utf-8")
        cls.reference_source = POWERPOINT_REFERENCE.read_text(encoding="utf-8")

    def test_core_and_advanced_groups_partition_every_powerpoint_action(self) -> None:
        self.assertEqual(self.core_actions & self.advanced_actions, set())
        self.assertEqual(
            self.core_actions | self.advanced_actions, self.powerpoint_actions
        )
        self.assertEqual(self.advanced_group.get("reference"), "powerpoint.md")

    def test_every_advanced_action_has_a_contract_and_addin_dispatch(self) -> None:
        dispatch_actions = set(
            re.findall(
                r"^\s*case\s+['\"]([A-Za-z][A-Za-z0-9]*)['\"]\s*:",
                self.addin_source,
                re.MULTILINE,
            )
        )

        self.assertEqual(self.advanced_actions - dispatch_actions, set())
        for action_name in sorted(self.advanced_actions):
            with self.subTest(action=action_name):
                action = self.actions[action_name]
                self.assertEqual(action["application"], "powerpoint")
                self.assertEqual(action["parameters"]["type"], "object")
                self.assertEqual(action["result"]["type"], "object")
                self.assertFalse(action["parameters"]["additionalProperties"])
                self.assertFalse(action["result"]["additionalProperties"])
                self.assertIn(action["risk"], {"read", "write", "destructive"})

    def test_advanced_reference_catalog_is_complete(self) -> None:
        start = self.reference_source.index("<!-- powerpoint-advanced-actions:start -->")
        end = self.reference_source.index("<!-- powerpoint-advanced-actions:end -->")
        documented = set(
            re.findall(r"`([a-z][A-Za-z0-9]*)`", self.reference_source[start:end])
        )

        self.assertEqual(documented, self.advanced_actions)

    def test_external_slide_and_image_actions_have_migration_dispatch(self) -> None:
        dispatch_actions = set(
            re.findall(
                r"^\s*case\s+['\"]([A-Za-z][A-Za-z0-9]*)['\"]\s*:",
                self.addin_source,
                re.MULTILINE,
            )
        )
        self.assertTrue(
            {
                "insertSlidesFromFile",
                "replacePptImage",
                "setFontColor",
                "setShapeFill",
                "setSlideTheme",
            }.issubset(dispatch_actions)
        )

    def test_external_slide_result_uses_numeric_insert_count(self) -> None:
        self.assertEqual(
            self.actions["insertSlidesFromFile"]["result"]["properties"][
                "inserted"
            ]["type"],
            "number",
        )
        self.assertEqual(
            set(self.actions["replacePptImage"]["result"]["required"]),
            {"name", "left", "top", "width", "height", "path"},
        )
        self.assertEqual(
            self.actions["replacePptImage"]["parameters"]["anyOf"],
            [{"required": ["name"]}, {"required": ["shapeIndex"]}],
        )

    def test_advanced_destructive_actions_are_gated(self) -> None:
        for action_name in (
            "removeAnimation",
            "removePptHyperlink",
            "removeSlideTransition",
            "replacePptImage",
            "replacePptText",
        ):
            with self.subTest(action=action_name):
                self.assertEqual(self.actions[action_name]["risk"], "destructive")


if __name__ == "__main__":
    unittest.main()
