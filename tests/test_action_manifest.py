"""Behavior tests for the WPS Action baseline validator CLI."""

from __future__ import annotations

from contextlib import contextmanager
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
VALIDATOR = REPOSITORY_ROOT / "scripts" / "validate_action_manifest.py"
REPRESENTATIVE_FIXTURES = (
    REPOSITORY_ROOT / "tests" / "fixtures" / "representative-actions.json"
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


if __name__ == "__main__":
    unittest.main()
