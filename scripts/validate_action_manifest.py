#!/usr/bin/env python3
"""Validate the checked-in WPS Action baseline.

Input: A repository root containing the production manifest, packaged Add-in, and frozen migration ledger.
Output: A human-readable validation summary or actionable errors on stderr.
Position: Standard-library guard for the migration's canonical WPS Action contract.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
from typing import Any, Dict


MANIFEST_PATH = Path("skills/wps-office/references/action-manifest.json")
MIGRATION_MAP_PATH = Path("doc/migration/legacy-tool-action-map.json")
JAVASCRIPT_DISPATCH_PATH = Path("skills/wps-office/assets/wps-addin/main.js")
ACTION_FIELDS = {
    "action",
    "application",
    "description",
    "parameters",
    "result",
    "prerequisites",
    "risk",
    "examples",
}
REFERENCE_GROUP_FIELDS = {"actions", "reference"}
REQUIRED_ACTION_FIELDS = ACTION_FIELDS - {"examples"}
MAPPING_FIELDS = {"tool", "actions", "status", "reason", "decision"}
MAPPING_STATUSES = {"mapped", "workflow", "retired"}
RETIRED_ACTION_FIELDS = {"action", "decision", "reason", "replacements"}
RETIRED_CONTRACT_FIELDS = {"action", "decision", "reason", "retired_parameters"}
REVIEWED_RETIREMENT_DECISION = "ADR-0014"
ADR_0014_RETIRED_ACTIONS = {"convertFormat", "openFile"}
ADR_0014_RETIRED_LEGACY_TOOLS = {
    "wps_cache_data",
    "wps_clear_cache",
    "wps_convert_format",
    "wps_excel_find_replace",
    "wps_execute_method",
    "wps_get_cached_data",
    "wps_list_cache",
    "wps_word_proofread_basic",
}
ADR_0014_RETIRED_CONTRACTS = {
    "addArrow": {"height", "left", "top", "width"},
}
FROZEN_LEGACY_TOOL_COUNT = 250
FROZEN_LEGACY_TOOL_IDENTIFIERS_DIGEST = (
    "d64663b94f87377bd8b0636b8003a01b6252237195a1a587d553ce6d1eebac4f"
)
JSON_TYPES = {"array", "boolean", "null", "number", "object", "string"}
SCHEMA_FIELDS = {
    "additionalProperties",
    "anyOf",
    "description",
    "enum",
    "items",
    "maximum",
    "minimum",
    "pattern",
    "properties",
    "required",
    "type",
}


def load_json(root: Path, relative_path: Path) -> Dict[str, Any]:
    path = root / relative_path
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as error:
        raise ValueError(f"missing required file: {relative_path}") from error
    except json.JSONDecodeError as error:
        raise ValueError(f"invalid JSON in {relative_path}: {error}") from error
    if not isinstance(value, dict):
        raise ValueError(f"{relative_path} must contain a JSON object")
    return value


def matches_json_type(value: Any, expected_type: str) -> bool:
    type_checks = {
        "array": lambda item: isinstance(item, list),
        "boolean": lambda item: isinstance(item, bool),
        "null": lambda item: item is None,
        "number": lambda item: isinstance(item, (int, float)) and not isinstance(item, bool),
        "object": lambda item: isinstance(item, dict),
        "string": lambda item: isinstance(item, str),
    }
    check = type_checks.get(expected_type)
    return bool(check and check(value))


def valid_declared_type(value: Any) -> bool:
    if isinstance(value, str):
        return value in JSON_TYPES
    return (
        isinstance(value, list)
        and bool(value)
        and len(value) == len(set(value))
        and all(isinstance(item, str) and item in JSON_TYPES for item in value)
    )


def validate_schema_node(schema: Any, context: str) -> None:
    if not isinstance(schema, dict):
        raise ValueError(f"{context} must be a schema object")
    unknown_fields = set(schema) - SCHEMA_FIELDS
    if unknown_fields:
        field = sorted(unknown_fields)[0]
        raise ValueError(f"{context} has unknown schema field '{field}'")
    if not valid_declared_type(schema.get("type")):
        raise ValueError(f"{context} requires a valid JSON type")
    if "enum" in schema:
        enum = schema["enum"]
        if not isinstance(enum, list) or not enum:
            raise ValueError(f"{context}.enum must be a non-empty array")
        if any(
            not any(matches_json_type(value, expected_type) for expected_type in (
                [schema["type"]] if isinstance(schema["type"], str) else schema["type"]
            ))
            for value in enum
        ):
            raise ValueError(f"{context}.enum contains a value outside its declared type")
    if "pattern" in schema:
        if schema.get("type") != "string" or not isinstance(schema["pattern"], str):
            raise ValueError(f"{context}.pattern requires a string schema")
        try:
            re.compile(schema["pattern"])
        except re.error as error:
            raise ValueError(f"{context}.pattern must be a valid regular expression") from error
    for bound_name in ("minimum", "maximum"):
        if bound_name not in schema:
            continue
        bound = schema[bound_name]
        declared_types = schema.get("type")
        if (
            "number"
            not in (
                [declared_types]
                if isinstance(declared_types, str)
                else declared_types or []
            )
            or not matches_json_type(bound, "number")
        ):
            raise ValueError(f"{context}.{bound_name} requires a number schema")
    if (
        "minimum" in schema
        and "maximum" in schema
        and schema["minimum"] > schema["maximum"]
    ):
        raise ValueError(f"{context}.minimum must not exceed maximum")
    if "properties" in schema:
        properties = schema["properties"]
        if not isinstance(properties, dict):
            raise ValueError(f"{context}.properties must be an object")
        for property_name, property_schema in properties.items():
            validate_schema_node(property_schema, f"{context}.{property_name}")
    if "required" in schema:
        required = schema["required"]
        if not isinstance(required, list) or not all(
            isinstance(item, str) for item in required
        ):
            raise ValueError(f"{context}.required must be an array of property names")
        properties = schema.get("properties")
        if isinstance(properties, dict):
            unknown_required = set(required) - set(properties)
            if unknown_required:
                property_name = sorted(unknown_required)[0]
                raise ValueError(f"{context} requires unknown property '{property_name}'")
    if "items" in schema:
        validate_schema_node(schema["items"], f"{context}.items")
    if "anyOf" in schema:
        any_of = schema["anyOf"]
        if not isinstance(any_of, list) or not any_of:
            raise ValueError(f"{context}.anyOf must be a non-empty array")
        properties = schema.get("properties", {})
        for branch_index, branch in enumerate(any_of):
            if not isinstance(branch, dict) or not branch:
                raise ValueError(
                    f"{context}.anyOf[{branch_index}] must be a schema object"
                )
            if set(branch) != {"required"}:
                validate_schema_node(branch, f"{context}.anyOf[{branch_index}]")
                continue
            required = branch["required"]
            if not isinstance(required, list) or not required or not all(
                isinstance(item, str) for item in required
            ):
                raise ValueError(
                    f"{context}.anyOf[{branch_index}].required must be a non-empty array"
                )
            unknown_required = set(required) - set(properties)
            if unknown_required:
                property_name = sorted(unknown_required)[0]
                raise ValueError(
                    f"{context}.anyOf[{branch_index}] requires unknown property "
                    f"'{property_name}'"
                )


def validate_nested_contract(
    schema: Dict[str, Any], context: str, contract_label: str
) -> None:
    declared_types = schema.get("type")
    if not valid_declared_type(declared_types):
        return
    if isinstance(declared_types, str):
        declared_types = [declared_types]
    if "array" in declared_types:
        if not isinstance(schema.get("items"), dict):
            raise ValueError(
                f"{context}: array {contract_label} requires an items schema"
            )
        validate_nested_contract(
            schema["items"], f"{context}.items", contract_label
        )
    if "object" in declared_types:
        if not isinstance(schema.get("properties"), dict):
            raise ValueError(
                f"{context}: object {contract_label} requires a properties schema"
            )
        for property_name, property_schema in schema["properties"].items():
            validate_nested_contract(
                property_schema,
                f"{context}.{property_name}",
                contract_label,
            )


def validate_contract_schema(schema: Any, context: str) -> None:
    if not isinstance(schema, dict) or schema.get("type") != "object":
        raise ValueError(f"{context} must be an object schema")
    properties = schema.get("properties")
    if not isinstance(properties, dict):
        raise ValueError(f"{context}.properties must be an object")
    required = schema.get("required")
    if not isinstance(required, list) or not all(
        isinstance(item, str) for item in required
    ):
        raise ValueError(f"{context}.required must be an array of property names")
    unknown_required = set(required) - set(properties)
    if unknown_required:
        property_name = sorted(unknown_required)[0]
        raise ValueError(f"{context} requires unknown property '{property_name}'")
    if schema.get("additionalProperties") is not False:
        raise ValueError(f"{context}.additionalProperties must be false")
    label = "parameter" if context.endswith("parameters") else "result field"
    for property_name, property_schema in properties.items():
        if not isinstance(property_schema, dict) or not valid_declared_type(
            property_schema.get("type")
        ):
            raise ValueError(
                f"{context} {label} '{property_name}' requires a valid JSON type"
            )
        validate_nested_contract(
            property_schema,
            f"{context}.{property_name}",
            label,
        )
    validate_schema_node(schema, context)


def validate_example_value(value: Any, schema: Dict[str, Any], context: str) -> None:
    expected_types = schema.get("type")
    if isinstance(expected_types, str):
        expected_types = [expected_types]
    if expected_types and not any(matches_json_type(value, item) for item in expected_types):
        raise ValueError(f"{context}: value does not match declared type")
    if "enum" in schema and value not in schema["enum"]:
        raise ValueError(f"{context}: value is not one of the declared enum values")
    if (
        matches_json_type(value, "number")
        and "minimum" in schema
        and value < schema["minimum"]
    ):
        raise ValueError(f"{context}: value is below the declared minimum")
    if (
        matches_json_type(value, "number")
        and "maximum" in schema
        and value > schema["maximum"]
    ):
        raise ValueError(f"{context}: value exceeds the declared maximum")
    if (
        isinstance(value, str)
        and "pattern" in schema
        and re.search(schema["pattern"], value) is None
    ):
        raise ValueError(f"{context}: value does not match declared pattern")
    if isinstance(value, dict):
        properties = schema.get("properties", {})
        for required_property in schema.get("required", []):
            if required_property not in value:
                raise ValueError(
                    f"{context}: missing required property '{required_property}'"
                )
        if schema.get("additionalProperties") is False:
            unknown_properties = set(value) - set(properties)
            if unknown_properties:
                property_name = sorted(unknown_properties)[0]
                raise ValueError(f"{context}: unknown property '{property_name}'")
        for property_name, property_value in value.items():
            property_schema = properties.get(property_name)
            if isinstance(property_schema, dict):
                validate_example_value(
                    property_value,
                    property_schema,
                    f"{context}.{property_name}",
                )
        any_of = schema.get("anyOf", [])
        if any_of and not any(example_matches_schema(value, branch) for branch in any_of):
            raise ValueError(f"{context}: must satisfy at least one anyOf branch")
    if isinstance(value, list) and isinstance(schema.get("items"), dict):
        for item_index, item in enumerate(value):
            validate_example_value(item, schema["items"], f"{context}[{item_index}]")


def example_matches_schema(value: Any, schema: Dict[str, Any]) -> bool:
    try:
        validate_example_value(value, schema, "anyOf branch")
    except ValueError:
        return False
    return True


def legacy_tool_application(tool_name: str) -> str:
    prefixes = {
        "wps_excel_": "excel",
        "wps_word_": "word",
        "wps_ppt_": "powerpoint",
    }
    for prefix, application in prefixes.items():
        if tool_name.startswith(prefix):
            return application
    return ""


def validate_retired_actions(
    migration_map: Dict[str, Any], action_names: set
) -> set:
    retired_actions = migration_map.get("retired_actions")
    if not isinstance(retired_actions, list):
        raise ValueError("legacy-tool-action-map.json field 'retired_actions' must be an array")
    retired_names = set()
    for index, retirement in enumerate(retired_actions):
        if not isinstance(retirement, dict) or set(retirement) != RETIRED_ACTION_FIELDS:
            raise ValueError(
                f"retired_actions[{index}] must contain action, decision, reason, and replacements"
            )
        action_name = retirement["action"]
        if not isinstance(action_name, str) or not re.fullmatch(
            r"[a-z][A-Za-z0-9]*", action_name
        ):
            raise ValueError(f"retired_actions[{index}].action must be a WPS Action name")
        if action_name in retired_names:
            raise ValueError(f"duplicate retired WPS Action '{action_name}'")
        retired_names.add(action_name)
        if retirement["decision"] != REVIEWED_RETIREMENT_DECISION:
            raise ValueError(
                f"retired WPS Action '{action_name}' requires a reviewed decision "
                f"({REVIEWED_RETIREMENT_DECISION})"
            )
        if not isinstance(retirement["reason"], str) or not retirement["reason"].strip():
            raise ValueError(f"retired WPS Action '{action_name}' requires a reason")
        replacements = retirement["replacements"]
        if not isinstance(replacements, list) or not all(
            isinstance(replacement, str) and replacement
            for replacement in replacements
        ) or len(replacements) != len(set(replacements)):
            raise ValueError(
                f"retired WPS Action '{action_name}'.replacements must be a unique Action array"
            )
        unknown_replacements = set(replacements) - action_names
        if unknown_replacements:
            replacement = sorted(unknown_replacements)[0]
            raise ValueError(
                f"retired WPS Action '{action_name}' references unknown replacement "
                f"'{replacement}'"
            )
    exposed_actions = retired_names & action_names
    if exposed_actions:
        action_name = sorted(exposed_actions)[0]
        raise ValueError(f"retired WPS Action '{action_name}' is still exposed by the manifest")
    return retired_names


def validate_retired_contracts(
    migration_map: Dict[str, Any], actions_by_name: Dict[str, Dict[str, Any]]
) -> Dict[str, set]:
    retired_contracts = migration_map.get("retired_contracts")
    if not isinstance(retired_contracts, list):
        raise ValueError("legacy-tool-action-map.json field 'retired_contracts' must be an array")
    retired_contracts_by_action = {}
    for index, retirement in enumerate(retired_contracts):
        if not isinstance(retirement, dict) or set(retirement) != RETIRED_CONTRACT_FIELDS:
            raise ValueError(
                f"retired_contracts[{index}] must contain action, decision, reason, and retired_parameters"
            )
        action_name = retirement["action"]
        if action_name not in actions_by_name:
            raise ValueError(
                f"retired contract references unknown WPS Action '{action_name}'"
            )
        if action_name in retired_contracts_by_action:
            raise ValueError(f"duplicate retired contract for WPS Action '{action_name}'")
        if retirement["decision"] != REVIEWED_RETIREMENT_DECISION:
            raise ValueError(
                f"retired contract for WPS Action '{action_name}' requires a reviewed "
                f"decision ({REVIEWED_RETIREMENT_DECISION})"
            )
        if not isinstance(retirement["reason"], str) or not retirement["reason"].strip():
            raise ValueError(
                f"retired contract for WPS Action '{action_name}' requires a reason"
            )
        parameters = retirement["retired_parameters"]
        if not isinstance(parameters, list) or not parameters or not all(
            isinstance(parameter, str) and parameter for parameter in parameters
        ) or len(parameters) != len(set(parameters)):
            raise ValueError(
                f"retired contract for WPS Action '{action_name}'.retired_parameters "
                "must be a non-empty unique parameter array"
            )
        exposed_parameters = set(parameters) & set(
            actions_by_name[action_name]["parameters"]["properties"]
        )
        if exposed_parameters:
            parameter = sorted(exposed_parameters)[0]
            raise ValueError(
                f"retired parameter '{parameter}' is still exposed by WPS Action "
                f"'{action_name}'"
            )
        retired_contracts_by_action[action_name] = set(parameters)
    return retired_contracts_by_action


def validate_reviewed_retirement_boundary(
    migration_map: Dict[str, Any],
    retired_actions: set,
    retired_legacy_tools: set,
    retired_contracts: Dict[str, set],
) -> None:
    schema_version = migration_map.get("schema_version")
    if not isinstance(schema_version, int) or schema_version < 1:
        raise ValueError("legacy-tool-action-map.json requires a positive schema_version")
    if schema_version < 2:
        if retired_actions or retired_legacy_tools or retired_contracts:
            raise ValueError(
                "retirement entries require legacy-tool-action-map.json schema_version 2"
            )
        return
    if retired_actions != ADR_0014_RETIRED_ACTIONS:
        raise ValueError("retired WPS Actions are not the reviewed ADR-0014 boundary")
    if retired_legacy_tools != ADR_0014_RETIRED_LEGACY_TOOLS:
        raise ValueError("retired legacy WPS Tools are not the reviewed ADR-0014 boundary")
    if retired_contracts != ADR_0014_RETIRED_CONTRACTS:
        raise ValueError("retired WPS Action contracts are not the reviewed ADR-0014 boundary")


def validate_frozen_legacy_tool_inventory(migration_map: Dict[str, Any], tools: set) -> None:
    if migration_map.get("schema_version", 0) < 2:
        return
    identifiers_digest = hashlib.sha256(
        "\n".join(sorted(tools)).encode("utf-8")
    ).hexdigest()
    if (
        len(tools) != FROZEN_LEGACY_TOOL_COUNT
        or identifiers_digest != FROZEN_LEGACY_TOOL_IDENTIFIERS_DIGEST
    ):
        raise ValueError(
            "legacy-tool-action-map.json does not match the frozen WPS Tool inventory"
        )


def validate(root: Path) -> int:
    manifest = load_json(root, MANIFEST_PATH)
    migration_map = load_json(root, MIGRATION_MAP_PATH)
    actions = manifest.get("actions")
    if not isinstance(actions, list):
        raise ValueError("action-manifest.json field 'actions' must be an array")
    action_names = set()
    action_applications = {}
    actions_by_name = {}
    for index, action in enumerate(actions):
        if not isinstance(action, dict):
            raise ValueError(f"actions[{index}] must be an object")
        unknown_fields = set(action) - ACTION_FIELDS
        if unknown_fields:
            field = sorted(unknown_fields)[0]
            raise ValueError(f"actions[{index}] has unknown field '{field}'")
        missing_fields = REQUIRED_ACTION_FIELDS - set(action)
        if missing_fields:
            field = sorted(missing_fields)[0]
            raise ValueError(f"actions[{index}] is missing required field '{field}'")
        if action.get("risk") not in {"read", "write", "destructive"}:
            raise ValueError(
                f"actions[{index}].risk must be one of read, write, destructive"
            )
        if action.get("application") not in {"common", "excel", "word", "powerpoint"}:
            raise ValueError(
                f"actions[{index}].application must be one of common, excel, word, powerpoint"
            )
        action_name = action.get("action")
        if not isinstance(action_name, str) or not re.fullmatch(
            r"[a-z][A-Za-z0-9]*", action_name
        ):
            raise ValueError(
                f"actions[{index}].action must be an original lowerCamelCase identifier"
            )
        if not isinstance(action.get("description"), str) or not action[
            "description"
        ].strip():
            raise ValueError(f"actions[{index}].description must be a non-empty string")
        prerequisites = action.get("prerequisites")
        allowed_prerequisites = {
            "wps_running",
            "active_document",
            "active_workbook",
            "active_presentation",
        }
        if not isinstance(prerequisites, list) or not all(
            isinstance(item, str) and item in allowed_prerequisites
            for item in prerequisites
        ):
            raise ValueError(f"actions[{index}].prerequisites contains an invalid value")
        examples = action.get("examples", [])
        if not isinstance(examples, list) or not all(
            isinstance(example, dict) for example in examples
        ):
            raise ValueError(f"actions[{index}].examples must be an array of objects")
        validate_contract_schema(action["parameters"], f"actions[{index}].parameters")
        validate_contract_schema(action["result"], f"actions[{index}].result")
        if action_name in action_names:
            raise ValueError(f"duplicate WPS Action '{action_name}'")
        action_names.add(action_name)
        action_applications[action_name] = action["application"]
        actions_by_name[action_name] = action
        for example in action.get("examples", []):
            example_name = example.get("name", "unnamed")
            validate_example_value(
                example.get("params"),
                action["parameters"],
                f"example '{example_name}' params",
            )
            validate_example_value(
                example.get("result"),
                action["result"],
                f"example '{example_name}' result",
            )
    retired_action_names = validate_retired_actions(migration_map, action_names)
    retired_contracts = validate_retired_contracts(migration_map, actions_by_name)
    reference_groups = manifest.get("reference_groups", {})
    if not isinstance(reference_groups, dict):
        raise ValueError("action-manifest.json field 'reference_groups' must be an object")
    for group_name, group in reference_groups.items():
        if not isinstance(group_name, str) or not re.fullmatch(r"[a-z][a-z0-9_]*", group_name):
            raise ValueError("reference group names must be lower_snake_case identifiers")
        if not isinstance(group, dict) or set(group) != REFERENCE_GROUP_FIELDS:
            raise ValueError(
                f"reference_groups.{group_name} must contain actions and reference"
            )
        reference = group["reference"]
        if not isinstance(reference, str) or not reference.endswith(".md"):
            raise ValueError(f"reference_groups.{group_name}.reference must be a Markdown file")
        group_actions = group["actions"]
        if (
            not isinstance(group_actions, list)
            or not group_actions
            or not all(isinstance(item, str) and item for item in group_actions)
            or len(group_actions) != len(set(group_actions))
        ):
            raise ValueError(
                f"reference_groups.{group_name}.actions must be a non-empty unique Action array"
            )
        unknown_group_actions = set(group_actions) - action_names
        if unknown_group_actions:
            action_name = sorted(unknown_group_actions)[0]
            raise ValueError(
                f"reference group '{group_name}' references unknown WPS Action '{action_name}'"
            )
    javascript_source = (root / JAVASCRIPT_DISPATCH_PATH).read_text(encoding="utf-8")
    javascript_actions = set(
        re.findall(r"^\s*case\s+['\"]([A-Za-z][A-Za-z0-9]*)['\"]\s*:", javascript_source, re.MULTILINE)
    )
    unmanifested_javascript_actions = javascript_actions - action_names
    if unmanifested_javascript_actions:
        action_name = sorted(unmanifested_javascript_actions)[0]
        raise ValueError(
            f"JavaScript dispatch has unmanifested WPS Action '{action_name}'"
        )
    javascript_omissions = action_names - javascript_actions
    if javascript_omissions:
        action_name = sorted(javascript_omissions)[0]
        raise ValueError(
            f"WPS Action '{action_name}' is missing from the packaged JavaScript Add-in"
        )
    mappings = migration_map.get("legacy_tools", [])
    if not isinstance(mappings, list):
        raise ValueError("legacy-tool-action-map.json field 'legacy_tools' must be an array")
    mapped_tools = set()
    retired_legacy_tools = set()
    for index, mapping in enumerate(mappings):
        if not isinstance(mapping, dict):
            raise ValueError(f"legacy_tools[{index}] must be an object")
        unknown_fields = set(mapping) - MAPPING_FIELDS
        if unknown_fields:
            field = sorted(unknown_fields)[0]
            raise ValueError(f"legacy_tools[{index}] has unknown field '{field}'")
        tool_name = mapping.get("tool")
        if not isinstance(tool_name, str) or not tool_name:
            raise ValueError(f"legacy_tools[{index}].tool must be a non-empty string")
        if tool_name in mapped_tools:
            raise ValueError(f"duplicate legacy WPS Tool mapping '{tool_name}'")
        mapped_tools.add(tool_name)
        status = mapping.get("status")
        if status not in MAPPING_STATUSES:
            raise ValueError(
                f"legacy WPS Tool '{tool_name}' has invalid status '{status}'"
            )
        mapped_actions = mapping.get("actions")
        if not isinstance(mapped_actions, list) or not all(
            isinstance(action, str) and action for action in mapped_actions
        ):
            raise ValueError(
                f"legacy WPS Tool '{tool_name}'.actions must be an array of Action names"
            )
        if status == "mapped" and len(mapped_actions) != 1:
            raise ValueError(
                f"legacy WPS Tool '{tool_name}' status 'mapped' requires exactly one WPS Action"
            )
        if status == "workflow" and not mapped_actions:
            raise ValueError(
                f"legacy WPS Tool '{tool_name}' status 'workflow' requires WPS Actions"
            )
        if status == "retired":
            if mapped_actions:
                raise ValueError(
                    f"legacy WPS Tool '{tool_name}' status 'retired' cannot map WPS Actions"
                )
            if not isinstance(mapping.get("reason"), str) or not mapping["reason"].strip():
                raise ValueError(
                    f"legacy WPS Tool '{tool_name}' status 'retired' requires a reason"
                )
        if status == "retired" and mapping.get("decision") != REVIEWED_RETIREMENT_DECISION:
            raise ValueError(
                f"legacy WPS Tool '{tool_name}' status 'retired' requires a reviewed "
                f"decision ({REVIEWED_RETIREMENT_DECISION})"
            )
        if status == "retired":
            retired_legacy_tools.add(tool_name)
    for mapping in mappings:
        if not isinstance(mapping, dict):
            continue
        for mapped_action in mapping.get("actions", []):
            if mapped_action not in action_names:
                raise ValueError(
                    f"legacy WPS Tool '{mapping.get('tool')}' maps to unknown "
                    f"WPS Action '{mapped_action}'"
                )
            tool_application = legacy_tool_application(mapping["tool"])
            action_application = action_applications[mapped_action]
            if tool_application and tool_application != action_application:
                raise ValueError(
                    f"legacy WPS Tool '{mapping['tool']}' owned by {tool_application} "
                    f"maps to {action_application} WPS Action '{mapped_action}'"
                )
    validate_frozen_legacy_tool_inventory(migration_map, mapped_tools)
    validate_reviewed_retirement_boundary(
        migration_map,
        retired_action_names,
        retired_legacy_tools,
        retired_contracts,
    )
    return len(actions)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        default=Path(__file__).resolve().parents[1],
        help="repository root (defaults to the validator's repository)",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        action_count = validate(args.root.resolve())
    except ValueError as error:
        print(f"Action baseline validation failed: {error}", file=sys.stderr)
        return 1
    print(f"Action baseline valid: {action_count} WPS Actions")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
