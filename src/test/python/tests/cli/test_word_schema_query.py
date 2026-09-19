import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from wps_skills.cli.build_skill import MAIN
from wps_skills.cli.build_word_skill import build_word_skill
from wps_skills.word.actions.registry import WORD_ACTIONS


ENTRY = MAIN / "resources/skills/wps-word/scripts/schema.py"
spec = importlib.util.spec_from_file_location("word_schema_query", ENTRY)
query = importlib.util.module_from_spec(spec)
spec.loader.exec_module(query)


def expand(value, definitions):
    if isinstance(value, dict):
        if "$ref" in value:
            return expand(definitions[value["$ref"].split("/")[-1]], definitions)
        return {key: expand(child, definitions) for key, child in value.items()}
    if isinstance(value, list):
        return [expand(child, definitions) for child in value]
    return value


class WordSchemaQueryTests(unittest.TestCase):
    def test_all_action_schemas_are_lossless_and_references_are_closed(self):
        batch = query.query(list(WORD_ACTIONS))
        for name, action in WORD_ACTIONS.items():
            wire = action.contract.to_wire("word")
            expected = {key: wire[key] for key in ("parameters", "result", "examples")}
            single = query.query([name])
            self.assertEqual(expected, expand(single["actions"][name], single["$defs"]))
            self.assertEqual(expected, expand(batch["actions"][name], batch["$defs"]))
        encoded = query.render(batch)
        self.assertEqual(batch, json.loads(encoded))
        self.assertLess(max(map(len, encoded.splitlines())), 2000)

    def test_query_reads_only_selected_actions_and_transitive_dependencies_once(self):
        reads = []
        original = Path.read_text

        def record(path, *args, **kwargs):
            reads.append(path)
            return original(path, *args, **kwargs)

        with patch.object(Path, "read_text", record):
            result = query.query(["replaceContent", "save", "replaceContent"])
        self.assertEqual(["replaceContent", "save"], list(result["actions"]))
        expected = {query.REFERENCE_ROOT / "actions" / (name + ".json") for name in result["actions"]}
        expected.update(query.REFERENCE_ROOT / "defs" / (name + ".json") for name in result["$defs"])
        self.assertEqual(expected, set(reads))
        self.assertEqual(len(reads), len(set(reads)))

    def test_invalid_queries_emit_no_partial_stdout(self):
        for names in (["save", "typo"], ["../actions"], []):
            with self.subTest(names=names):
                result = subprocess.run([sys.executable, str(ENTRY), *names], capture_output=True, text=True)
                self.assertEqual(2, result.returncode)
                self.assertEqual("", result.stdout)
                self.assertTrue(result.stderr)

    def test_missing_or_cyclic_definitions_fail_without_incomplete_output(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / "actions").mkdir()
            (root / "defs").mkdir()
            (root / "actions/save.json").write_text('{"parameters":{"$ref":"#/$defs/A"}}')
            with self.assertRaisesRegex(ValueError, "Missing"):
                query.query(["save"], root)
            (root / "defs/A.json").write_text('{"$ref":"#/$defs/B"}')
            (root / "defs/B.json").write_text('{"$ref":"#/$defs/A"}')
            with self.assertRaisesRegex(ValueError, "Cyclic"):
                query.query(["save"], root)

    def test_relocated_query_needs_no_runtime_or_combined_contract_file(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            built = build_word_skill(root / "built")
            installed = root / "moved"
            built.rename(installed)
            self.assertFalse((installed / "references/actions.json").exists())
            result = subprocess.run(
                [sys.executable, "-I", str(installed / "scripts/schema.py"),
                 "openDocument", "inspectDocument", "findContent", "replaceContent", "save"],
                cwd=root, capture_output=True, text=True, check=True,
            )
            value = json.loads(result.stdout)
            self.assertEqual(5, len(value["actions"]))
            self.assertEqual("", result.stderr)
            self.assertLess(len(result.stdout.encode()), 20000)
            self.assertLess(max(map(len, result.stdout.splitlines())), 2000)


if __name__ == "__main__":
    unittest.main()
