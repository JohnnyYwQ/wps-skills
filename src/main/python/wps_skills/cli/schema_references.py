"""Build lossless, shared-schema references for application planning interfaces."""

from collections import Counter
import json
from pathlib import Path


def _key(value):
    return json.dumps(value, ensure_ascii=True, sort_keys=True, separators=(",", ":"))


def _nodes(value):
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from _nodes(child)
    elif isinstance(value, list):
        for child in value:
            yield from _nodes(child)


def write_schema_references(skill_directory, contracts, *, schemas=None):
    if schemas is None:
        from wps_skills.word.contracts import schemas

    # The Agent view keeps complete input/output schemas and examples. Runtime
    # contracts remain authoritative and are never replaced by this projection.
    actions = {
        name: {field: contract[field] for field in ("parameters", "result", "examples")}
        for name, contract in contracts.items()
    }
    counts = Counter(_key(node) for node in _nodes(actions))
    shared = {}
    names = {}
    for name, value in sorted(vars(schemas).items()):
        if not name.isupper() or not isinstance(value, dict) or not ({"type", "oneOf", "anyOf", "const"} & set(value)):
            continue
        value = json.loads(json.dumps(value))
        key = _key(value)
        if len(key) >= 200 and counts[key] > 1 and key not in names:
            shared[name] = value
            names[key] = name

    def project(value, *, root=False):
        if isinstance(value, dict):
            name = names.get(_key(value))
            if name and not root:
                return {"$ref": "#/$defs/" + name}
            return {key: project(child) for key, child in value.items()}
        if isinstance(value, list):
            return [project(child) for child in value]
        return value

    destination = Path(skill_directory) / "references" / "schemas"
    for folder, values in (("actions", actions), ("defs", shared)):
        directory = destination / folder
        directory.mkdir(parents=True, exist_ok=True)
        expected = {name + ".json" for name in values}
        for old in directory.glob("*.json"):
            if old.name not in expected:
                old.unlink()  # This directory contains generated references only.
        for name, value in values.items():
            (directory / (name + ".json")).write_text(
                json.dumps(project(value, root=True), ensure_ascii=True, separators=(",", ":")) + "\n",
                encoding="utf-8",
            )
    return destination
