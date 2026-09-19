#!/usr/bin/env python3
"""Return complete selected Action schemas and their shared definitions; no WPS."""

import argparse
import json
from pathlib import Path
import re
import sys


REFERENCE_ROOT = Path(__file__).resolve().parents[1] / "references" / "schemas"
NAME = re.compile(r"[A-Za-z][A-Za-z0-9_]*\Z")


def references(value):
    if isinstance(value, dict):
        if "$ref" in value:
            ref = value["$ref"]
            if not isinstance(ref, str) or not ref.startswith("#/$defs/"):
                raise ValueError("Unsupported schema reference: " + repr(ref))
            name = ref[len("#/$defs/"):]
            if not NAME.fullmatch(name):
                raise ValueError("Invalid definition name: " + name)
            yield name
        for child in value.values():
            yield from references(child)
    elif isinstance(value, list):
        for child in value:
            yield from references(child)


def query(names, root=REFERENCE_ROOT):
    names = list(dict.fromkeys(names))
    if not names or any(not NAME.fullmatch(name) for name in names):
        raise ValueError("Provide one or more Action names, not file paths")

    def read(folder, name):
        path = root / folder / (name + ".json")
        if not path.is_file():
            raise ValueError("Unknown Action: " + name if folder == "actions"
                             else "Missing schema definition: " + name)
        return json.loads(path.read_text(encoding="utf-8"))

    actions = {name: read("actions", name) for name in names}
    definitions = {}
    visiting = set()

    def include(name):
        if name in visiting:
            raise ValueError("Cyclic schema definition: " + name)
        if name in definitions:
            return
        visiting.add(name)
        value = read("defs", name)
        for dependency in references(value):
            include(dependency)
        visiting.remove(name)
        definitions[name] = value

    for name in references(actions):
        include(name)
    return {"actions": actions, "$defs": definitions}


def render(value, width=160):
    """Wrap only between JSON tokens, preserving strings and all constraints."""
    lines = []
    line = ""
    encoder = json.JSONEncoder(ensure_ascii=True, separators=(",", ":"))
    for token in encoder.iterencode(value):
        if line and len(line) + len(token) > width:
            lines.append(line)
            line = ""
        line += token
    if line:
        lines.append(line)
    return "\n".join(lines) + "\n"


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("actions", nargs="+", help="Action names; batch related Actions in one call")
    args = parser.parse_args(argv)
    try:
        output = render(query(args.actions))
    except (ValueError, OSError) as exc:
        parser.exit(2, str(exc) + "\n")
    # ASCII JSON survives Windows shell encoding without changing string values.
    sys.stdout.write(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
