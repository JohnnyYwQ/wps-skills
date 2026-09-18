"""Verify the standalone Word distribution without importing its runtime."""
import hashlib
import json
from pathlib import Path
import sys


def main():
    root = Path(__file__).resolve().parent / "wps-word"
    manifest = root / "runtime" / "files.sha256.json"
    try:
        entries = json.loads(manifest.read_text(encoding="utf-8"))
        failures = []
        for relative, expected in entries.items():
            path = (root / relative).resolve()
            if root.resolve() not in path.parents:
                failures.append(relative + ": invalid path")
            elif not path.is_file():
                failures.append(relative + ": missing")
            elif hashlib.sha256(path.read_bytes()).hexdigest() != expected:
                failures.append(relative + ": hash mismatch")
        if failures:
            print("\n".join(failures), file=sys.stderr)
            return 1
        print("Verified {} Word Skill files.".format(len(entries)))
        return 0
    except (OSError, ValueError, TypeError, AttributeError) as exc:
        print("Package verification failed: {}".format(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
