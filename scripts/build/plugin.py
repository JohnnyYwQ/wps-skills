#!/usr/bin/env python3
"""Build the installable WPS Skills plugin for Codex and Claude Code."""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src/main/python"))
from wps_skills.cli.build_plugin import main

if __name__ == "__main__":
    raise SystemExit(main())
