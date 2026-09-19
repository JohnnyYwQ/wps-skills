#!/usr/bin/env python3
"""Run isolated startup, communication and COM diagnostics from the repository."""
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "src/test/python/diagnostics/startup_communication"))
from debug import main

if __name__ == "__main__":
    raise SystemExit(main())
