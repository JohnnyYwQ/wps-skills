#!/usr/bin/env python3
"""Repository entry point for Task submission, receipts and Action discovery."""

from pathlib import Path
import sys


MAIN_PYTHON = Path(__file__).resolve().parents[1] / "src" / "main" / "python"
if str(MAIN_PYTHON) not in sys.path:
    sys.path.insert(0, str(MAIN_PYTHON))

from wps_skills.cli.call import main  # noqa: E402


if __name__ == "__main__":
    raise SystemExit(main())
