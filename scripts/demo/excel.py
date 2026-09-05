#!/usr/bin/env python3
"""Run the visible Excel desktop demonstration from a repository checkout."""
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'src/main/python'))
from wps_skills.excel.demo import main

if __name__ == '__main__':
    raise SystemExit(main())
