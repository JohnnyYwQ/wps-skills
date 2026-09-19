#!/usr/bin/env python3
"""Run word Action acceptance on an interactive Windows desktop."""
from pathlib import Path
import sys
ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'src/test/python'))
from tests.applications.action_acceptance import main
if __name__ == '__main__':
    raise SystemExit(main('word'))
