#!/usr/bin/env python3
"""Run repository-only common Excel Action acceptance on Windows WPS."""
from pathlib import Path
import runpy
import sys

repo = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(repo / 'src/main/python'))
runpy.run_path(str(repo / 'src/test/python/tests/excel/common_live_acceptance.py'), run_name='__main__')
