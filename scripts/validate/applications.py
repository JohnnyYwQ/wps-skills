#!/usr/bin/env python3
"""Opt-in Windows native acceptance for current application Tasks."""
from pathlib import Path
import sys
root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(root / 'src/test/python'))
from tests.applications.native_acceptance import run
if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, required=True)
    raise SystemExit(run(parser.parse_args().root))
