#!/usr/bin/env python3
"""Build the complete Windows Action acceptance kit."""
import argparse
from pathlib import Path
import sys
ROOT=Path(__file__).resolve().parents[2]
sys.path[:0]=[str(ROOT/'src/main/python'),str(ROOT/'src/test/python')]
from tests.applications.build_acceptance import build
if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--output',type=Path,required=True)
    print(build(p.parse_args().output))
