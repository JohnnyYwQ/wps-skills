#!/usr/bin/env python3
"""Repository-only three-application native persistence acceptance."""
from pathlib import Path
import sys
ROOT=Path(__file__).resolve().parents[2]
sys.path[:0]=[str(ROOT/'src/main/python'),str(ROOT/'src/test/python')]
from tests.persistence.live_acceptance import main
if __name__=='__main__': raise SystemExit(main())
