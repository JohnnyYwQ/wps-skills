#!/usr/bin/env python3
"""Check the local Windows environment; optionally run real WPS acceptance."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src/test/python/diagnostics/startup_communication'))
from doctor import main
if __name__ == '__main__':
    raise SystemExit(main())
