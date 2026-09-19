"""Windows font catalog and production readback assertions, without a WPS process."""
import os
from pathlib import Path
import subprocess
import unittest

from wps_skills.word.windows.task_factory import BRIDGE_SCRIPT


@unittest.skipUnless(os.name == 'nt', 'Requires Windows installed font catalog')
class WordFontNameTests(unittest.TestCase):
    def test_installed_aliases_and_strict_mismatch_verification(self):
        root = Path(__file__).resolve().parents[6]
        script = root / 'src/test/resources/wps_skills/word/font-aliases/verify.ps1'
        result = subprocess.run(
            ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass',
             '-File', str(script), '-BridgePath', str(BRIDGE_SCRIPT)],
            capture_output=True, text=True, errors='replace', timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn('"passed":13', result.stdout)
