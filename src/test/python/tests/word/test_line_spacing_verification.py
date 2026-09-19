"""Exercise production PowerShell read-back verification with native-shaped facts."""

import json
import os
from pathlib import Path
import subprocess
import unittest


@unittest.skipUnless(os.name == "nt", "Requires Windows PowerShell")
class LineSpacingVerificationTest(unittest.TestCase):
    def test_equivalent_rules_pass_and_different_effects_fail(self):
        src = Path(__file__).resolve().parents[4]
        script = src / "test/resources/wps_skills/word/line_spacing_verification.ps1"
        bridge = src / "main/resources/wps_skills/word/windows/word_bridge.ps1"
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy",
             "Bypass", "-File", str(script), "-BridgePath", str(bridge)],
            capture_output=True, timeout=30, creationflags=subprocess.CREATE_NO_WINDOW,
        )
        self.assertEqual(result.returncode, 0, result.stdout.decode("utf-8", errors="replace")
                         + result.stderr.decode("utf-8", errors="replace"))
        report = json.loads(result.stdout)
        self.assertGreaterEqual(report["cases"], 24)
        self.assertEqual(report["failures"], 0)


if __name__ == "__main__":
    unittest.main()
