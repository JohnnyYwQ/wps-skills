"""Exercise the production live XML reader without starting WPS."""
import json
import os
from pathlib import Path
import subprocess
import unittest


@unittest.skipUnless(os.name == "nt", "Requires Windows PowerShell")
class StoryXmlVerificationTest(unittest.TestCase):
    def test_story_relationships_and_visible_text(self):
        src = Path(__file__).resolve().parents[4]
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-File", str(src / "test/resources/wps_skills/word/story_xml_verification.ps1"),
             "-BridgePath", str(src / "main/resources/wps_skills/word/windows/word_bridge.ps1")],
            capture_output=True, timeout=30, creationflags=subprocess.CREATE_NO_WINDOW,
        )
        self.assertEqual(result.returncode, 0, result.stdout.decode("utf-8", errors="replace")
                         + result.stderr.decode("utf-8", errors="replace"))
        report = json.loads(result.stdout)
        self.assertGreaterEqual(report["cases"], 16)
        self.assertEqual(report["failures"], 0)
