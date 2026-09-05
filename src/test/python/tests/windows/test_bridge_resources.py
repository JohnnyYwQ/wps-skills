from pathlib import Path
import unittest


class SharedPowerShellResourceTests(unittest.TestCase):
    def test_bridge_emits_ascii_safe_json_across_windows_code_pages(self):
        bridge_source = (Path(__file__).resolve().parents[4] / "main/resources/wps_skills/windows/bridge_common.ps1").read_text(encoding="utf-8")

        self.assertIn("$json.ToCharArray()", bridge_source)
        self.assertIn("$ascii.AppendFormat('\\u{0:x4}'", bridge_source)
        self.assertIn(
            "[Console]::Out.WriteLine($ascii.ToString())",
            bridge_source,
        )

