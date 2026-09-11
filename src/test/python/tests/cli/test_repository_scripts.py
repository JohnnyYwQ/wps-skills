import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


class RepositoryEntryTests(unittest.TestCase):
    def test_python_entries_resolve_from_outside_the_repository(self):
        root = Path(__file__).resolve().parents[5]
        scripts = root / 'scripts'
        entries = [scripts / 'call.py']
        for group in ('build', 'demo', 'validate'):
            for app in ('word', 'excel', 'ppt'):
                entries.append(scripts / group / (app + '.py'))
        entries.append(scripts / 'validate' / 'persistence.py')
        entries.extend(scripts / 'validate' / (app + '_common.py') for app in ('excel', 'ppt'))
        with tempfile.TemporaryDirectory() as cwd:
            for entry in entries:
                with self.subTest(entry=entry.relative_to(root)):
                    result = subprocess.run([sys.executable, str(entry), '--help'], cwd=cwd,
                                            env=dict(os.environ, PYTHONPATH='', PYTHONNOUSERSITE='1', PYTHONUTF8='1'),
                                            capture_output=True, text=True, encoding='utf-8', timeout=15)
                    self.assertEqual(0, result.returncode, result.stderr)
                    self.assertIn('usage:', result.stdout)
