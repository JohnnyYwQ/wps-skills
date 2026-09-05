import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from wps_skills.cli.build_ppt_skill import build_ppt_skill
from wps_skills.ppt import contracts


class PptSkillBuildTests(unittest.TestCase):
    def test_relocated_distribution_is_self_contained_and_has_no_type_library_dependency(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            built = build_ppt_skill(root / 'built')
            installed = root / 'installed'
            built.rename(installed)
            entry = installed / 'scripts/ppt.py'
            environment = dict(os.environ, PYTHONPATH='', PYTHONNOUSERSITE='1')
            result = subprocess.run([sys.executable, str(entry), '--app', 'ppt', '--index'],
                                    cwd=root, env=environment, capture_output=True, text=True)
            production = contracts.PPT_PRODUCTION_CONTRACT_SET
            if production is None:
                self.assertEqual(4, result.returncode)
                self.assertEqual('', result.stdout)
                self.assertIn('WPS_DISCOVERY_UNAVAILABLE', result.stderr)
            else:
                self.assertEqual(0, result.returncode, result.stderr)
                self.assertEqual([e.to_wire() for e in production.action_index()], json.loads(result.stdout)['actions'])
            code = "import sys;sys.path.insert(0,sys.argv[1]);import ppt;from wps_skills.ppt.windows.session import BRIDGE_SCRIPT;assert BRIDGE_SCRIPT.is_file();assert (BRIDGE_SCRIPT.parents[2]/'windows/bridge_common.ps1').is_file();assert not ppt.open_session().can_execute;assert 'win32com' not in sys.modules"
            subprocess.run([sys.executable, '-c', code, str(entry.parent)], cwd=root, env=environment, check=True)
            manifest = json.loads((installed / 'runtime/files.sha256.json').read_text())
            for name, digest in manifest.items():
                self.assertEqual(digest, hashlib.sha256((installed / name).read_bytes()).hexdigest())
            self.assertFalse((installed / 'runtime/src/test').exists())
            self.assertFalse(any(installed.rglob('wps_ppt_api.py')))

    def test_does_not_overwrite_an_existing_install(self):
        with tempfile.TemporaryDirectory() as temporary:
            marker = Path(temporary) / 'user.txt'
            marker.write_text('keep')
            with self.assertRaises(FileExistsError):
                build_ppt_skill(temporary)
            self.assertEqual('keep', marker.read_text())
