"""Verify actual isolated distributions have no cross-application dependencies."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from wps_skills.cli.build_skill import build_application_skill


class ScopedSkillBuildTests(unittest.TestCase):
    def test_each_package_imports_and_discovers_without_other_applications(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            for app, count in (('word', 14), ('excel', 34), ('ppt', 37)):
                with self.subTest(app=app):
                    package = build_application_skill(app, root / app)
                    runtime = package / 'runtime/src/main'
                    python = runtime / 'python/wps_skills'
                    resources = runtime / 'resources/wps_skills'
                    for other in {'word', 'excel', 'ppt'} - {app}:
                        self.assertFalse((python / other).exists())
                        self.assertFalse((resources / other).exists())
                    self.assertFalse(list(package.rglob('demo*')))
                    self.assertFalse(list(python.rglob('build_*.py')))
                    env = dict(os.environ, PYTHONPATH='', PYTHONNOUSERSITE='1', PYTHONDONTWRITEBYTECODE='1')
                    # Import every shipped module from a relocated, isolated path.
                    code = '''import importlib, pkgutil, sys
sys.path.insert(0, sys.argv[1])
import wps_skills
for module in pkgutil.walk_packages(wps_skills.__path__, 'wps_skills.'):
    importlib.import_module(module.name)
app = sys.argv[2]
contracts = importlib.import_module('wps_skills.' + app + '.contracts')
assert len(getattr(contracts, app.upper() + '_PRODUCTION_CONTRACT_SET').action_names) == int(sys.argv[3])
session = importlib.import_module('wps_skills.' + app + '.windows.session')
assert session.BRIDGE_SCRIPT.is_file()
'''
                    subprocess.run([sys.executable, '-I', '-c', code, str(python.parent), app, str(count)], cwd=root, env=env, check=True)
                    entry = package / 'scripts' / (app + '.py')
                    for other in {'word', 'excel', 'ppt'} - {app}:
                        result = subprocess.run([sys.executable, str(entry), '--app', other, '--index'], cwd=root, env=env, capture_output=True, text=True)
                        self.assertEqual(4, result.returncode, result.stderr)
                        self.assertEqual('', result.stdout)
                        self.assertEqual('WPS_DISCOVERY_UNAVAILABLE app=' + other + '\n', result.stderr)
