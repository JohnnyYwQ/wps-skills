"""Distribution boundaries for the combined Codex and Claude Code plugin."""

import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile

from wps_skills.cli.build_plugin import build_plugin


class PluginBuildTests(unittest.TestCase):
    def test_relocated_archive_contains_both_hosts_and_self_contained_skills(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            built = build_plugin(root / 'original' / 'wps-skills')
            archive_path = built.parent / 'wps-skills.zip'
            self.assertEqual(hashlib.sha256(archive_path.read_bytes()).hexdigest(),
                             (built.parent / 'wps-skills.zip.sha256').read_text().split()[0])
            relocated = root / '安装 plugin' / 'wps-skills'
            with zipfile.ZipFile(archive_path) as archive:
                self.assertIsNone(archive.testzip())
                archive.extractall(relocated.parent)
            shutil.rmtree(built.parent)
            recorded = json.loads((relocated / 'files.sha256.json').read_text(encoding='utf-8'))
            actual = {p.relative_to(relocated).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
                      for p in relocated.rglob('*') if p.is_file() and p != relocated / 'files.sha256.json'}
            self.assertEqual(recorded['files'], actual)
            for name in actual:
                self.assertTrue({'experiments', 'test', '__pycache__'}.isdisjoint(Path(name).parts), name)
                self.assertNotIn(Path(name).name, {'demo.py', 'build_plugin.py', 'demo_launcher.ps1'})
            codex_market = json.loads((relocated / '.agents/plugins/marketplace.json').read_text(encoding='utf-8'))
            claude_market = json.loads((relocated / '.claude-plugin/marketplace.json').read_text(encoding='utf-8'))
            self.assertEqual(codex_market['name'], claude_market['name'])
            plugin = (relocated / codex_market['plugins'][0]['source']['path']).resolve()
            self.assertEqual(plugin, (relocated / claude_market['plugins'][0]['source']).resolve())
            codex = json.loads((plugin / '.codex-plugin/plugin.json').read_text(encoding='utf-8'))
            claude = json.loads((plugin / '.claude-plugin/plugin.json').read_text(encoding='utf-8'))
            self.assertEqual(plugin.name, codex['name'])
            self.assertEqual(codex['name'], claude['name'])
            self.assertEqual(recorded['version'], codex['version'])
            self.assertEqual(codex['version'], claude['version'])
            self.assertNotIn('interface', claude)
            self.assertEqual({'wps-word', 'wps-excel', 'wps-ppt'},
                             {p.name for p in (plugin / 'skills').iterdir()})
            environment = dict(os.environ, PYTHONPATH='', PYTHONNOUSERSITE='1',
                               PYTHONDONTWRITEBYTECODE='1', PYTHONUTF8='1')
            for app, count in (('word', 14), ('excel', 34), ('ppt', 37)):
                with self.subTest(app=app):
                    skill = plugin / 'skills' / ('wps-' + app)
                    entry = skill / 'scripts' / (app + '.py')
                    result = subprocess.run([sys.executable, str(entry), '--app', app, '--index'],
                                            cwd=root, env=environment, capture_output=True,
                                            text=True, encoding='utf-8', timeout=30)
                    self.assertEqual(0, result.returncode, result.stderr)
                    names = [a['action'] for a in json.loads(result.stdout)['actions']]
                    self.assertEqual(count, len(names))
                    result = subprocess.run([sys.executable, str(skill / 'scripts/schema.py'), *names],
                                            cwd=root, env=environment, capture_output=True,
                                            text=True, encoding='utf-8', timeout=30)
                    self.assertEqual(0, result.returncode, result.stderr)
                    self.assertEqual(set(names), set(json.loads(result.stdout)['actions']))
                    manifest = json.loads((skill / 'runtime/files.sha256.json').read_text(encoding='utf-8'))
                    for name, digest in manifest.items():
                        self.assertEqual(digest, hashlib.sha256((skill / name).read_bytes()).hexdigest())

    def test_existing_directory_or_archive_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            destination = root / 'wps-skills'
            destination.mkdir()
            marker = destination / 'user.txt'
            marker.write_text('keep', encoding='utf-8')
            with self.assertRaises(FileExistsError):
                build_plugin(destination)
            self.assertEqual('keep', marker.read_text(encoding='utf-8'))
            archive = root / 'other.zip'
            archive.write_bytes(b'previous package')
            with self.assertRaises(FileExistsError):
                build_plugin(root / 'other')
            self.assertEqual(b'previous package', archive.read_bytes())
            self.assertFalse((root / 'other').exists())


if __name__ == '__main__':
    unittest.main()
