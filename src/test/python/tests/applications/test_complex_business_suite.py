"""Verify frozen suite size, legal requests, Action coverage and observable oracles."""
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
from tests.applications.complex_business_suite import build_cases
from tests.applications.action_suite import coverage


def result_path_exists(schema,path):
    if not path:return True
    if 'oneOf' in schema:return any(result_path_exists(s,path) for s in schema['oneOf'])
    head,*tail=path
    if head=='*' or isinstance(head,(int,dict)):
        return 'items' in schema and result_path_exists(schema['items'],tail)
    return head in schema.get('properties',{}) and result_path_exists(schema['properties'][head],tail)


class ComplexBusinessTests(unittest.TestCase):
    def test_relocated_complex_kit_prepares_all_apps_without_repository_imports(self):
        repository = Path(__file__).resolve().parents[5]
        environment = dict(os.environ, PYTHONPATH='', PYTHONNOUSERSITE='1',
                           PYTHONDONTWRITEBYTECODE='1', PYTHONUTF8='1')
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            built = root / 'build' / 'complex-kit'
            result = subprocess.run(
                [sys.executable, str(repository / 'scripts/build/complex_acceptance.py'),
                 '--output', str(built)], cwd=root, env=environment,
                capture_output=True, text=True, encoding='utf-8', timeout=60)
            self.assertEqual(0, result.returncode, result.stderr)
            installed = root / 'relocated kit'
            with zipfile.ZipFile(built.with_suffix('.zip')) as archive:
                archive.extractall(installed)
            shutil.rmtree(built)
            self.assertEqual({'run_word.py', 'run_excel.py', 'run_ppt.py'},
                             {entry.name for entry in installed.glob('*.py')})
            manifest = json.loads((installed / 'kit-manifest.json').read_text(encoding='utf-8'))
            for name, digest in manifest['files'].items():
                self.assertEqual(digest, hashlib.sha256((installed / name).read_bytes()).hexdigest(), name)
            for app, count in (('word', 14), ('excel', 34), ('ppt', 37)):
                with self.subTest(app=app):
                    destination = root / 'prepared' / app
                    result = subprocess.run(
                        [sys.executable, str(installed / ('run_' + app + '.py')),
                         '--prepare-only', '--windows-root', 'C:/wps-tests/' + app,
                         '--root', str(destination)], cwd=root, env=environment,
                        capture_output=True, text=True, encoding='utf-8', timeout=30)
                    self.assertEqual(0, result.returncode, result.stderr)
                    prepared = json.loads((destination / 'manifest.json').read_text(encoding='utf-8'))
                    self.assertEqual(20, prepared['businessTaskCount'])
                    self.assertEqual(20, sum(not case['setupOnly'] for case in prepared['cases']))
                    self.assertEqual(count, len(prepared['coverage']))
                    self.assertFalse((destination / 'run.claim').exists())
                    for case in prepared['cases']:
                        for field in ('request', 'expected'):
                            self.assertEqual(case[field + 'Sha256'], hashlib.sha256(
                                (destination / case[field]).read_bytes()).hexdigest())

    def test_all_business_requests_and_oracles(self):
        root=Path(__file__).resolve().parents[5]
        for app,count,prefix in [('word',14,'CW'),('excel',34,'CE'),('ppt',37,'CP')]:
            with self.subTest(app=app):
                cases=build_cases(app,'C:/wps-tests/contract-check/'+app)
                measured=[c for c in cases if not c.setup]
                self.assertEqual([c.id for c in measured],[prefix+f'{i:02d}' for i in range(1,21)])
                self.assertEqual(len(coverage(app,measured)),count)
                schemas=json.loads((root/f'src/main/resources/skills/wps-{app}/references/actions.json').read_text())
                for case in cases:
                    oracle=case.expectation()
                    for step in case.request['steps']:
                        if step['address']['action']=='findContent':
                            for rule in case.checks.get(step['id'],[]):
                                if rule['path']==['matches'] and rule['op']=='length':
                                    self.assertLessEqual(rule['value'],step['params']['limit'])
                    self.assertEqual(oracle['kind'],'positive')
                    self.assertLessEqual(len(case.request['steps']),125)
                    for action in oracle['actions']:
                        self.assertTrue(action['dataAssertions'],(case.id,action['id']))
                        for rule in action['dataAssertions']:
                            self.assertTrue(result_path_exists(schemas[action['address']['action']]['result'],rule['path']),(case.id,action['id'],rule))

if __name__=='__main__':unittest.main()
