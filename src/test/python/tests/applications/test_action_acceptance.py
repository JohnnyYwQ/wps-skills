import unittest
from wps_skills.client.applications import contracts_for


class AcceptanceSuiteTests(unittest.TestCase):
    def test_every_registered_action_has_a_success_case_and_data_assertion(self):
        from tests.applications.action_suite import build_cases, coverage
        for app in ('word', 'excel', 'ppt'):
            with self.subTest(app=app):
                cases = build_cases(app, 'C:/tests/fresh')
                matrix = coverage(app, cases)
                self.assertEqual({c.name for c in contracts_for(app).contracts}, set(matrix))
                for name, row in matrix.items():
                    self.assertTrue(row['positiveCases'], name)
                    self.assertTrue(row['dataAssertionCases'], name)

    def test_same_run_basename_in_different_directories_never_reuses_output_names(self):
        from tests.applications.action_suite import build_cases
        from pathlib import PureWindowsPath
        first=build_cases('word','C:/tests/run-a/word')[0]
        second=build_cases('word','C:/tests/run-b/word')[0]
        self.assertNotEqual(PureWindowsPath(first.path('.docx')).name, PureWindowsPath(second.path('.docx')).name)

    def test_all_assertion_paths_and_result_references_exist_in_current_contracts(self):
        from tests.applications.action_suite import build_cases
        from tests.applications.test_complex_plan import schema_at, refs
        from wps_skills.client.applications import compile_request
        for app in ('word','excel','ppt'):
            contracts=contracts_for(app)
            for case in build_cases(app,'C:/tests/fresh'):
                schemas={s['id']:contracts.resolve(s['address']['action']).to_wire(app)['result'] for s in compile_request(case.request,app)['steps']}
                for step in compile_request(case.request,app)['steps']:
                    for r in refs(step['params']):schema_at(schemas[r['step']],r['path'][1:])
                for step in case.expectation()['actions']:
                    for rule in step['dataAssertions']:
                        schema_at(schemas[step['id']],rule['path'])
                        for r in refs(rule):schema_at(schemas[r['step']],r['path'][1:])

    def test_preflight_rejection_cannot_hide_execution_or_consumption(self):
        from tests.applications.action_acceptance import evaluate
        import copy
        expected={'kind':'preflight_rejection'}
        response={'state':'rejected','outcome':'failed','taskId':None,'recordPath':None,
                  'stop':{'phase':'validation','error':{'code':'INVALID_TASK_REQUEST'}},
                  'steps':[{'state':'not_executed','response':None}], 'taskFile':{'state':'retained'}}
        self.assertTrue(evaluate(expected,{},response,4)['passed'])
        for patch in ({'taskId':'unexpected'},{'taskFile':{'state':'removed'}},{'steps':[{'state':'succeeded','response':{}}]}):
            changed=copy.deepcopy(response);changed.update(patch)
            self.assertFalse(evaluate(expected,{},changed,4)['passed'])

    def test_missing_or_multiple_final_responses_never_count_as_success(self):
        from tests.applications.action_acceptance import parse_response
        for output in ('', '{"type":"task.progress"}', '{"type":"task.response"}\n{"type":"task.response"}'):
            with self.assertRaises(ValueError):parse_response(output)

    def test_open_artifact_size_is_historical_after_in_place_save(self):
        import tempfile, zipfile
        from pathlib import Path
        from tests.applications.action_acceptance import artifacts
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);path=root/'changed.docx'
            with zipfile.ZipFile(path,'w') as z:z.writestr('word/document.xml','<document/>')
            def step(name,size):
                return {'address':{'app':'word','action':name},'response':{'data':{'artifact':{'path':str(path),'sizeBytes':size}}}}
            response={'document':step('openDocument',1),'steps':[],'completion':{'save':step('save',path.stat().st_size),'pdf':None}}
            self.assertEqual(1,len(artifacts(response,root,'fixture')))

    def test_portable_kit_prepares_all_three_apps_without_repository_imports(self):
        import tempfile, subprocess, sys, os, json
        from pathlib import Path
        from tests.applications.build_acceptance import build
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);kit=root/'kit';build(kit)
            for app,count in (('word',14),('excel',34),('ppt',37)):
                run=root/app
                process=subprocess.run([sys.executable,'-B',str(kit/('run_'+app+'.py')),
                    '--prepare-only','--windows-root','C:/acceptance/'+app,'--root',str(run)],
                    cwd=tmp,env=dict(os.environ,PYTHONPATH='',PYTHONNOUSERSITE='1'),capture_output=True,text=True,timeout=30)
                self.assertEqual(0,process.returncode,process.stderr)
                manifest=json.loads((run/'manifest.json').read_text())
                self.assertEqual(count,len(manifest['coverage']))
                self.assertFalse((run/'run.claim').exists())
