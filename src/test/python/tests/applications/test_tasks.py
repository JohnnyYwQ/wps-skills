"""Shared Task guarantees at the public JSON and packaged CLI interfaces."""
import copy
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from wps_skills.client import task_client, task_store
from wps_skills.client.applications import contracts_for, compile_request
from wps_skills.cli.build_skill import build_application_skill
from tests.applications.native_acceptance import Plan, ref, excel_plan, ppt_plan


class TaskMigrationTests(unittest.TestCase):
    def test_full_preplanned_requests_validate_with_forward_result_references(self):
        for app,build in [('excel',excel_plan),('ppt',ppt_plan)]:
            with self.subTest(app=app):
                request=build(Path('C:/test/run/outputs'))
                task_client._preflight(compile_request(request,app),contracts_for(app))

    def test_all_examples_are_admitted_and_handlers_complete(self):
        import importlib
        for app in ('excel','ppt'):
            contracts=contracts_for(app)
            handlers=getattr(importlib.import_module('wps_skills.'+app+'.handlers'),app.upper()+'_HANDLERS')
            for contract in contracts.contracts:
                with self.subTest(app=app,action=contract.name):
                    if contract.binding_role=='required':self.assertIn(contract.name,handlers)
                    for example in contract.to_wire(app)['examples']:
                        contracts.validate_params(contract.name,example['params'])

    def test_lifecycle_and_reference_rejections_are_pre_dispatch(self):
        for app in ('excel','ppt'):
            base=Plan(app).finish()
            cases=[]
            req=copy.deepcopy(base);req['version']=2;cases.append(req)
            req=copy.deepcopy(base);req['completion']=[Plan(app).action('save')];cases.append(req)
            req=copy.deepcopy(base);req['document']['address']['app']='word';cases.append(req)
            req=copy.deepcopy(base);req['steps']=[{'id':'read','address':{'app':app,'action':'listSlides' if app=='ppt' else 'listWorksheets'},'params':{'bad':ref('later','token')}}];cases.append(req)
            for request in cases:
                factory=lambda **kwargs: self.fail('Rejected plan must not acquire resources')
                result=task_client.execute(request,app,executor_factory=factory,source_path='invalid.json')
                self.assertEqual('rejected',result['state']);self.assertEqual(app,result['app'])

    def test_receipts_stop_no_replay_and_existing_changes_protection(self):
        for app,read in [('excel','getWorkbookInfo'),('ppt','getPresentationInfo')]:
            with tempfile.TemporaryDirectory() as tmp,patch.dict(os.environ,WPS_SKILLS_TASK_DIR=tmp):
                p=Plan(app);first=p.add(read);p.add(read)
                calls=[]
                class Executor:
                    can_execute=True
                    def execute(self,address,params):
                        calls.append(address['action'])
                        if address['action']==read:return {'outcome':'failed','error':{'code':'TEST_FAILURE','message':'definite failure'}}
                        return {'outcome':'succeeded','data':{'documentState':{'persistenceState':'unsaved'}}}
                    def close(self):return {'outcome':'succeeded','error':None}
                result=task_client.execute(p.finish(),app,source_path=str(Path(tmp)/'in.json'),executor_factory=lambda **kwargs:Executor())
                self.assertEqual(first,result['stop']['stepId']);self.assertEqual('not_executed',result['steps'][1]['state'])
                repeated=task_client.execute(p.finish(),app,source_path=str(Path(tmp)/'in.json'),executor_factory=lambda **kwargs:self.fail('No replay'))
                self.assertEqual(result,repeated);self.assertEqual(2,len(calls))
                self.assertEqual(result,task_store.status_file(app,str(Path(tmp)/'in.json')))
                p=Plan(app,'C:/test/existing.'+('xlsx' if app=='excel' else 'pptx'));p.add(read)
                result=task_client.execute(p.finish(save=True),app,source_path=str(Path(tmp)/'modified.json'),executor_factory=lambda **kwargs:Executor())
                self.assertEqual('TASK_EXISTING_CHANGES_CONFIRMATION_REQUIRED',result['stop']['error']['code'])
                self.assertEqual('not_executed',result['completion']['save']['state'])

    def test_annotation_is_informational_and_does_not_relax_validation(self):
        from wps_skills.core.action_runtime import ActionContract, ApplicationContractSet
        from dataclasses import replace
        from wps_skills.excel.contracts import EXCEL_FORMAT_VALIDATORS
        base=contracts_for('excel').resolve('readRange')
        wire=base.to_wire('excel')
        params=wire['parameters']
        params['properties']['address']['description']='Uppercase A1 region.'
        updated=replace(base,parameters=params)
        contracts=ApplicationContractSet(application='excel',contracts=(updated,),format_validators=EXCEL_FORMAT_VALIDATORS)
        contracts.validate_params('readRange',{'sheet':'Sheet1','address':'A1'})
        with self.assertRaises(ValueError):contracts.validate_params('readRange',{'sheet':'Sheet1','address':'wrong'})
        params['properties']['address']['description']=42
        with self.assertRaises(ValueError):ApplicationContractSet(application='excel',contracts=(replace(base,parameters=params),))

    def test_all_factories_construct_the_shared_executor_without_starting_wps(self):
        from wps_skills.windows.task_factory import build_task
        from wps_skills.core.task_executor import ApplicationTask
        from types import SimpleNamespace
        with tempfile.TemporaryDirectory() as tmp,patch.dict(os.environ,WPS_TRACE_DIR=tmp):
            for app in ('word','excel','ppt'):
                launcher=SimpleNamespace(close=lambda:())
                with self.subTest(app=app),patch('wps_skills.windows.owned_process.WindowsOwnedProcessLauncher',return_value=launcher):
                    task=build_task(application=app,task_id='test-'+app)
                    self.assertIsInstance(task,ApplicationTask)
                    self.assertEqual('succeeded',task.close()['outcome'])

    def test_relocated_packages_query_lossless_schemas_without_other_apps(self):
        for app in ('excel','ppt'):
            with self.subTest(app=app),tempfile.TemporaryDirectory() as tmp:
                root=Path(tmp);built=build_application_skill(app,root/('wps-'+app))
                runtime=built/'runtime/src/main/python/wps_skills'
                for other in {'excel','ppt','word'}-{app}:self.assertFalse((runtime/other).exists())
                manifest=json.loads((built/'runtime/files.sha256.json').read_text())
                for path,expected in manifest.items():self.assertEqual(expected,hashlib.sha256((built/path).read_bytes()).hexdigest())
                env=dict(os.environ,PYTHONPATH='',PYTHONNOUSERSITE='1',PYTHONDONTWRITEBYTECODE='1')
                entry=built/'scripts'/(app+'.py')
                result=subprocess.run([sys.executable,str(entry),'--app',app,'--index'],cwd=root,env=env,capture_output=True,text=True,check=True)
                names=[c.name for c in contracts_for(app).contracts]
                self.assertEqual(set(names),{v['action'] for v in json.loads(result.stdout)['actions']})
                result=subprocess.run([sys.executable,str(built/'scripts/schema.py'),*names],cwd=root,env=env,capture_output=True,text=True,check=True)
                view=json.loads(result.stdout)
                def expand(value):
                    if isinstance(value,dict):
                        if set(value)=={'$ref'}:return expand(view['$defs'][value['$ref'].split('/')[-1]])
                        return {k:expand(v) for k,v in value.items()}
                    if isinstance(value,list):return [expand(v) for v in value]
                    return value
                for c in contracts_for(app).contracts:
                    wire=c.to_wire(app)
                    self.assertEqual({k:wire[k] for k in ('parameters','result','examples')},expand(view['actions'][c.name]))
                code="import sys,importlib,pkgutil;sys.path.insert(0,sys.argv[1]);import wps_skills;[importlib.import_module(m.name) for m in pkgutil.walk_packages(wps_skills.__path__,'wps_skills.')]"
                subprocess.run([sys.executable,'-c',code,str(runtime.parent)],env=env,cwd=root,capture_output=True,text=True,check=True)
                opposite='ppt' if app=='excel' else 'excel'
                result=subprocess.run([sys.executable,str(entry),'--app',opposite,'--index'],cwd=root,env=env,capture_output=True,text=True)
                self.assertNotEqual(0,result.returncode)
