"""Offline checks for prepared complex cases, not native acceptance results."""
import copy
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from tests.applications.complex_plan import excel_cases,ppt_cases,write_plan
from tests.applications.response_expectations import compare,select,matches
from wps_skills.client.applications import contracts_for,compile_request

ROOT='C:/Users/yim/wps-complex-review-test'


def schema_at(schema,path):
    if not path:return schema
    if 'oneOf' in schema or 'anyOf' in schema:
        errors=[]
        for branch in schema.get('oneOf',schema.get('anyOf')):
            try:return schema_at(branch,path)
            except (KeyError,AssertionError):errors.append(branch)
        raise AssertionError('No schema branch contains '+str(path))
    head,*tail=path
    if isinstance(head,dict):
        assert schema['type']=='array'
        schema_at(schema['items'],[head['field']])
        return schema_at(schema['items'],tail)
    if head=='*' or type(head)==int:
        assert schema['type']=='array'
        if type(head)==int and 'maxItems' in schema:assert head<schema['maxItems']
        return schema_at(schema['items'],tail)
    return schema_at(schema['properties'][head],tail)


def refs(value):
    if isinstance(value,dict):
        if '$ref' in value:yield value['$ref']
        for v in value.values():yield from refs(v)
    elif isinstance(value,list):
        for v in value:yield from refs(v)


class ComplexPlanTests(unittest.TestCase):
    def test_all_requests_preflight_and_every_assertion_path_exists_in_contract(self):
        for builder in (excel_cases,ppt_cases):
            cases,setup=builder(ROOT)
            self.assertEqual(12,len(cases));self.assertEqual(1,len(setup))
            for case in setup+cases:
                with self.subTest(case=case.id):
                    expectation=case.expectation()
                    contracts=contracts_for(case.app)
                    plan=compile_request(case.request,case.app)
                    schemas={s['id']:contracts.resolve(s['address']['action']).to_wire(case.app)['result'] for s in plan['steps']}
                    for s in plan['steps']:
                        for r in refs(s['params']):
                            self.assertEqual('data',r['path'][0]);schema_at(schemas[r['step']],r['path'][1:])
                    for s in expectation['actions']:
                        for rule in s['dataAssertions']:
                            schema_at(schemas[s['id']],rule['path'])
                            for r in refs(rule):schema_at(schemas[r['step']],r['path'][1:])
                    self.assertEqual(len(plan['steps']),len(expectation['actions']))

    def test_preparation_writes_concrete_utf8_inputs_and_never_executes(self):
        with tempfile.TemporaryDirectory() as tmp,patch('wps_skills.client.task_client.execute',side_effect=AssertionError('must not execute')):
            output=Path(tmp)/'plan';manifest=write_plan(ROOT,output)
            self.assertEqual('prepared_not_executed',manifest['status'])
            self.assertEqual(24,len(manifest['cases']))
            paths=[]
            for item in manifest['setup']+manifest['cases']:
                request=json.loads((output/item['request']).read_text(encoding='utf-8'))
                self.assertEqual({'app','document','steps','completion'},set(request))
                text=json.dumps(request);self.assertNotIn('{{ROOT}}',text)
                for step in request['completion']:
                    if 'outputPath' in step['params']:paths.append(step['params']['outputPath'])
            self.assertEqual(len(paths),len(set(paths)))
            with self.assertRaises(FileExistsError):write_plan(ROOT,output)

    def test_dynamic_identity_matrix_and_strict_missing_observations(self):
        responses={'create':{'data':{'shapes':[{'id':91}]}},'copy':{'data':{'cells':[[{'value':10},{'value':20}]]}}}
        path=['shapes',{'field':'id','value':{'$ref':{'step':'create','path':['data','shapes',0,'id']}}},'left']
        self.assertEqual(40,select({'shapes':[{'id':91,'left':40}]},path,responses))
        self.assertTrue(matches([[10,20]],{'op':'equals','value':{'$response':{'step':'copy','path':['cells','*','*','value']}}},responses))
        self.assertFalse(matches(True,{'op':'equals','value':1},responses))
        with self.assertRaises(ValueError):select({'cells':[]},['cells','*','value'],responses)
        with self.assertRaises(ValueError):select({'shapes':[{'id':91},{'id':91}]},path,responses)

    def test_negative_case_requires_exact_failure_and_unexecuted_tail(self):
        case=excel_cases(ROOT)[0][-1];expected=case.expectation()
        # Only exercise structural negative-case reporting here; document data
        # oracles are covered by the selector test, not fabricated WPS results.
        for s in expected['actions']:s['dataAssertions']=[]
        records=[]
        for s in expected['actions']:
            r=None if s['state']=='not_executed' else {'outcome':s['responseOutcome'],'address':s['address'],'taskId':'t','traceId':'trace-'+s['id']}
            if 'errorCode' in s:r['error']={'code':s['errorCode']}
            records.append({'id':s['id'],'address':s['address'],'state':s['state'],'response':r})
        body=len(case.request['steps'])
        response={'type':'task.response','taskId':'t','recordPath':'receipt.json','app':'excel','state':'stopped','outcome':'failed','stop':expected['task']['stop'],'cleanup':{'outcome':'succeeded'},'document':records[0],'steps':records[1:1+body],'completion':{'save':records[-2],'pdf':records[-1]}}
        self.assertTrue(compare(expected,response)['passed'])
        broken=copy.deepcopy(response);broken['completion']['save']['state']='succeeded'
        self.assertFalse(compare(expected,broken)['passed'])
        broken=copy.deepcopy(response);broken['steps'][0]['response']['taskId']='other'
        self.assertFalse(compare(expected,broken)['passed'])
