"""Exercise batch scheduling at the subprocess/response boundary without WPS."""
import json
import os
from types import SimpleNamespace
from pathlib import Path, PureWindowsPath
import subprocess
import tempfile
import unittest
from unittest.mock import patch
from tests.applications import action_acceptance as runner


class BatchRunnerTests(unittest.TestCase):
    def run_scenario(self,first_cleanup='succeeded',timeout=False,dependency=False):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);entries=[]
            for name in ('first','second','third'):
                request='requests/'+name+'.json';expected='expected/'+name+'.json'
                runner.write(root/request,{'app':'word'});runner.write(root/expected,{'kind':'positive','caseId':name,'setupOnly':name=='first' and dependency})
                entries.append({'id':name,'request':request,'expected':expected,'requestSha256':runner.sha(root/request),'expectedSha256':runner.sha(root/expected),'dependsOn':['first'] if dependency and name=='second' else []})
            runner.write(root/'manifest.json',{'app':'word','windowsRoot':PureWindowsPath(root).as_posix(),'skillManifestSha256':'hash','assets':{},'coverage':{'writeContent':{}},'cases':entries,'validationMode':'response_only','caseFailurePolicy':'continue_independent'})
            responses={};submitted=[]
            def run(cmd,**kw):
                if '-ProgId' in cmd:return subprocess.CompletedProcess(cmd,0,json.dumps({'registered':True,'userInteractive':True,'sessionId':1,'echo':'中文往返'}),'')
                file=Path(cmd[-1]);name=file.stem
                if '--task-file' in cmd:
                    submitted.append(name)
                    if timeout and name=='first':raise subprocess.TimeoutExpired(cmd,300)
                    r={'type':'task.response','app':'word','taskId':name,'state':'completed','outcome':'succeeded','document':None,'steps':[],'completion':{},'cleanup':{'outcome':first_cleanup if name=='first' else 'succeeded'},'stop':None}
                    responses[name]=r;file.unlink();kw['stdout'].write(json.dumps(r)+'\n')
                    return subprocess.CompletedProcess(cmd,0)
                return subprocess.CompletedProcess(cmd,0,json.dumps(responses[name])+'\n','')
            def verdict(expected,*unused):return {'passed':expected['caseId']!='first','mismatches':[] if expected['caseId']!='first' else [{'field':'business oracle'}],'responseVerifiedActions':['writeContent']}
            with patch.object(runner,'os',SimpleNamespace(name='nt',environ=os.environ,getpid=os.getpid,replace=os.replace)),patch.object(runner,'verify_skill',return_value='hash'),patch('wps_skills.windows.bridge_runtime.windows_powershell_executable',return_value='powershell.exe'),patch.object(runner.subprocess,'run',side_effect=run),patch.object(runner,'evaluate',side_effect=verdict):
                runner.execute_prepared(root,root/'skill',root/'resources')
            return submitted,runner.read(root/'report.json')

    def test_failed_oracle_continues_without_retry(self):
        submitted,report=self.run_scenario()
        self.assertEqual(submitted,['first','second','third'])
        self.assertEqual([c['state'] for c in report['cases']],['failed','passed','passed'])
        self.assertEqual(report['state'],'completed_with_failures')

    def test_failed_setup_blocks_only_dependent_case(self):
        submitted,report=self.run_scenario(dependency=True)
        self.assertEqual(submitted,['first','third'])
        self.assertEqual([c['state'] for c in report['cases']],['failed','blocked','passed'])

    def test_unproved_cleanup_stops_application(self):
        submitted,report=self.run_scenario(first_cleanup='failed')
        self.assertEqual(submitted,['first']);self.assertEqual(report['state'],'failed')

    def test_timeout_does_not_continue(self):
        submitted,report=self.run_scenario(timeout=True)
        self.assertEqual(submitted,['first']);self.assertEqual(report['state'],'failed')

if __name__=='__main__':unittest.main()
