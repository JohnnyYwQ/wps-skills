"""Windows experiments through real Task CLI, receipts and independent WPS observations."""
import argparse
import copy
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import traceback
import uuid
import zipfile
from xml.etree import ElementTree as ET

from .plans import BASE, MARK, SHEET, seed, edit, primary_action
from . import GROUPS


def write(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp')
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    os.replace(temporary, path)


def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def require(condition, message, **observed):
    if not condition:
        raise AssertionError(message + ': ' + json.dumps(observed, ensure_ascii=False, default=str))


def response(text):
    values=[]
    for line in text.splitlines():
        try: value=json.loads(line)
        except ValueError: continue
        if isinstance(value,dict) and value.get('type')=='task.response':values.append(value)
    require(len(values)==1,'Expected one Task Response',count=len(values))
    return values[0]


def document_text(path):
    """Read the saved OOXML without importing any WPS Action verification code."""
    path=Path(path)
    with zipfile.ZipFile(path) as archive:
        require(archive.testzip() is None,'OOXML CRC')
        if path.suffix=='.docx':
            return ''.join(t.text or '' for t in ET.fromstring(archive.read('word/document.xml')).findall('.//{*}t'))
        if path.suffix=='.pptx':
            parts=[n for n in archive.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')]
            return '\n'.join(''.join(t.text or '' for t in ET.fromstring(archive.read(n)).findall('.//{*}t')) for n in parts)
        shared=[]
        if 'xl/sharedStrings.xml' in archive.namelist():
            shared=[''.join(t.text or '' for t in si.findall('.//{*}t')) for si in ET.fromstring(archive.read('xl/sharedStrings.xml'))]
        values=[]
        for name in archive.namelist():
            if not name.startswith('xl/worksheets/sheet') or not name.endswith('.xml'):continue
            for c in ET.fromstring(archive.read(name)).findall('.//{*}c'):
                v=c.find('{*}v')
                if c.get('t')=='s' and v is not None:values.append(shared[int(v.text)])
                elif c.get('t')=='inlineStr':values.append(''.join(t.text or '' for t in c.findall('.//{*}t')))
                elif v is not None:values.append(v.text or '')
        return '\n'.join(values)


NATIVE_OPERATION={'word':'insert_structured_body_content','excel':'insertRows','ppt':'addSlide'}


class ProcessWitness:
    """Keep a native handle before failure injection so PID reuse cannot fool cleanup checks."""
    def __init__(self,pid):
        import ctypes
        from ctypes import wintypes
        self.api=ctypes.WinDLL('kernel32',use_last_error=True)
        self.api.OpenProcess.argtypes=[wintypes.DWORD,wintypes.BOOL,wintypes.DWORD]
        self.api.OpenProcess.restype=wintypes.HANDLE
        self.api.WaitForSingleObject.argtypes=[wintypes.HANDLE,wintypes.DWORD]
        self.api.WaitForSingleObject.restype=wintypes.DWORD
        self.api.CloseHandle.argtypes=[wintypes.HANDLE]
        self.handle=self.api.OpenProcess(0x100000,False,pid)
        require(bool(self.handle),'Cannot witness owned bridge process',pid=pid)
        self.pid=pid

    def assert_exited(self):
        try:require(self.api.WaitForSingleObject(self.handle,2000)==0,'Owned bridge survived owner cleanup',pid=self.pid)
        finally:self.api.CloseHandle(self.handle)


class Experiments:
    def __init__(self,kit,root):
        self.kit,self.root=kit,root
        self.ps=os.path.join(os.environ.get('WINDIR','C:/Windows'),'System32/WindowsPowerShell/v1.0/powershell.exe')
        self.env=dict(os.environ,PYTHONUTF8='1',PYTHONIOENCODING='utf-8',WPS_TRACE_DIR=str(root/'traces'),WPS_SKILLS_TASK_DIR=str(root/'receipts'))
        self.serial=0
        self.tasks=[]

    def launch(self,app,request,label,*,config=None,source=None,restore=False,timeout=60,lost_output=False):
        self.serial+=1
        directory=self.root/'calls'/('%04d-'%self.serial+label)
        directory.mkdir(parents=True)
        write(directory/'request.json',request)
        source=Path(source) if source else self.root/'submissions'/('%04d-'%self.serial+label+'.json')
        if not source.exists() and (restore or source.parent==self.root/'submissions' and not any(t['source']==str(source) for t in self.tasks)):
            write(source,request)
        spec=dict(app=app,caseDir=str(directory),ownedRoot=str(self.root),stage='',operation='',mode='',sheet=SHEET)
        spec.update(config or {})
        write(directory/'config.json',spec)
        command=[sys.executable,str(self.kit/'design_value/child.py'),'--skill',str(self.kit/'skills'/('wps-'+app)),
                 '--config',str(directory/'config.json'),'--','--app',app,'--task-file',str(source),'--timeout',str(timeout)]
        out=(directory/'stdout.jsonl').open('w',encoding='utf-8');err=(directory/'stderr.jsonl').open('w',encoding='utf-8')
        started=time.perf_counter()
        process=subprocess.Popen(command,env=self.env,stdout=subprocess.DEVNULL if lost_output else out,stderr=err)
        record=dict(app=app,label=label,source=str(source),directory=str(directory),pid=process.pid,started=started,
                    requestSha256=sha(directory/'request.json'),configSha256=sha(directory/'config.json'),lostOutput=lost_output)
        self.tasks.append(record)
        write(directory/'launch.json',{k:v for k,v in record.items() if k!='started'})
        return process,out,err,record

    def finish(self,handle,*,killed=False):
        process,out,err,record=handle
        try:code=process.wait(timeout=120)
        except subprocess.TimeoutExpired:
            process.kill();process.wait(timeout=10)
            raise AssertionError('Experiment child exceeded watchdog; effects preserved')
        finally:out.close();err.close()
        record.update(exitCode=code,seconds=time.perf_counter()-record.pop('started'))
        write(Path(record['directory'])/'process.json',record)
        if killed:return record,None
        if record['lostOutput']:value=self.query(record['app'],record['source'],Path(record['directory'])/'recovered')
        else:value=response((Path(record['directory'])/'stdout.jsonl').read_text(encoding='utf-8-sig'))
        write(Path(record['directory'])/'response.json',value)
        return record,value

    def task(self,app,request,label,**kwargs):
        return self.finish(self.launch(app,request,label,**kwargs))

    def query(self,app,source,target):
        cp=subprocess.run([sys.executable,str(self.kit/'skills'/('wps-'+app)/'scripts'/(app+'.py')),'--app',app,'--task-status-file',str(source)],env=self.env,capture_output=True,text=True,encoding='utf-8',timeout=30)
        write(target.with_suffix('.json'),{'exitCode':cp.returncode,'stdout':cp.stdout,'stderr':cp.stderr})
        return response(cp.stdout)

    def observe(self,app,path,label,mode='read'):
        cp=subprocess.run([self.ps,'-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(self.kit/'resources/observe.ps1'),
                           '-Application',app,'-Path',str(path),'-OwnedRoot',str(self.root),'-Mode',mode],capture_output=True,text=True,encoding='utf-8',timeout=30)
        write(self.root/'observations'/(label+'.json'),{'exitCode':cp.returncode,'stdout':cp.stdout,'stderr':cp.stderr})
        require(cp.returncode==0,'Independent COM observation failed',stderr=cp.stderr)
        return json.loads(cp.stdout)

    def count(self,app,path,label):
        observed=self.observe(app,path,label)
        return '\n'.join(str(t or '') for t in observed['texts']).count(MARK)

    def close_saved(self,app,path,label):
        observed=self.observe(app,path,label)
        if observed['saved']:self.observe(app,path,label+'-close',mode='close')

    def fixture(self,app,label):
        plan=seed(app,self.root.as_posix(),label)
        record,result=self.task(app,plan.request,label+'-seed')
        require(result['outcome']=='succeeded' and result['cleanup']['outcome']=='succeeded','Seed Task failed',result=result)
        require(BASE in document_text(plan.path) and MARK not in document_text(plan.path),'Independent seed artifact incorrect')
        return plan

    def wait_ready(self,handle):
        process,_,_,record=handle
        path=Path(record['directory'])/'ready.json';end=time.monotonic()+30
        while not path.exists():
            require(process.poll() is None,'Child exited before injection barrier',directory=record['directory'])
            require(time.monotonic()<end,'Injection barrier timeout')
            time.sleep(.02)
        return read(path)

    def native_events(self,record):
        path=Path(record['directory'])/'native-events.jsonl'
        return [json.loads(line) for line in path.read_text(encoding='utf-8-sig').splitlines()] if path.exists() else []

    def preflight(self,app,label):
        observations=[]
        for candidate in (False,True):
            plan=self.fixture(app,label+('-late' if candidate else '-early'))
            request=edit(app,self.root.as_posix(),label,plan.path).request
            bad=copy.deepcopy(request['steps'][-1]);bad['id']='invalid_tail';bad['params']['__invalid__']=True
            request['steps'].append(bad)
            record,result=self.task(app,request,label+('-candidate' if candidate else '-baseline'),config={'deferredPreflight':candidate})
            count=self.count(app,plan.path,label+str(candidate))
            require(result['outcome']=='failed','Bad plan reported success',result=result)
            require(count==(1 if candidate else 0),'Unexpected partial effects',candidate=candidate,count=count)
            require(MARK not in document_text(plan.path),'Failure unexpectedly saved document')
            if not candidate:require(result['state']=='rejected' and result['taskId'] is None,'Preflight did not reject before admission')
            if not candidate:require(not self.native_events(record),'Static rejection dispatched native operation')
            observations.append({'candidate':candidate,'liveEffects':count,'diskEffects':0,'seconds':record['seconds'],'stop':result['stop']})
            self.close_saved(app,plan.path,label+str(candidate)+'-end')
        return observations

    def replay(self,app,label):
        plan=self.fixture(app,label)
        request=edit(app,self.root.as_posix(),label,plan.path).request
        first,result=self.task(app,request,label+'-lost',lost_output=True)
        require(result['outcome']=='succeeded','Recovered completion failed',result=result)
        require(self.count(app,plan.path,label+'-once')==1,'Initial effect count')
        _,again=self.task(app,request,label+'-repeat',source=first['source'])
        require(again['taskId']==result['taskId'] and again['steps']==result['steps'],'Repeated identity changed')
        _,restored=self.task(app,request,label+'-restored',source=first['source'],restore=True)
        require(restored['taskId']==result['taskId'],'Restored same file replayed')
        # Change a valid Task field so rejection is about identity, not malformed JSON.
        changed=copy.deepcopy(request);changed['includeExistingChanges']=True
        write(Path(first['source']),changed)
        _,conflict=self.task(app,changed,label+'-conflict',source=first['source'])
        require(conflict['state']=='rejected' and conflict['stop']['error']['code']=='TASK_INPUT_CONFLICT','Input conflict not rejected',result=conflict)
        require(self.count(app,plan.path,label+'-deduplicated')==1,'Duplicate effect')
        _,fresh=self.task(app,request,label+'-new-path')
        require(fresh['outcome']=='succeeded' and fresh['taskId']!=result['taskId'],'New path boundary failed',result=fresh)
        require(document_text(plan.path).count(MARK)==2,'New path must perform new effect')
        self.close_saved(app,plan.path,label+'-end')
        return {'samePathEffects':1,'newPathEffects':2,'taskId':result['taskId'],'newTaskId':fresh['taskId']}

    def focus(self,app,label):
        modes=['focus']+(['active_candidate'] if app=='excel' else [])
        results=[]
        for mode in modes:
            a=self.fixture(app,label+mode+'-A');b=self.fixture(app,label+mode+'-B')
            original_a=self.observe(app,a.path,label+mode+'-A-before')['texts']
            original_b=self.observe(app,b.path,label+mode+'-B-before')['texts']
            request=edit(app,self.root.as_posix(),label,a.path).request
            operation=NATIVE_OPERATION[app]
            if app=='excel':
                # A scalar write avoids structure-edit cache invalidation changing the target a second time.
                from tests.applications.complex_plan import Case
                c=Case(label,app,'精确绑定对照',MARK,self.root.as_posix(),a.path)
                c.region('write','writeRange',SHEET,'A1',values=[[MARK]]);c.finish(save=True)
                request=c.request;operation='write_literal_rectangle'
            record,result=self.task(app,request,label+mode,config={'mode':mode,'operation':operation,'stage':'before','otherPath':b.path})
            require(any(e['event']=='focus_confirmed' for e in self.native_events(record)),'Focus injection did not execute')
            ac=self.count(app,a.path,label+mode+'-A');bc=self.count(app,b.path,label+mode+'-B')
            require(result['outcome']=='succeeded','Focus Task did not complete',result=result)
            require((ac,bc)==((0,1) if mode=='active_candidate' else (1,0)),'Wrong binding effects',a=ac,b=bc,mode=mode)
            other_path=a.path if mode=='active_candidate' else b.path
            original=original_a if mode=='active_candidate' else original_b
            require(self.observe(app,other_path,label+mode+'-unchanged')['texts']==original,'Non-target text changed')
            results.append({'mode':mode,'targetEffects':ac,'otherEffects':bc,'reportedOutcome':result['outcome']})
            self.close_saved(app,a.path,label+mode+'-A-end');self.close_saved(app,b.path,label+mode+'-B-end')
        return results

    def concurrency(self,app,label):
        plan=self.fixture(app,label)
        request=edit(app,self.root.as_posix(),label,plan.path).request
        first=self.launch(app,request,label+'-owner',config={'pausePhase':'before','action':primary_action(app)})
        self.wait_ready(first)
        try:
            _,duplicate=self.task(app,request,label+'-duplicate',source=first[3]['source'])
            require(duplicate['state']=='running','Duplicate did not observe current owner',result=duplicate)
            _,conflict=self.task(app,request,label+'-competitor')
            require(conflict['outcome']=='failed' and conflict['stop']['error']['code']=='DOCUMENT_LEASE_CONFLICT','Competitor entered leased document',result=conflict)
            require(self.count(app,plan.path,label+'-blocked')==0,'Mutation happened before released barrier')
        finally:(Path(first[3]['directory'])/'release').write_text('release')
        record,result=self.finish(first)
        require(result['outcome']=='succeeded' and result['taskId']==duplicate['taskId'],'Owner failed',result=result)
        require(document_text(plan.path).count(MARK)==1,'Owner effect count')
        self.close_saved(app,plan.path,label+'-end')
        return {'ownerTaskId':result['taskId'],'duplicateTaskId':duplicate['taskId'],'competitorError':conflict['stop']['error']['code'],'effects':1}

    def lease_paths(self,app,label):
        rows=[]
        for mode in ('save','saveAs'):
            plan=self.fixture(app,label+mode)
            request=edit(app,self.root.as_posix(),label,plan.path).request
            target=Path(plan.path)
            if mode=='saveAs':
                target=target.with_name(target.stem+'-saved-as'+target.suffix)
                request['completion']=[{'address':{'app':app,'action':'saveAs'},'params':{'outputPath':str(target),'overwritePolicy':'failIfExists'}}]
            owner=self.launch(app,request,label+mode,config={'pausePhase':'after','action':mode})
            self.wait_ready(owner)
            try:
                require(document_text(target).count(MARK)==1,'Owner persistence not complete before competition')
                competing=edit(app,self.root.as_posix(),label,str(target)).request
                _,blocked=self.task(app,competing,label+mode+'-competitor')
                require(blocked['outcome']=='failed' and blocked['stop']['error']['code']=='DOCUMENT_LEASE_CONFLICT','Persisted target lost its lease fence',result=blocked)
            finally:(Path(owner[3]['directory'])/'release').write_text('release')
            _,result=self.finish(owner)
            require(result['outcome']=='succeeded' and document_text(target).count(MARK)==1,'Owner lost persistence or competitor mutated')
            if mode=='saveAs':require(MARK not in document_text(plan.path),'Save As modified original file')
            self.close_saved(app,target,label+mode+'-end')
            rows.append({'persistence':mode,'competingTarget':str(target),'error':blocked['stop']['error']['code'],'effects':1})
        return rows

    def interruption(self,app,label):
        results=[]
        mutation={'word':'writeContent','excel':'writeRange','ppt':'addTextBox'}[app]
        for window in ('before_effect','after_effect','after_saved_receipt'):
            plan=self.fixture(app,label+window)
            request=edit(app,self.root.as_posix(),label,plan.path).request
            action=mutation;phase='before' if window=='before_effect' else 'after'
            if window=='after_saved_receipt':
                request['completion'].append({'address':{'app':app,'action':'exportPdf'},'params':{'outputPath':str(self.root/'outputs'/(label+window+'.pdf')),'overwritePolicy':'failIfExists'}})
                action='exportPdf';phase='before'
            handle=self.launch(app,request,label+window,config={'pausePhase':phase,'action':action})
            self.wait_ready(handle)
            witness=ProcessWitness(self.native_events(handle[3])[0]['pid'])
            handle[0].kill()
            record,_=self.finish(handle,killed=True)
            witness.assert_exited()
            result=self.query(app,record['source'],Path(record['directory'])/'interrupted')
            require(result['state']=='interrupted' and result['outcome']=='unknown','Lost owner not reported uncertain',result=result)
            effects=self.count(app,plan.path,label+window+'-observed')
            require(effects==(0 if window=='before_effect' else 1),'Unexpected interrupted effects',effects=effects)
            require(result['document']['state']=='succeeded','Known acquisition was lost')
            if window=='after_saved_receipt':
                require(result['completion']['save']['state']=='succeeded' and document_text(plan.path).count(MARK)==1,'Known saved effect was lost')
                require(result['completion']['pdf']['state']=='unknown','Possible export misreported')
            else:
                require(result['completion']['save']['state']=='not_executed' and MARK not in document_text(plan.path),'Unexecuted save occurred')
            _,again=self.task(app,request,label+window+'-repeat',source=record['source'])
            require(again['taskId']==result['taskId'] and again['state']=='interrupted','Interrupted Task replayed')
            require(self.count(app,plan.path,label+window+'-after-repeat')==effects,'Repeated interrupted effects')
            results.append({'window':window,'effects':effects,'stop':result['stop'],'saveState':result['completion']['save']['state'],'taskId':result['taskId'],'bridgePid':witness.pid,'bridgeExited':True})
        return results

    def isolation(self,app,label):
        plan=self.fixture(app,label)
        request=edit(app,self.root.as_posix(),label,plan.path).request
        operation={'word':'insert_structured_body_content','excel':'write_literal_rectangle','ppt':'addTextBox'}[app]
        handle=self.launch(app,request,label+'-timeout',timeout=5,config={'mode':'pause','operation':operation,'stage':'inflight_after'})
        ready=self.wait_ready(handle)
        require(read(ready['coordinationPath'])['inFlight'] is True,'In-flight marker missing')
        witness=ProcessWitness(ready['pid'])
        record,result=self.finish(handle)
        witness.assert_exited()
        require(result['outcome']=='unknown','Deadline did not preserve uncertainty',result=result)
        require(record['seconds']<25,'Deadline and cleanup exceeded declared test bound',seconds=record['seconds'])
        require(result['completion']['save']['state']=='not_executed','Timeout allowed save')
        require(self.count(app,plan.path,label+'-effect')==1,'Pre-timeout effect missing or duplicated')
        write(Path(record['directory'])/'coordination-after.json',read(ready['coordinationPath']))
        _,blocked=self.task(app,request,label+'-reentry')
        require(blocked['outcome']=='failed' and blocked['stop']['error']['code']=='DOCUMENT_QUARANTINED','Unsafe reentry not quarantined',result=blocked)
        other=self.fixture(app,label+'-independent')
        _,independent=self.task(app,edit(app,self.root.as_posix(),label,other.path).request,label+'-independent-edit')
        require(independent['outcome']=='succeeded' and document_text(other.path).count(MARK)==1,'Unrelated document blocked',result=independent)
        self.close_saved(app,other.path,label+'-independent-end')
        return {'effectCount':1,'outcome':result['outcome'],'stop':result['stop'],'seconds':record['seconds'],
                'reentryError':blocked['stop']['error']['code'],'otherDocument':'succeeded',
                'injection':'effect complete, response pending inside coordinated window; not a hung WPS server call',
                'cleanup':result['cleanup'],'bridgePid':ready['pid'],'bridgeExited':True}

    def stale(self,app,label):
        from tests.applications.complex_plan import Case
        from tests.applications.native_acceptance import ref
        from tests.applications.action_suite import paragraph
        plan=self.fixture(app,label)
        c=Case(label,app,'过期观察不得覆盖当前状态',MARK,self.root.as_posix(),plan.path)
        forbidden='FORBIDDEN_STALE_WRITE'
        if app=='word':
            c.add('read','findContent',query={'scope':{'kind':'document'},'text':BASE,'caseSensitive':True,'wholeWord':False},limit=50)
            c.add('change','writeContent',anchor={'kind':'documentEnd'},blocks=[paragraph(MARK)])
            c.add('stale','replaceContent',target={'kind':'range','range':ref('read','matches',0,'range')},replacement={'kind':'blocks','blocks':[{'kind':'text','runs':[{'text':forbidden}]}]})
            code='STALE_CONTENT_RANGE'
        elif app=='excel':
            c.add('read','readRange',sheet=SHEET,address='A1')
            c.add('change','writeRange',sheet=SHEET,address='A1',expectedToken=ref('read','token'),values=[[MARK]])
            c.add('stale','writeRange',sheet=SHEET,address='A1',expectedToken=ref('read','token'),values=[[forbidden]])
            code='STALE_RANGE'
        else:
            c.add('slides','listSlides');sid=ref('slides','slides',0,'id')
            c.add('read','getSlideInfo',slideId=sid);shape=ref('read','shapes',0,'id')
            c.add('change','setShapeText',slideId=sid,shapeId=shape,expectedToken=ref('read','token'),text=MARK)
            c.add('stale','setShapeText',slideId=sid,shapeId=shape,expectedToken=ref('read','token'),text=forbidden)
            code='STALE_CONTENT'
        c.finish(save=True)
        _,result=self.task(app,c.request,label)
        require(result['outcome']=='failed' and result['stop']['error']['code']==code,'Stale observation not rejected',result=result)
        require(result['completion']['save']['state']=='not_executed','Stale failure allowed save')
        observed=self.observe(app,plan.path,label+'-effect');text='\n'.join(str(x or '') for x in observed['texts'])
        require(text.count(MARK)==1 and forbidden not in text,'Stale consumer changed live state',text=text)
        require(MARK not in document_text(plan.path),'Stale failure saved document')
        return {'error':code,'earlierEffects':1,'staleEffects':0,'diskEffects':0}

    def response_loss(self,app,label):
        plan=self.fixture(app,label)
        operation={'word':'insert_structured_body_content','excel':'write_literal_rectangle','ppt':'addTextBox'}[app]
        request=edit(app,self.root.as_posix(),label,plan.path).request
        record,result=self.task(app,request,label,config={'mode':'drop_response','stage':'after','operation':operation})
        require(any(e['event']=='injection' for e in self.native_events(record)),'Response-loss injection missing')
        require(result['outcome']=='unknown' and result['stop']['error']['code']=='RESPONSE_LOST','Lost native response falsely certain',result=result)
        require(result['completion']['save']['state']=='not_executed' and self.count(app,plan.path,label+'-effect')==1,'Lost response had wrong effects')
        require(MARK not in document_text(plan.path),'Unknown Task saved')
        _,repeat=self.task(app,request,label+'-repeat',source=record['source'])
        require(repeat['taskId']==result['taskId'] and repeat['outcome']=='unknown','Lost response replayed')
        require(self.count(app,plan.path,label+'-again')==1,'Lost response duplicate effects')
        return {'outcome':result['outcome'],'error':result['stop']['error']['code'],'liveEffects':1,'diskEffects':0,'cleanup':result['cleanup']}

    def cross_task(self,app,label):
        from tests.applications.complex_plan import Case
        plan=self.fixture(app,label)
        previous=read(Path(self.tasks[-1]['directory'])/'response.json')
        data=previous['steps'][-1]['response']['data']
        c=Case(label,app,'跨 Task 观察不可复用',MARK,self.root.as_posix(),plan.path)
        if app=='word':
            c.add('foreign','replaceContent',target={'kind':'range','range':data['range']},replacement={'kind':'blocks','blocks':[{'kind':'text','runs':[{'text':MARK}]}]})
            code='STALE_CONTENT_RANGE'
        elif app=='excel':
            c.add('foreign','writeRange',sheet=SHEET,address='A1',expectedToken=data['token'],values=[[MARK]])
            code='STALE_RANGE'
        else:
            c.add('foreign','setShapeText',slideId=data['slide']['id'],shapeId=data['shapes'][0]['id'],expectedToken=data['token'],text=MARK)
            code='STALE_CONTENT'
        c.finish(save=True)
        _,result=self.task(app,c.request,label)
        require(result['document']['state']=='succeeded' and result['outcome']=='failed' and result['stop']['error']['code']==code,'Foreign Task observation not rejected',result=result)
        require(result['completion']['save']['state']=='not_executed','Foreign observation allowed save')
        require(self.count(app,plan.path,label+'-effect')==0 and MARK not in document_text(plan.path),'Foreign observation changed document')
        self.close_saved(app,plan.path,label+'-end')
        return {'oldTaskId':previous['taskId'],'newTaskId':result['taskId'],'error':code,'effects':0}

    def receipts(self,app,label):
        from wps_skills.client.task_request import result_steps
        rows=[]
        for candidate in (False,True):
            plan=self.fixture(app,label+str(candidate))
            request=edit(app,self.root.as_posix(),label,plan.path).request
            request['completion'].append({'address':{'app':app,'action':'exportPdf'},'params':{
                'outputPath':str(self.root/'outputs'/(label+str(candidate)+'.pdf')),'overwritePolicy':'failIfExists'}})
            handle=self.launch(app,request,label+str(candidate),config={'pausePhase':'before','action':'exportPdf','finalReceiptOnly':candidate})
            self.wait_ready(handle)
            witness=ProcessWitness(self.native_events(handle[3])[0]['pid'])
            require(document_text(plan.path).count(MARK)==1,'Save did not finish before interruption')
            handle[0].kill();record,_=self.finish(handle,killed=True);witness.assert_exited()
            result=self.query(app,record['source'],Path(record['directory'])/'recovered')
            require(result['state']=='interrupted' and result['outcome']=='unknown','Interrupted receipt falsely complete',result=result)
            known=sum(s['state']=='succeeded' for s in result_steps(result))
            unknown=sum(s['state']=='unknown' for s in result_steps(result))
            if candidate:
                require(known==0 and all(s['state']=='unknown' for s in result_steps(result)),'Final-only candidate not active',result=result)
                events=[json.loads(line) for line in (Path(record['directory'])/'child-events.jsonl').read_text(encoding='utf-8-sig').splitlines()]
                require(any(e['event']=='receipt_suppressed' for e in events),'Intermediate receipt was not suppressed')
            else:require(result['completion']['save']['state']=='succeeded' and known>0,'Progress receipts lost known save')
            _,again=self.task(app,request,label+str(candidate)+'-repeat',source=record['source'])
            require(again['taskId']==result['taskId'] and document_text(plan.path).count(MARK)==1,'Receipt comparison replayed')
            rows.append({'finalReceiptOnly':candidate,'knownSucceeded':known,'unknown':unknown,'saveState':result['completion']['save']['state'],'diskEffects':1,'bridgeExited':True})
        return rows

    def persistence(self,app,label):
        results=[]
        for mode in ('no_save','pdf_only','dirty_existing','save_conflict','pdf_conflict','cleanup_failure'):
            plan=self.fixture(app,label+mode)
            before=sha(plan.path)
            if mode=='dirty_existing':self.observe(app,plan.path,label+mode+'-dirty',mode='dirty')
            request=edit(app,self.root.as_posix(),label,plan.path,save=mode not in ('no_save','pdf_only')).request
            output=self.root/'outputs'/(label+mode+'.pdf')
            if mode in ('pdf_only','save_conflict','pdf_conflict'):
                request['completion'].append({'address':{'app':app,'action':'exportPdf'},'params':{'outputPath':str(output),'overwritePolicy':'failIfExists'}})
            if mode=='save_conflict':
                occupied=self.root/'outputs'/(label+'-occupied'+Path(plan.path).suffix);occupied.write_bytes(b'occupied-output')
                request['completion'][0]={'address':{'app':app,'action':'saveAs'},'params':{'outputPath':str(occupied),'overwritePolicy':'failIfExists'}}
            if mode=='pdf_conflict':output.write_bytes(b'occupied-output')
            record,result=self.task(app,request,label+mode,config={'cleanupFailure':mode=='cleanup_failure'})
            expected='failed' if mode in ('dirty_existing','save_conflict','pdf_conflict') else 'succeeded'
            require(result['outcome']==expected,'Persistence outcome mismatch',mode=mode,result=result)
            effects=self.count(app,plan.path,label+mode+'-effect')
            require(effects==(0 if mode=='dirty_existing' else 1),'Persistence live effect mismatch',mode=mode,effects=effects)
            persisted=mode in ('pdf_conflict','cleanup_failure')
            require(document_text(plan.path).count(MARK)==(1 if persisted else 0),'Unexpected disk effect',mode=mode)
            if not persisted:require(sha(plan.path)==before,'Original file unexpectedly changed',mode=mode)
            if mode=='dirty_existing':require(result['stop']['error']['code']=='TASK_EXISTING_CHANGES_CONFIRMATION_REQUIRED','Dirty document not protected',result=result)
            if mode=='save_conflict':require(occupied.read_bytes()==b'occupied-output' and not output.exists(),'Failed save overwrote output or allowed PDF')
            if mode=='pdf_conflict':require(result['completion']['save']['state']=='succeeded' and output.read_bytes()==b'occupied-output','Saved result lost or PDF overwritten')
            if mode=='pdf_only':require(output.read_bytes().startswith(b'%PDF-') and result['completion']['save'] is None,'Export-only failed')
            if mode=='cleanup_failure':
                require(result['cleanup']['outcome']=='failed' and record['exitCode']!=0,'Cleanup failure not separate')
                _,again=self.task(app,request,label+mode+'-repeat',source=record['source'])
                require(again['taskId']==result['taskId'] and self.count(app,plan.path,label+mode+'-again')==1,'Cleanup failure caused replay')
                self.close_saved(app,plan.path,label+mode+'-end')
            results.append({'mode':mode,'outcome':result['outcome'],'liveEffects':effects,'diskEffects':int(persisted),'cleanup':result['cleanup']['outcome'],'stop':result['stop']})
        return results

    def dynamic(self,app,label):
        from tests.applications.complex_plan import Case
        from tests.applications.native_acceptance import ref
        from tests.applications.action_suite import INSPECT,paragraph
        results=[]
        for bad_path in (False,True):
            plan=self.fixture(app,label+str(bad_path))
            c=Case(label,app,'动态非法参数',MARK,self.root.as_posix(),plan.path)
            field='missing_field' if bad_path else {'word':'truncated','excel':'cells','ppt':'token'}[app]
            if app=='word':
                c.add('read','inspectDocument',**INSPECT)
                c.add('invalid','writeContent',anchor={'kind':'documentEnd'},blocks=[paragraph(ref('read',field))])
            elif app=='excel':
                c.add('read','readRange',sheet=SHEET,address='A1')
                c.add('invalid','writeRange',sheet=SHEET,address='A1',expectedToken=ref('read','token'),values=ref('read',field))
            else:
                c.add('read','listSlides')
                c.add('invalid','addSlide',position=ref('read',field),expectedToken=ref('read','token'))
            c.finish(save=True)
            _,result=self.task(app,c.request,label+str(bad_path))
            code='TASK_REFERENCE_UNAVAILABLE' if bad_path else 'TASK_PARAMS_INVALID'
            require(result['outcome']=='failed' and result['stop']['error']['code']==code,'Wrong dynamic rejection',result=result)
            require(result['steps'][1]['state']=='not_executed' and result['completion']['save']['state']=='not_executed','Invalid consumer executed')
            require(self.count(app,plan.path,label+str(bad_path))==0,'Invalid consumer had effects')
            self.close_saved(app,plan.path,label+str(bad_path)+'-end')
            results.append({'badReferencePath':bad_path,'error':code,'consumerDispatched':False})
        return results

    def references(self,app,label):
        from tests.applications.complex_plan import Case
        from tests.applications.native_acceptance import ref
        plan=self.fixture(app,label)
        c=Case(label,app,'真实结果引用贯通',MARK,self.root.as_posix(),plan.path)
        expected_count=1
        if app=='word':
            for index,(before,after) in enumerate(((BASE,MARK),(MARK,MARK+'_FINAL'))):
                step='find_'+str(index)
                c.add(step,'findContent',query={'scope':{'kind':'document'},'text':before,'caseSensitive':True,'wholeWord':False},limit=50)
                c.add('replace_'+str(index),'replaceContent',target={'kind':'range','range':ref(step,'matches',0,'range')},replacement={'kind':'blocks','blocks':[{'kind':'text','runs':[{'text':after}]}]})
        elif app=='excel':
            c.add('sheets','listWorksheets',offset=0,limit=100);sheet=ref('sheets','worksheets',0,'name')
            c.add('source','readRange',sheet=sheet,address='A1')
            c.add('target','readRange',sheet=sheet,address='A2')
            c.add('copy_value','writeRange',sheet=sheet,address='A2',expectedToken=ref('target','token'),values=[[ref('source','cells',0,0,'value')]])
            c.add('mark_token','readRange',sheet=sheet,address='A3')
            c.add('mark','writeRange',sheet=sheet,address='A3',expectedToken=ref('mark_token','token'),values=[[MARK]])
        else:
            c.add('slides','listSlides');slide=ref('slides','slides',0,'id')
            c.add('shape','getSlideInfo',slideId=slide)
            c.add('text','setShapeText',slideId=slide,shapeId=ref('shape','shapes',0,'id'),expectedToken=ref('shape','token'),text=MARK)
            c.add('changed','getSlideInfo',slideId=slide)
            c.add('copy_text','addTextBox',slideId=slide,expectedToken=ref('changed','token'),text=ref('changed','shapes',0,'text'),left=40,top=160,width=500,height=60)
            expected_count=2
        c.finish(save=True)
        record,result=self.task(app,c.request,label)
        require(result['outcome']=='succeeded','Dynamic reference chain failed',result=result)
        require(self.count(app,plan.path,label+'-live')==expected_count,'Reference chain live effect mismatch')
        text=document_text(plan.path)
        require(text.count(MARK)==expected_count,'Reference chain saved effect mismatch',text=text)
        if app=='word':require(MARK+'_FINAL' in text and BASE not in text,'Word dynamic range replacement failed')
        if app=='excel':require(text.count(BASE)==2,'Excel dynamic source value was not copied')
        self.close_saved(app,plan.path,label+'-end')
        return {'contentActions':len(c.request['steps']),'taskId':result['taskId'],'markerEffects':expected_count,'seconds':record['seconds'],'callerRefilledParameters':False}

    def readback(self,app,label):
        if app!='excel':return {'state':'not_applicable','reason':'Controlled native write/readback comparison uses Excel; no claim for other Actions.'}
        results=[]
        for optimistic in (False,True):
            plan=self.fixture(app,label+str(optimistic))
            from tests.applications.complex_plan import Case
            c=Case(label,app,'回读对照',MARK,self.root.as_posix(),plan.path)
            c.region('write','writeRange',SHEET,'A1',values=[[MARK]]);c.finish(save=True)
            record,result=self.task(app,c.request,label+str(optimistic),config={'noWrite':True,'optimisticReadback':optimistic})
            require(any(e['event']=='write_suppressed' for e in self.native_events(record)),'Write fault not injected')
            require(any(e['event']=='optimistic_result' for e in self.native_events(record))==optimistic,'Readback candidate activation mismatch')
            require(self.count(app,plan.path,label+str(optimistic))==0 and MARK not in document_text(plan.path),'Suppressed write unexpectedly took effect')
            require((result['outcome']=='succeeded')==optimistic,'Readback result mismatch',result=result)
            if not optimistic:require(result['completion']['save']['state']=='not_executed','Failed verification allowed save')
            results.append({'optimistic':optimistic,'actualEffects':0,'outcome':result['outcome'],'seconds':record['seconds'],'stop':result['stop']})
            self.close_saved(app,plan.path,label+str(optimistic)+'-end')
        return results

    def schema(self,app,label):
        from .measurements import schema
        return schema(self,app,label)

    def protocol(self,app,label):
        from .protocol import protocol
        return protocol(self,app,label)

    def startup(self,app,label):
        from .protocol import startup
        return startup(self,app,label)

    def task_cost(self,app,label):
        from .measurements import task_cost
        return task_cost(self,app,label)

    def readback_cost(self,app,label):
        from .measurements import readback_cost
        return readback_cost(self,app,label)


def main(kit):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root',required=True,type=Path)
    parser.add_argument('--groups',nargs='+',choices=GROUPS,default=['preflight','replay','focus','readback'])
    parser.add_argument('--apps',nargs='+',default=['word','excel','ppt'],choices=['word','excel','ppt'])
    parser.add_argument('--trials',type=int,default=3)
    args=parser.parse_args()
    require(os.name=='nt','Windows required')
    import ctypes
    session=ctypes.c_ulong()
    require(ctypes.windll.kernel32.ProcessIdToSessionId(os.getpid(),ctypes.byref(session)) and session.value!=0,'Interactive Windows session required')
    require(1<=args.trials<=30,'Trial count must be 1..30')
    root=args.root.resolve();root.mkdir(parents=True,exist_ok=False);(root/'outputs').mkdir()
    manifest=read(kit/'manifest.json')
    for name,digest in manifest['files'].items():require(sha(kit/name)==digest,'Changed test package',path=name)
    write(root/'plan.json',{'groups':args.groups,'apps':args.apps,'trials':args.trials,'kitManifestSha256':sha(kit/'manifest.json'),'document':'PLAN.md' if (kit/'PLAN.md').is_file() else None})
    write(root/'environment.json',{'python':sys.version,'executable':sys.executable,'sessionId':session.value,'pid':os.getpid()})
    experiment=Experiments(kit,root)
    report={'state':'running','startedUtc':datetime.now(timezone.utc).isoformat(),'experiments':[]}
    try:
        for group in args.groups:
            for app in args.apps:
                for trial in range(1,args.trials+1):
                    label=f'{group}-{app}-{trial}'
                    row={'id':label,'state':'running'};report['experiments'].append(row);write(root/'report.json',report)
                    print('RUN '+label,flush=True)
                    try:
                        observed=getattr(experiment,group)(app,label)
                        state='not_applicable' if isinstance(observed,dict) and observed.get('state')=='not_applicable' else 'passed'
                        row.update(state=state,observations=observed)
                    except BaseException as error:
                        row.update(state='failed',error=str(error));raise
                    finally:write(root/'report.json',report)
                    print(('SKIP ' if row['state']=='not_applicable' else 'PASS ')+label,flush=True)
        report['state']='passed'
    except BaseException as error:
        report.update(state='failed',error=str(error));traceback.print_exc()
    finally:
        report['endedUtc']=datetime.now(timezone.utc).isoformat()
        write(root/'report.json',report);write(root/'calls.json',experiment.tasks)
        print('Report: '+str(root/'report.json'),flush=True)
    return 0 if report['state']=='passed' else 1
