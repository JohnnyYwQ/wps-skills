"""Host-independent schema size and execution-cost measurements with saved raw samples."""
import json
from pathlib import Path
import subprocess
import sys
import time


def schema(experiment,app,label):
    from .runner import require,write
    skill=experiment.kit/'skills'/('wps-'+app)
    entry=skill/'scripts/schema.py'
    selected={'word':['openDocument','findContent','replaceContent','save'],
              'excel':['openWorkbook','readRange','writeRange','save'],
              'ppt':['openPresentation','getSlideInfo','setShapeText','save']}[app]
    all_names=sorted(p.stem for p in (skill/'references/schemas/actions').glob('*.json'))
    rows=[]
    for trial in range(10):
        observed={}
        for mode in (('selected','all') if trial%2==0 else ('all','selected')):
            names=selected if mode=='selected' else all_names
            start=time.perf_counter()
            cp=subprocess.run([sys.executable,str(entry),*names],capture_output=True,timeout=30)
            elapsed=time.perf_counter()-start
            require(cp.returncode==0,'Schema query failed',stderr=cp.stderr.decode('utf-8'))
            data=json.loads(cp.stdout)
            observed[mode]=data
            rows.append({'trial':trial,'mode':mode,'seconds':elapsed,'bytes':len(cp.stdout),'actionCount':len(data['actions'])})
            if trial==0:
                output=experiment.root/'measurements'/label/(mode+'.json');output.parent.mkdir(parents=True,exist_ok=True);output.write_bytes(cp.stdout)
        require(all(observed['all']['actions'][k]==v for k,v in observed['selected']['actions'].items()),'Selected Action schema differs')
        require(all(observed['all']['$defs'][k]==v for k,v in observed['selected']['$defs'].items()),'Selected shared definition differs')
    write(experiment.root/'measurements'/(label+'.json'),rows)
    return {'selected':selected,'allActionCount':len(all_names),'pairs':10,'samples':rows,'tokens':'not measured','agentBehavior':'not measured'}


def task_cost(experiment,app,label):
    from .runner import require,document_text,read
    from .plans import BASE,MARK,SHEET
    from tests.applications.complex_plan import Case
    from tests.applications.action_suite import INSPECT
    plan=experiment.fixture(app,label)
    rows=[]
    # Rotate order across outer trials to reduce monotonic warm-up/order bias.
    counts=[1,8,32];rotation=(int(label.rsplit('-',1)[-1])-1)%3;counts=counts[rotation:]+counts[:rotation]
    for count in counts:
        c=Case(label,app,'计划长度成本',BASE,experiment.root.as_posix(),plan.path)
        for i in range(count):
            if app=='word':c.add('read_'+str(i),'inspectDocument',**INSPECT)
            elif app=='excel':c.add('read_'+str(i),'readRange',sheet=SHEET,address='A1')
            else:c.add('read_'+str(i),'listSlides')
        record,result=experiment.task(app,c.request,label+'-'+str(count))
        require(result['outcome']=='succeeded' and all(s['state']=='succeeded' for s in result['steps']),'Read chain failed',result=result)
        pids={e['pid'] for e in experiment.native_events(record)}
        require(len(pids)==1,'Task did not reuse one owned bridge',pids=list(pids))
        require(BASE in document_text(plan.path) and MARK not in document_text(plan.path),'Read chain changed disk')
        require(experiment.count(app,plan.path,label+'-'+str(count))==0,'Read chain changed live document')
        totals={}
        for path in (experiment.root/'traces/requests').glob('*.jsonl'):
            events=[json.loads(line) for line in path.read_text(encoding='utf-8-sig').splitlines()]
            if not any(e.get('taskId')==result['taskId'] for e in events):continue
            for event in events:
                if event.get('event')=='span.finished':
                    name=event['name'];totals[name]=totals.get(name,0)+event['durationMs']
        require('task.receipt_publish' in totals and 'action.execute' in totals,'Missing phase evidence')
        rows.append({'contentActions':count,'wallSeconds':record['seconds'],'taskId':result['taskId'],'phaseMs':totals,'bridgeProcesses':len(pids),
                     'phaseWarning':'Nested inclusive spans; do not sum all phase durations.'})
    experiment.close_saved(app,plan.path,label+'-end')
    return rows


def readback_cost(experiment,app,label):
    from .runner import require,document_text
    from .plans import MARK,SHEET
    from tests.applications.complex_plan import Case
    if app!='excel':return {'state':'not_applicable'}
    plan=experiment.fixture(app,label)
    rows=[]
    for pair in range(8):
        for optimistic in ((False,True) if pair%2==0 else (True,False)):
            marker=MARK+'_'+str(pair)+'_'+str(optimistic)
            c=Case(label,app,'正常写入回读成本',marker,experiment.root.as_posix(),plan.path)
            c.region('write','writeRange',SHEET,'A1',values=[[marker]]);c.finish(save=True)
            record,result=experiment.task(app,c.request,label+'-'+str(pair)+'-'+str(optimistic),config={'optimisticReadback':optimistic})
            require(result['outcome']=='succeeded' and marker in document_text(plan.path),'Measured write is incorrect',result=result)
            observed=experiment.observe(app,plan.path,label+'-'+str(pair)+'-'+str(optimistic))
            require(observed['texts'][0]==marker,'Measured write does not match live WPS')
            if optimistic:require(any(e['event']=='optimistic_result' for e in experiment.native_events(record)),'Readback candidate not active')
            rows.append({'pair':pair,'optimistic':optimistic,'seconds':record['seconds'],'taskId':result['taskId']})
    experiment.close_saved(app,plan.path,label+'-end')
    return rows
