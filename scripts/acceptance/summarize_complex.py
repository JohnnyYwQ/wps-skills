"""Offline summary of frozen complex-suite manifests, responses and raw trace timing."""
import argparse
import csv
import hashlib
import json
from pathlib import Path
import statistics


def read(p):return json.loads(p.read_text(encoding='utf-8-sig'))
def stat(values):
    return {'count':len(values),'sumSeconds':sum(values),'medianSeconds':statistics.median(values),'minSeconds':min(values),'maxSeconds':max(values)} if values else {'count':0}

def summarize(root):
    summaries={};tasks=[];actions=[]
    for app in ('word','excel','ppt'):
        base=root/app;manifest=read(base/'manifest.json')
        report=read(base/'report.json') if (base/'report.json').exists() else {'state':'not_run','cases':[]}
        actual={c['id']:c for c in report['cases']};traces={}
        for path in (base/'traces/requests').glob('*.jsonl'):
            entries=[json.loads(l) for l in path.read_text(encoding='utf-8-sig').splitlines() if l.strip()]
            action_spans=[e for e in entries if e.get('name')=='action.execute' and e.get('event')=='span.finished']
            if not action_spans:continue
            tids={e['taskId'] for e in action_spans};assert len(tids)==1
            tid=tids.pop();assert tid not in traces
            total=[e for e in entries if e.get('name')=='request.total' and e.get('event')=='span.finished']
            traces[tid]=(total,action_spans)
        business={'passed':0,'failed':0,'blocked':0,'running':0,'not_run':0};setup={'passed':0,'failed':0,'blocked':0,'running':0,'not_run':0}
        times=[];action_times=[]
        for spec in manifest['cases']:
            record=actual.get(spec['id'],{'state':'not_run'});status=record['state']
            (setup if spec.get('setupOnly') else business)[status]+=1
            rp=base/'results'/spec['id']/'response.json'
            r=read(rp) if rp.exists() else {}
            row={'app':app,'caseId':spec['id'],'title':spec['title'],'setupOnly':spec.get('setupOnly',False),'status':status,
                 'taskId':r.get('taskId'),'taskOutcome':r.get('outcome'),'cleanupOutcome':(r.get('cleanup') or {}).get('outcome'),
                 'plannedActions':spec['actionCount'],'taskSeconds':None,'stop':r.get('stop')}
            if r.get('taskId') in traces:
                totals,spans=traces[r['taskId']]
                if totals:assert len(totals)==1;row['taskSeconds']=totals[0]['durationNs']/1e9
                row['observedActions']=len(spans)
                for e in spans:
                    ar={'app':app,'caseId':spec['id'],'setupOnly':row['setupOnly'],'taskId':r['taskId'],'stepId':e['stepId'],'action':e['action'],'outcome':e.get('outcome'),'seconds':e['durationNs']/1e9};actions.append(ar)
                row['slowestActions']=sorted([{'stepId':e['stepId'],'action':e['action'],'seconds':e['durationNs']/1e9} for e in spans],key=lambda x:x['seconds'],reverse=True)[:3]
                if status=='passed':
                    assert len(totals)==1 and len(spans)==spec['actionCount']
                    assert all(e.get('outcome')=='succeeded' for e in spans)
                    if not row['setupOnly']:times.append(row['taskSeconds']);action_times.extend(e['durationNs']/1e9 for e in spans)
            elif status=='passed':raise ValueError('Missing trace for passed case '+spec['id'])
            tasks.append(row)
        summaries[app]={'state':report['state'],'business':business,'setup':setup,'successfulBusinessTaskTiming':stat(times),'successfulBusinessActionTiming':stat(action_times)}
    return {'scope':'execution-side fixed JSON, one attempt per case; successful business timing excludes setup and failed assertions','applications':summaries,'tasks':tasks,'actions':actions}


def main():
    p=argparse.ArgumentParser();p.add_argument('--root',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
    s=summarize(a.root);a.output.mkdir(parents=True,exist_ok=True)
    (a.output/'summary.json').write_text(json.dumps(s,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    for name in ('tasks','actions'):
        rows=s[name];keys=list(dict.fromkeys(k for row in rows for k in row))
        with (a.output/(name+'.csv')).open('w',encoding='utf-8-sig',newline='') as f:
            w=csv.DictWriter(f,fieldnames=keys);w.writeheader();w.writerows(rows)
    hashes={p.relative_to(a.root).as_posix():hashlib.sha256(p.read_bytes()).hexdigest() for p in a.root.rglob('*') if p.is_file()}
    (a.output/'source-sha256.json').write_text(json.dumps(hashes,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(s['applications'],ensure_ascii=False,indent=2))
if __name__=='__main__':main()
