"""Compare a recorded Task Response with a prepared oracle; never call WPS."""
import argparse
import json
from pathlib import Path

from wps_skills.client.task_request import result_steps
from wps_skills.core.action_runtime import _json_equal


def select(value, path, responses):
    if not path:return value
    head,*tail=path
    if head=='*':
        if not isinstance(value,list) or not value:raise ValueError('Wildcard requires a nonempty observed array')
        return [select(v,tail,responses) for v in value]
    if isinstance(head,dict):
        wanted=resolve_expected(head['value'],responses)
        matches=[v for v in value if _json_equal(v.get(head['field']),wanted)]
        if len(matches)!=1:raise ValueError('Selector must identify exactly one observed item')
        return select(matches[0],tail,responses)
    return select(value[head],tail,responses)


def resolve_expected(value,responses):
    if isinstance(value,dict):
        if set(value)=={'$ref'}:
            r=value['$ref'];return select(responses[r['step']],r['path'],responses)
        if set(value)=={'$response'}:
            r=value['$response'];return select(responses[r['step']]['data'],r['path'],responses)
        return {k:resolve_expected(v,responses) for k,v in value.items()}
    if isinstance(value,list):return [resolve_expected(v,responses) for v in value]
    return value


def flatten(value):
    if isinstance(value,list):
        for v in value:yield from flatten(v)
    else:yield value


def matches(actual,rule,responses):
    wanted=resolve_expected(rule.get('value'),responses)
    op=rule['op']
    if op=='equals':return _json_equal(actual,wanted)
    if op=='one_of':return any(_json_equal(actual,item) for item in wanted)
    if op=='not_null':return actual is not None
    if op=='greater_than':return isinstance(actual,(int,float)) and not isinstance(actual,bool) and actual>wanted
    if op=='length':return len(actual)==wanted
    if op=='not_contains':return not any(_json_equal(v,wanted) for v in actual)
    if op=='all_equals':return all(_json_equal(v,wanted) for v in flatten(actual))
    if op in {'close','all_close'}:
        values=list(flatten(actual)) if op=='all_close' else [actual]
        return all(isinstance(v,(int,float)) and not isinstance(v,bool) and abs(v-wanted)<=rule['tolerance'] for v in values)
    raise ValueError('Unknown expectation operator: '+op)


def compare(expected,response):
    """Return exact mismatches; missing or unexpected records never silently pass."""
    errors=[]
    def require(label,actual,wanted):
        if not _json_equal(actual,wanted):errors.append({'field':label,'expected':wanted,'actual':actual})
    try:
        task=expected['task']
        require('type',response.get('type'),'task.response')
        for k in ('app','state','outcome'):require(k,response.get(k),task[k])
        if task['stop'] is None:require('stop',response.get('stop'),None)
        else:
            stop=response.get('stop') or {}
            for k in ('phase','stepId'):require('stop.'+k,stop.get(k),task['stop'][k])
            require('stop.error.code',(stop.get('error') or {}).get('code'),task['stop']['error']['code'])
        require('cleanup.outcome',(response.get('cleanup') or {}).get('outcome'),task['cleanupOutcome'])
        # taskFile exists on the initial CLI response but is not in the stored receipt.
        if 'taskFile' in response:require('taskFile.state',response['taskFile'].get('state'),task['taskFileState'])
        if not isinstance(response.get('taskId'),str) or not response['taskId']:errors.append({'field':'taskId','error':'missing identity'})
        if not isinstance(response.get('recordPath'),str) or not response['recordPath']:errors.append({'field':'recordPath','error':'missing receipt locator'})
        actual_steps=result_steps(response)
        require('ordered step IDs',[s['id'] for s in actual_steps],[s['id'] for s in expected['actions']])
        by_id={s['id']:s for s in actual_steps}
        responses={s['id']:s['response'] for s in actual_steps if s.get('response') is not None}
        for item in expected['actions']:
            identifier=item['id']
            if identifier not in by_id:errors.append({'stepId':identifier,'error':'missing record'});continue
            s=by_id[identifier];prefix=identifier+'.'
            require(prefix+'address',s.get('address'),item['address']);require(prefix+'state',s.get('state'),item['state'])
            if item['state']=='not_executed':require(prefix+'response',s.get('response'),None);continue
            r=s.get('response') or {}
            require(prefix+'response.outcome',r.get('outcome'),item['responseOutcome'])
            require(prefix+'response.address',r.get('address'),item['address'])
            require(prefix+'response.taskId',r.get('taskId'),response.get('taskId'))
            if not r.get('traceId'):errors.append({'stepId':identifier,'error':'missing traceId'})
            if 'errorCode' in item:require(prefix+'error.code',(r.get('error') or {}).get('code'),item['errorCode'])
            for rule in item['dataAssertions']:
                try:
                    actual=select(r['data'],rule['path'],responses)
                    if not matches(actual,rule,responses):errors.append({'stepId':identifier,'rule':rule,'actual':actual})
                except (KeyError,IndexError,TypeError,ValueError) as exc:
                    errors.append({'stepId':identifier,'rule':rule,'error':str(exc)})
    except (KeyError,IndexError,TypeError,ValueError) as exc:
        errors.append({'field':'response','error':str(exc)})
    return {'caseId':expected['caseId'],'passed':not errors,'mismatches':errors}


def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--expected',required=True,type=Path);p.add_argument('--response',required=True,type=Path)
    args=p.parse_args(argv)
    result=compare(json.loads(args.expected.read_text(encoding='utf-8-sig')),json.loads(args.response.read_text(encoding='utf-8-sig')))
    print(json.dumps(result,ensure_ascii=False,indent=2))
    return 0 if result['passed'] else 1


if __name__=='__main__':raise SystemExit(main())
