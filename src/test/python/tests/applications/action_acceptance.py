"""Prepare/run a complete Action baseline through standalone Task CLI files."""
import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path, PureWindowsPath
import shutil
import subprocess
import sys
import time
import uuid
import zipfile
from xml.etree import ElementTree as ET


def write(path, value):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    temporary=path.with_suffix(path.suffix+'.tmp')
    temporary.write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    os.replace(temporary,path)


def read(path):return json.loads(Path(path).read_text(encoding='utf-8-sig'))
def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def verify_skill(skill):
    files=read(skill/'runtime/files.sha256.json')
    for name,expected in files.items():
        path=(skill/name).resolve()
        if skill.resolve() not in path.parents or not path.is_file() or sha(path)!=expected:
            raise ValueError('Skill hash mismatch or invalid manifest path: '+name)
    return sha(skill/'runtime/files.sha256.json')


def prepare(app, root, skill, windows_root=None):
    from tests.applications.action_suite import build_cases, coverage, invalid_cases
    from tests.applications.native_acceptance import make_png
    if root.exists():raise FileExistsError('Use a new run directory: '+str(root))
    target=windows_root or str(root)
    if not PureWindowsPath(target).is_absolute():raise ValueError('Preparation requires an absolute Windows run path (use --windows-root on Mac).')
    target=PureWindowsPath(target).as_posix().rstrip('/')
    cases=build_cases(app,target);matrix=coverage(app,cases);negatives=invalid_cases(app,cases)
    package_hash=verify_skill(skill)
    root.mkdir(parents=True)
    (root/'assets').mkdir();(root/'outputs').mkdir()
    make_png(root/'assets/sample.png')
    rows=[]
    # Validate invalid inputs first: no WPS document should be acquired.
    for action,request in negatives.items():
        identifier='invalid-'+action
        expected={'caseId':identifier,'kind':'preflight_rejection','action':action,
                  'task':{'state':'rejected','outcome':'failed','phase':'validation','code':'INVALID_TASK_REQUEST'}}
        rows.append((identifier,request,expected))
    rows += [(c.id,c.request,c.expectation()) for c in cases]
    manifest={'schemaVersion':1,'app':app,'windowsRoot':target,'skill':str(skill.resolve()),
              'skillManifestSha256':package_hash,'createdUtc':datetime.now(timezone.utc).isoformat(),
              'scope':'all admitted Action names; scenario baseline, not exhaustive parameter or fault coverage',
              'coverage':matrix,'cases':[],'state':'prepared','visualReview':'pending'}
    for identifier,request,expected in rows:
        rp='requests/'+identifier+'.json';ep='expected/'+identifier+'.json'
        write(root/rp,request);write(root/ep,expected)
        manifest['cases'].append({'id':identifier,'request':rp,'expected':ep,'requestSha256':sha(root/rp),'expectedSha256':sha(root/ep)})
    manifest['assets']={'assets/sample.png':sha(root/'assets/sample.png')}
    write(root/'manifest.json',manifest)
    return manifest


def parse_response(text):
    rows=[]
    for line in text.splitlines():
        try:row=json.loads(line)
        except ValueError:continue
        if isinstance(row,dict) and row.get('type')=='task.response':rows.append(row)
    if len(rows)!=1:raise ValueError('Expected exactly one Task Response, observed '+str(len(rows)))
    return rows[0]


def evaluate(expected, request, response, exit_code):
    from wps_skills.client.applications import contracts_for, compile_request
    from wps_skills.client.task_request import result_steps, resolve
    from tests.applications.response_expectations import compare
    errors=[];verified=[]
    if expected['kind']=='preflight_rejection':
        checks={'state':'rejected','outcome':'failed','taskId':None,'recordPath':None}
        for key,value in checks.items():
            if response.get(key)!=value:errors.append({'field':key,'expected':value,'actual':response.get(key)})
        stop=response.get('stop') or {}
        if stop.get('phase')!='validation' or (stop.get('error') or {}).get('code')!='INVALID_TASK_REQUEST':errors.append({'field':'stop','actual':stop})
        if any(s.get('state')!='not_executed' or s.get('response') is not None for s in response.get('steps',[])):errors.append({'field':'steps','error':'rejected input had dispatched steps'})
        if exit_code==0:errors.append({'field':'exitCode','error':'rejection reported zero exit'})
        if (response.get('taskFile') or {}).get('state')!='retained':errors.append({'field':'taskFile','error':'unadmitted input must be retained'})
        return {'passed':not errors,'mismatches':errors,'responseVerifiedActions':[]}
    errors.extend(compare(expected,response)['mismatches'])
    if exit_code != (0 if expected['kind']=='positive' else 2):errors.append({'field':'exitCode','actual':exit_code})
    try:
        plan=compile_request(request,request['app']);contracts=contracts_for(request['app']);responses={}
        observed={s['id']:s for s in result_steps(response)}
        for step in plan['steps']:
            record=observed.get(step['id'],{});action=record.get('response')
            if record.get('state')=='succeeded' and action:
                params=resolve(step['params'],responses)
                contracts.validate_result(step['address']['action'],action['data'],params=params)
                responses[step['id']]=action
                verified.append(step['address']['action'])
    except (ValueError,KeyError,TypeError,IndexError) as error:errors.append({'field':'action.result.contract','error':str(error)})
    return {'passed':not errors,'mismatches':errors,'responseVerifiedActions':sorted(set(verified)) if not errors else []}


def artifacts(response, root, case_id):
    """Independent container checks; response correctness is a separate verdict."""
    from wps_skills.client.task_request import result_steps
    records=[]
    for step in result_steps(response):
        # Acquisition describes the input at open time, not the later saved file.
        if step['address']['action'] not in ('save','saveAs','exportPdf','exportSlideImage'):continue
        artifact=((step.get('response') or {}).get('data') or {}).get('artifact')
        if not artifact:continue
        path=Path(artifact['path']).resolve()
        if root.resolve() not in path.parents:raise ValueError('Artifact outside the owned run: '+str(path))
        if not path.is_file() or path.stat().st_size<=0:raise ValueError('Artifact missing or empty: '+str(path))
        if path.stat().st_size!=artifact['sizeBytes']:raise ValueError('Artifact size disagrees with response')
        suffix=path.suffix.lower();level='container'
        if suffix in ('.docx','.xlsx','.pptx'):
            with zipfile.ZipFile(path) as archive:
                if archive.testzip() is not None:raise ValueError('OOXML CRC failure')
                main={'.docx':'word/document.xml','.xlsx':'xl/workbook.xml','.pptx':'ppt/presentation.xml'}[suffix]
                xml=ET.fromstring(archive.read(main))
                if suffix=='.docx' and case_id in ('W01','W02','W03'):
                    text=''.join(n.text or '' for n in xml.findall('.//{*}t'))
                    if case_id in ('W01','W02') and ('执行基线：新值。' not in text or '旧值' in text):raise ValueError('Saved Word replacement disagrees with fixed oracle')
                    if case_id=='W02' and '重新打开后追加。' not in text:raise ValueError('Saved Word append missing')
                    if case_id=='W03':
                        tables=xml.findall('.//{*}tbl')
                        if len(tables)!=1:raise ValueError('Expected one Word table')
                        data=[[''.join(t.text or '' for t in cell.findall('.//{*}t')) for cell in row.findall('{*}tc')] for row in tables[0].findall('{*}tr')]
                        if data!=[['名称','数量'],['测试','2']]:raise ValueError('Saved Word table mismatch')
                        if not any(n.startswith('word/media/') for n in archive.namelist()):raise ValueError('Embedded image missing')
                    level='selected_content'
        elif suffix=='.pdf':
            with path.open('rb') as f:
                if f.read(5)!=b'%PDF-':raise ValueError('Invalid PDF signature')
        elif suffix=='.png':
            with path.open('rb') as f:
                if f.read(8)!=b'\x89PNG\r\n\x1a\n':raise ValueError('Invalid PNG signature')
        records.append({'path':str(path),'sha256':sha(path),'independentCheck':level})
    return records


def execute_prepared(root, skill, resources, timeout=300):
    app=read(root/'manifest.json')['app']
    manifest=read(root/'manifest.json')
    continue_cases=manifest.get('caseFailurePolicy')=='continue_independent'
    # No admission or resumption twice, including an interrupted prior run.
    with (root/'run.claim').open('x') as f:f.write(str(os.getpid()))
    report={'app':app,'state':'running','cases':[],'coverage':manifest['coverage'],
            'remaining':'not_executed','visualReview':'pending','startedUtc':datetime.now(timezone.utc).isoformat()}
    for row in report['coverage'].values():row.update(executedCases=[],responseVerifiedCases=[],invalidParameterPassed=False)
    write(root/'report.json',report)
    env=dict(os.environ,PYTHONUTF8='1',PYTHONIOENCODING='utf-8',WPS_SKILLS_TASK_DIR=str(root/'receipts'),WPS_TRACE_DIR=str(root/'traces'))
    try:
        if os.name!='nt':raise RuntimeError('Real execution requires Windows; use --prepare-only on Mac.')
        from wps_skills.windows.bridge_runtime import windows_powershell_executable
        if str(PureWindowsPath(root).as_posix()).lower()!=manifest['windowsRoot'].lower():raise ValueError('Prepared run was moved; generate a new plan for the actual path.')
        if verify_skill(skill)!=manifest['skillManifestSha256']:raise ValueError('Installed Skill changed since preparation')
        for relative,h in manifest['assets'].items():
            if sha(root/relative)!=h:raise ValueError('Fixture changed: '+relative)
        for row in manifest['cases']:
            if sha(root/row['request'])!=row['requestSha256'] or sha(root/row['expected'])!=row['expectedSha256']:raise ValueError('Prepared request/expectation changed: '+row['id'])
        # Basic environment check does not activate COM.
        probe=subprocess.run([windows_powershell_executable(),'-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(resources/'environment.ps1'),'-ProgId',{'word':'KWPS.Application','excel':'KET.Application','ppt':'KWPP.Application'}[app]],input=json.dumps({'payload':'中文往返'}),text=True,encoding='utf-8',capture_output=True,timeout=15)
        (root/'environment.stderr.txt').write_text(probe.stderr,encoding='utf-8')
        data=json.loads(probe.stdout);write(root/'environment.json',data)
        if probe.returncode or not data['registered'] or not data['userInteractive'] or data['sessionId']==0 or data['echo']!='中文往返':raise RuntimeError('Environment preflight failed; see environment.json')
        entry=skill/'scripts'/(app+'.py')
        for row in manifest['cases']:
            dependency_states={c['id']:c['state'] for c in report['cases']}
            blocked=[d for d in row.get('dependsOn',[]) if dependency_states.get(d)!='passed']
            if blocked:
                report['cases'].append({'id':row['id'],'state':'blocked','setupOnly':row.get('setupOnly',False),'reason':'setup_failed','dependencies':blocked})
                write(root/'report.json',report)
                print('SKIP '+row['id']+' (setup failed)',flush=True)
                continue
            identifier=row['id'];case_dir=root/'results'/identifier;case_dir.mkdir(parents=True)
            request=read(root/row['request']);expected=read(root/row['expected'])
            submitted=root/'submissions'/(identifier+'.json');write(submitted,request)
            case={'id':identifier,'state':'running','setupOnly':expected.get('setupOnly',False)};report['cases'].append(case);write(root/'report.json',report)
            print('RUN '+identifier,flush=True);started=time.perf_counter()
            response=None
            try:
                with (case_dir/'stdout.jsonl').open('w',encoding='utf-8') as out,(case_dir/'stderr.txt').open('w',encoding='utf-8') as err:
                    cp=subprocess.run([sys.executable,str(entry),'--app',app,'--task-file',str(submitted)],env=env,stdout=out,stderr=err,timeout=timeout)
                response=parse_response((case_dir/'stdout.jsonl').read_text(encoding='utf-8-sig'));write(case_dir/'response.json',response)
                case['exitCode']=cp.returncode
                from wps_skills.client.task_request import result_steps
                if response.get('document'):
                    for step in result_steps(response):
                        if step.get('response') is not None:
                            name=step['address']['action']
                            if identifier not in report['coverage'][name]['executedCases']:report['coverage'][name]['executedCases'].append(identifier)
                verdict=evaluate(expected,request,response,cp.returncode);write(case_dir/'assertions.json',verdict)
                if not verdict['passed']:raise AssertionError('Response assertions failed; inspect assertions.json')
                if expected['kind']=='preflight_rejection':
                    if not submitted.exists():raise AssertionError('Rejected request was consumed')
                    report['coverage'][expected['action']]['invalidParameterPassed']=True
                else:
                    if submitted.exists():raise AssertionError('Admitted request was not consumed')
                    query=subprocess.run([sys.executable,str(entry),'--app',app,'--task-status-file',str(submitted)],env=env,capture_output=True,text=True,encoding='utf-8',timeout=30)
                    (case_dir/'status.stdout.jsonl').write_text(query.stdout,encoding='utf-8')
                    (case_dir/'status.stderr.txt').write_text(query.stderr,encoding='utf-8')
                    receipt=parse_response(query.stdout)
                    for key in ('taskId','state','outcome','document','steps','completion','cleanup','stop'):
                        if receipt.get(key)!=response.get(key):raise AssertionError('Receipt mismatch: '+key)
                    for name in verdict['responseVerifiedActions']:report['coverage'][name]['responseVerifiedCases'].append(identifier)
                    if manifest.get('validationMode') != 'response_only':
                        observations=artifacts(response,root,identifier);write(case_dir/'artifacts.json',observations)
                    # For expected failures, planned completion/export files must not appear.
                    if expected['kind']=='negative':
                        for step in expected['actions']:
                            if step['state']!='not_executed':continue
                            source=next(s for s in [request['document'],*request['steps'],*request['completion']] if s['address']==step['address'] and ('id' not in s or s['id']==step['id']))
                            path=source['params'].get('outputPath')
                            if path and Path(path).exists():raise AssertionError('Unexecuted export/save produced a file')
                    elif response.get('completion',{}).get('save'):
                        saved=response['completion']['save']['response']['data']['artifact']['path']
                        cleanup=subprocess.run([windows_powershell_executable(),'-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(resources/'close_verified.ps1'),'-Application',app,'-Path',saved],capture_output=True,text=True,encoding='utf-8',timeout=20)
                        (case_dir/'close.stdout.txt').write_text(cleanup.stdout,encoding='utf-8');(case_dir/'close.stderr.txt').write_text(cleanup.stderr,encoding='utf-8')
                        if cleanup.returncode:raise AssertionError('Verified test document cleanup failed')
                case.update(state='passed',seconds=time.perf_counter()-started)
            except (Exception,KeyboardInterrupt) as error:
                case.update(state='failed',seconds=time.perf_counter()-started,error=str(error),effects='unknown if dispatched; preserved, never replayed')
                terminal_clean=bool(response and response.get('state') in ('completed','stopped') and (response.get('cleanup') or {}).get('outcome')=='succeeded')
                rejected=bool(response and response.get('state')=='rejected' and not response.get('taskId') and not any(s.get('response') for s in response.get('steps',[])))
                if not (continue_cases and not isinstance(error,(KeyboardInterrupt,subprocess.TimeoutExpired)) and (terminal_clean or rejected)):
                    case['continuation']='application_stopped';raise
                case['continuation']='next_independent_case'
                print('FAIL '+identifier+' (recorded; continuing independent cases)',flush=True)
            finally:write(root/'report.json',report)
            if case['state']=='passed':print('PASS '+identifier,flush=True)
        report['state']='completed_with_failures' if any(c['state']!='passed' for c in report['cases']) else 'passed'
        report['remaining']='none'
        if report['state']=='passed' and any(not v['responseVerifiedCases'] or (manifest.get('validationMode') != 'response_only' and not v['invalidParameterPassed']) for v in report['coverage'].values()):raise AssertionError('Observed Action coverage incomplete')
    except (Exception,KeyboardInterrupt) as error:
        report.update(state='failed',error=str(error))
        print('STOP '+str(error),flush=True)
    finally:
        report['endedUtc']=datetime.now(timezone.utc).isoformat();write(root/'report.json',report)
        print('Report: '+str(root/'report.json'),flush=True)
    return 0 if report['state']=='passed' else 1


def main(app, argv=None, package_root=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root',type=Path,help='New Windows evidence directory')
    parser.add_argument('--skill',type=Path,help='Complete standalone Skill to test')
    parser.add_argument('--prepare-only',action='store_true',help='Generate reviewable Task JSON and expected responses; do not execute')
    parser.add_argument('--windows-root',help='Actual target run path when preparing on Mac')
    parser.add_argument('--run-prepared',action='store_true',help='Execute an unchanged prepared directory once')
    parser.add_argument('--case-timeout',type=int,default=300)
    args=parser.parse_args(argv)
    if args.case_timeout<=0 or args.case_timeout>3600:parser.error('case-timeout must be 1..3600 seconds')
    if args.prepare_only and args.run_prepared:parser.error('Choose prepare-only or run-prepared')
    repo=Path(__file__).resolve().parents[5] if package_root is None else Path(package_root)
    skill=(args.skill or repo/('build/skills' if package_root is None else 'skills')/('wps-'+app)).resolve()
    sys.path.insert(0,str(skill/'runtime/src/main/python'))
    resources=repo/('src/test/resources/acceptance' if package_root is None else 'resources')
    root=(args.root or repo/('build/runs/action-acceptance' if package_root is None else 'runs')/(app+'-'+datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S')+'-'+uuid.uuid4().hex[:8])).resolve()
    try:
        if not args.run_prepared:
            manifest=prepare(app,root,skill,args.windows_root)
            print('Prepared '+str(len(manifest['cases']))+' cases; '+str(len(manifest['coverage']))+' supported Actions. '+str(root/'manifest.json'))
        elif read(root/'manifest.json')['app']!=app:raise ValueError('Prepared application differs from this entry')
        if args.prepare_only:return 0
        return execute_prepared(root,skill,resources,args.case_timeout)
    except (OSError,ValueError) as error:parser.exit(1,str(error)+'\n')
