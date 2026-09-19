"""Prepare or run the fixed 20/20/20 business suite through production Task CLI."""
import argparse
from datetime import datetime, timezone
from pathlib import Path, PureWindowsPath
import sys
import uuid
from tests.applications.action_acceptance import write, read, sha, verify_skill, execute_prepared


def prepare(app,root,skill,windows_root):
    from tests.applications.complex_business_suite import build_cases
    from tests.applications.action_suite import coverage
    from tests.applications.native_acceptance import make_png
    if root.exists():raise FileExistsError(root)
    target=PureWindowsPath(windows_root).as_posix().rstrip('/')
    if not PureWindowsPath(target).is_absolute():raise ValueError('Absolute Windows target required')
    cases=build_cases(app,target)
    business=[c for c in cases if not c.setup]
    matrix=coverage(app,business)
    for v in matrix.values():v['invalidParameterCase']=None
    root.mkdir(parents=True);(root/'outputs').mkdir();(root/'assets').mkdir();make_png(root/'assets/sample.png')
    manifest={'schemaVersion':1,'app':app,'windowsRoot':target,'skillManifestSha256':verify_skill(skill),
              'validationMode':'response_only','scope':'20 fixed positive business Tasks; setup counted separately',
              'createdUtc':datetime.now(timezone.utc).isoformat(),'state':'prepared','coverage':matrix,'cases':[],
              'businessTaskCount':20,'setupTaskCount':len(cases)-20,'visualReview':'not_in_scope',
              'caseFailurePolicy':'continue_independent'}
    setup_outputs={c.request['completion'][0]['params']['outputPath']:c.id for c in cases if c.setup}
    for c in cases:
        expected=c.expectation()
        if any(not s['dataAssertions'] for s in expected['actions']):raise ValueError('Missing data assertions: '+c.id)
        rp='requests/'+c.id+'.json';ep='expected/'+c.id+'.json';write(root/rp,c.request);write(root/ep,expected)
        manifest['cases'].append({'id':c.id,'title':c.title,'setupOnly':c.setup,'request':rp,'expected':ep,
            'requestSha256':sha(root/rp),'expectedSha256':sha(root/ep),'actionCount':len(expected['actions']),
            'distinctActionCount':len({s['address']['action'] for s in expected['actions']}),
            'dependsOn':[setup_outputs[c.existing]] if c.existing in setup_outputs else []})
    manifest['assets']={'assets/sample.png':sha(root/'assets/sample.png')};write(root/'manifest.json',manifest)
    return manifest


def main(package_root,argv=None,fixed_app=None):
    p=argparse.ArgumentParser(description=__doc__)
    if fixed_app:p.set_defaults(app=fixed_app)
    else:p.add_argument('--app',choices=['word','excel','ppt','all'],default='all')
    p.add_argument('--root',type=Path)
    p.add_argument('--prepare-only',action='store_true')
    p.add_argument('--windows-root')
    p.add_argument('--case-timeout',type=int,default=300)
    a=p.parse_args(argv);package=Path(package_root)
    if a.root is None:
        a.root=package/'runs'/(a.app+'-'+datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S')+'-'+uuid.uuid4().hex[:8])
    if not 1<=a.case_timeout<=3600:p.error('case-timeout must be 1..3600')
    # Each application runs in a fresh Python process via the top-level all wrapper.
    if a.app=='all':
        import subprocess
        failed=False
        for app in ('word','excel','ppt'):
            cmd=[sys.executable,str(package/('run_'+app+'.py')),'--root',str(a.root/app),'--case-timeout',str(a.case_timeout)]
            if a.prepare_only:cmd+=['--prepare-only','--windows-root',a.windows_root.rstrip('/')+'/'+app]
            result=subprocess.run(cmd)
            failed=failed or result.returncode!=0
        return 1 if failed else 0
    skill=package/'skills'/('wps-'+a.app)
    sys.path.insert(0,str(skill/'runtime/src/main/python'))
    if a.prepare_only:
        if not a.windows_root:p.error('--windows-root required for prepare-only')
        m=prepare(a.app,a.root,skill,a.windows_root)
        print('PREPARED',a.app,m['businessTaskCount'],'business,',m['setupTaskCount'],'setup',flush=True)
        return 0
    if not (a.root/'manifest.json').exists():
        import os
        if os.name!='nt':p.error('Execution requires Windows; use --prepare-only on Mac.')
        prepare(a.app,a.root,skill,str(a.root.resolve()))
    m=read(a.root/'manifest.json')
    if m['app']!=a.app or m.get('businessTaskCount')!=20:raise ValueError('Unexpected prepared manifest')
    # ZIP transports may omit empty directories. Materialize only the owned
    # output directory before WPS admission, not arbitrary request paths.
    if (a.root/'run.claim').exists():raise FileExistsError('Existing run: never replay')
    (a.root/'outputs').mkdir(exist_ok=True)
    return execute_prepared(a.root,skill,package/'resources',a.case_timeout)
