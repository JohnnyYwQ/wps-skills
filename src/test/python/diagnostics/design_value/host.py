"""Mac pack/start/status/collect entry; Windows runs independently."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
import uuid

from .build import build, REPO
from . import GROUPS


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    sub=parser.add_subparsers(dest='command',required=True)
    pack=sub.add_parser('pack');pack.add_argument('--output',type=Path,required=True)
    report=sub.add_parser('report');report.add_argument('--runs',nargs='+',type=Path,required=True);report.add_argument('--output',type=Path,required=True)
    start=sub.add_parser('start');start.add_argument('--kit',type=Path,required=True);start.add_argument('--name',required=True)
    start.add_argument('--groups',nargs='+',choices=GROUPS,default=['preflight','replay','focus','readback']);start.add_argument('--trials',type=int,default=3)
    start.add_argument('--apps',nargs='+',default=['word','excel','ppt'],choices=['word','excel','ppt'])
    for name in ('status','collect','resources'):
        child=sub.add_parser(name);child.add_argument('--run',type=Path,required=True)
    args=parser.parse_args()
    if args.command=='pack':print(build(args.output));return
    if args.command=='report':
        from .report import summarize
        summarize(args.runs,args.output);return
    sys.path.insert(0,str(REPO/'src/test/python/diagnostics/startup_communication'))
    from debug import remote, command, ps_literal, extract_snapshot
    from common import write_json, read_json
    if args.command=='start':
        if not re.fullmatch(r'[A-Za-z0-9_-]+',args.name):parser.error('Invalid run name')
        if any(not re.fullmatch(r'[a-z_]+',g) for g in args.groups):parser.error('Invalid group')
        run=REPO/'build/runs/design-value'/args.name;run.mkdir(parents=True,exist_ok=False)
        archive=args.kit.resolve().with_suffix('.zip');digest=hashlib.sha256(archive.read_bytes()).hexdigest()
        metadata={'state':'submission_started','name':args.name,'archiveSha256':digest,'kit':str(args.kit.resolve())}
        write_json(run/'run.json',metadata)
        prepared=json.loads(remote('win',f"""
$root=Join-Path $env:USERPROFILE 'wps-design-value/{args.name}'
if(Test-Path $root){{throw 'Fresh remote run required'}}
[void](New-Item -ItemType Directory -Path $root)
$python=Join-Path $env:USERPROFILE '.workbuddy/binaries/python/versions/3.13.12/python.exe'
if(!(Test-Path $python)){{throw 'Known Windows interpreter unavailable'}}
@{{root=$root;python=$python}}|ConvertTo-Json -Compress
""",run,'prepare'))
        metadata.update(prepared);write_json(run/'run.json',metadata)
        command(['scp',str(archive),'win:'+prepared['root'].replace('\\','/')+'/kit.zip'],run,'upload')
        launch='''import sys,runpy
from pathlib import Path
root=Path(__file__).resolve().parent
sys.stdout=(root/'stdout.txt').open('w',encoding='utf-8',buffering=1)
sys.stderr=(root/'stderr.txt').open('w',encoding='utf-8',buffering=1)
sys.path.insert(0,str(root/'kit'))
sys.argv=[str(root/'kit/run.py'),'--root',str(root/'results'),'--groups',GROUPS,'--apps',APPS,'--trials','TRIALS']
runpy.run_path(root/'kit/run.py',run_name='__main__')
'''.replace('GROUPS',','.join(repr(x) for x in args.groups)).replace('APPS',','.join(repr(x) for x in args.apps)).replace('TRIALS',str(args.trials))
        task='WpsDesign-'+args.name
        metadata.update(state='trigger_pending',taskName=task);write_json(run/'run.json',metadata)
        result=remote('win',f"""
$root={ps_literal(prepared['root'])}
$zip=Join-Path $root 'kit.zip'
if((Get-FileHash $zip -Algorithm SHA256).Hash.ToLowerInvariant() -ne '{digest}'){{throw 'Upload hash mismatch'}}
Expand-Archive $zip -DestinationPath (Join-Path $root 'kit')
[IO.File]::WriteAllText((Join-Path $root 'launch.py'),{ps_literal(launch)},(New-Object Text.UTF8Encoding($false)))
$action=New-ScheduledTaskAction -Execute {ps_literal(prepared['python'])} -Argument ('"'+(Join-Path $root 'launch.py')+'"')
$principal=New-ScheduledTaskPrincipal -UserId ([Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
$settings=New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 45) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
if(Get-ScheduledTask -TaskName '{task}' -ErrorAction SilentlyContinue){{throw 'Task exists'}}
Register-ScheduledTask -TaskName '{task}' -Action $action -Principal $principal -Settings $settings|Out-Null
Start-ScheduledTask -TaskName '{task}'
'Submitted'
""",run,'trigger')
        metadata['state']='submitted';write_json(run/'run.json',metadata);print(result,run);return
    run=args.run.resolve();metadata=read_json(run/'run.json')
    if args.command=='collect' and (run/'collected').exists():
        parser.error('Existing collection preserved. Use report to reverify it; do not overwrite evidence.')
    if args.command=='resources':
        from .resource_sampler import sample
        sample(run,metadata,remote,ps_literal);return
    prefix=(REPO/'src/test/resources/diagnostics/startup-communication/snapshot_io.ps1').read_text(encoding='utf-8-sig')+'\n$root='+ps_literal(metadata['root'])+'\n'
    if args.command=='status':
        print(remote('win',prefix+f"""
$file=Join-Path $root 'results/report.json'
if(Test-Path $file){{$r=ConvertFrom-Json -InputObject (Read-DiagnosticText -Path $file);$last=@($r.experiments)[-1];@{{state=$r.state;count=@($r.experiments).Count;last=@{{id=$last.id;state=$last.state;error=$last.error}};error=$r.error}}|ConvertTo-Json -Depth 15 -Compress}}
$scheduled=Get-ScheduledTask -TaskName '{metadata['taskName']}'
$info=Get-ScheduledTaskInfo -InputObject $scheduled
@{{scheduledState=[string]$scheduled.State;lastTaskResult=$info.LastTaskResult}}|ConvertTo-Json -Compress
if($r -and $r.state -eq 'running' -and $scheduled.State -ne 'Running'){{'Supervisor is not running; stale report is interrupted/unknown. Do not replay.'}}
if(Test-Path (Join-Path $root 'stderr.txt')){{Read-DiagnosticText -Path (Join-Path $root 'stderr.txt')}}
""",run,'status'));return
    snapshot='evidence-'+uuid.uuid4().hex
    observed=json.loads(remote('win',prefix+f"""
if((Get-ScheduledTask -TaskName '{metadata['taskName']}').State -eq 'Running'){{throw 'Wait for final evidence; no live snapshot in this entry'}}
$snapshot=Join-Path $root '{snapshot}'
[void](New-Item -ItemType Directory -Path $snapshot)
$files=@{{}}
$sourceFiles=@(Get-ChildItem (Join-Path $root 'results') -File -Recurse)+@(Get-Item (Join-Path $root 'stdout.txt'),(Join-Path $root 'stderr.txt'))
foreach($file in $sourceFiles){{
    $relative=$file.FullName.Substring($root.Length+1)
    $target=Join-Path $snapshot $relative
    [void](New-Item -ItemType Directory -Force -Path (Split-Path $target -Parent))
    Copy-DiagnosticEvidenceFile -Source $file.FullName -Destination $target
    $files[$relative.Replace([char]92,[char]47)]=(Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash.ToLowerInvariant()
}}
[IO.File]::WriteAllText((Join-Path $snapshot 'files.json'),(ConvertTo-Json -InputObject $files -Depth 3),(New-Object Text.UTF8Encoding($false)))
$zip=$snapshot+'.zip'
Add-Type -AssemblyName System.IO.Compression.FileSystem
[IO.Compression.ZipFile]::CreateFromDirectory($snapshot,$zip)
@{{sha256=(Get-FileHash $zip -Algorithm SHA256).Hash.ToLowerInvariant();count=$files.Count}}|ConvertTo-Json -Compress
""",run,'archive'))
    archive=run/'evidence.zip'
    command(['scp','win:'+metadata['root'].replace('\\','/')+'/'+snapshot+'.zip',str(archive)],run,'download')
    assert hashlib.sha256(archive.read_bytes()).hexdigest()==observed['sha256']
    target=run/'collected';target.mkdir(exist_ok=False);extract_snapshot(archive,target)
    for name,digest in read_json(target/'files.json').items():assert hashlib.sha256((target/name).read_bytes()).hexdigest()==digest,name
    write_json(run/'collection.json',dict(observed,verified=True));print('Verified',observed['count'],'files',target)
