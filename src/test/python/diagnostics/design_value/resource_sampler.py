"""Read-only resource samples for one known Windows run; never inspect document contents."""
import json


def sample(run,metadata,remote,ps_literal):
    root=metadata.get('root') or metadata['remoteRoot']
    script='''$root=ROOT
$rows=@(Get-CimInstance Win32_Process -Filter 'SessionId <> 0' | Where-Object {$_.CommandLine -and $_.CommandLine.Contains($root)} | ForEach-Object {
    $role=if($_.CommandLine -match 'worker.py'){'supervisor'}elseif($_.CommandLine -match 'launch.ps1'){'wrapper'}elseif($_.CommandLine -match 'probe.ps1'){'bridge'}elseif($_.CommandLine -match 'child.py'){'case'}else{'run_process'}
    @{pid=$_.ProcessId;parentPid=$_.ParentProcessId;name=$_.Name;role=$role;workingSetBytes=[long]$_.WorkingSetSize;privateBytes=[long]$_.PrivatePageCount;peakWorkingSetKB=[long]$_.PeakWorkingSetSize;handles=[int]$_.HandleCount;threads=[int]$_.ThreadCount;createdUtc=$_.CreationDate.ToUniversalTime().ToString('o')}
})
@{utc=[DateTime]::UtcNow.ToString('o');processes=$rows;scope='command line contains this run root; desktop sessions only; sampled, not every process peak'}|ConvertTo-Json -Depth 5 -Compress
'''.replace('ROOT',ps_literal(root),1)
    value=json.loads(remote(metadata.get('host','win'),script,run,'resources'))
    with (run/'resource-samples.jsonl').open('a',encoding='utf-8') as output:output.write(json.dumps(value,ensure_ascii=False)+'\n')
    print(json.dumps(value,ensure_ascii=False))
    return value
