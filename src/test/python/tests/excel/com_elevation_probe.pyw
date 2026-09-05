"""Desktop regression: elevated per-user COM registration must fail before demo creation."""
import argparse
import base64
import json
import os
from pathlib import Path
import subprocess

parser = argparse.ArgumentParser()
parser.add_argument('--repo', required=True, type=Path)
parser.add_argument('--output', required=True, type=Path)
args = parser.parse_args()
args.output.mkdir(parents=True, exist_ok=False)
config = {'repo': str(args.repo), 'output': str(args.output/'must-not-exist')}
encoded = base64.b64encode(json.dumps(config).encode('utf-8')).decode('ascii')
code = "$config=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('"+encoded+"'))|ConvertFrom-Json\n" + r'''
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$dir=Join-Path $config.repo 'src/main/resources/wps_skills/excel/windows'
. ([scriptblock]::Create([IO.File]::ReadAllText((Join-Path $dir 'excel_com.ps1'))))
$r=Get-ExcelComAvailability
if($r.available -or $r.reason -ne 'elevated_user_registration'){throw 'Expected elevated user-only registration to be unavailable'}
try {
 & ([scriptblock]::Create([IO.File]::ReadAllText((Join-Path $dir 'demo_launcher.ps1')))) -RepositoryRoot $config.repo -PythonPath (Join-Path $env:USERPROFILE '.conda/envs/yolov8app/python.exe') -OutputDirectory $config.output -Delay 0
 throw 'Expected the launcher to reject this elevated context'
} catch {
 if($_.Exception.Message -notlike '*PowerShell*'){throw}
 $r.message=$_.Exception.Message
}
if(Test-Path -LiteralPath $config.output){throw 'The blocked launcher created a demo directory'}
$r.status='passed'
$r|ConvertTo-Json -Compress
'''
result = subprocess.run([str(Path(os.environ['WINDIR'])/'System32/WindowsPowerShell/v1.0/powershell.exe'),
    '-NoProfile', '-NonInteractive', '-EncodedCommand', base64.b64encode(code.encode('utf-16le')).decode('ascii')],
    capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, timeout=20)
report = {'status': 'passed' if result.returncode == 0 else 'failed',
          'stdout': result.stdout.decode('utf-8', errors='replace'),
          'stderr': result.stderr.decode('utf-8', errors='replace')}
(args.output/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
