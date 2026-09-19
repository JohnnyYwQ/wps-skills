param([Parameter(Mandatory=$true)][string]$Root,
      [Parameter(Mandatory=$true)][string]$Python,
      [Parameter(Mandatory=$true)][ValidateSet('quick','long')][string]$Profile)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
$evidence = Join-Path $Root 'windows'
[void][IO.Directory]::CreateDirectory($evidence)
$record = [ordered]@{wrapperStartedUtc=[DateTime]::UtcNow.ToString('o'); pid=$PID; state='running'}
$encoding = New-Object Text.UTF8Encoding($false)
$recordPath = Join-Path $evidence 'launcher.json'
[IO.File]::WriteAllText($recordPath, ($record | ConvertTo-Json), $encoding)
try {
    # Native stderr is evidence, not a PowerShell terminating error. Preserve the full traceback.
    $arguments = @(('"'+(Join-Path $Root 'bundle/worker.py')+'"'), '--root', ('"'+$Root+'"'), '--profile', $Profile)
    $child = Start-Process -FilePath $Python -ArgumentList $arguments -NoNewWindow -PassThru -Wait `
        -RedirectStandardOutput (Join-Path $evidence 'worker.stdout') -RedirectStandardError (Join-Path $evidence 'worker.stderr')
    $record.exitCode = $child.ExitCode
    $record.state = if ($child.ExitCode -eq 0) {'completed'} else {'failed'}
} catch {
    $record.state = 'failed'
    $record.error = ($_ | Out-String)
} finally {
    $record.finishedUtc = [DateTime]::UtcNow.ToString('o')
    [IO.File]::WriteAllText($recordPath, ($record | ConvertTo-Json), $encoding)
}
if ($record.state -eq 'failed') { exit 1 }
