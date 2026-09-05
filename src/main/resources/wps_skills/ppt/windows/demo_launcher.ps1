# Run the visible Ppt demo without allocating an auxiliary console.
# Progress is forwarded to the caller's existing PowerShell window.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$RepositoryRoot,
    [Parameter(Mandatory=$true)][string]$PythonPath,
    [string]$OutputDirectory = '',
    [ValidateRange(0,10)][double]$Delay = 1.5
)
$ErrorActionPreference = 'Stop'
$repository = (Resolve-Path -LiteralPath $RepositoryRoot).Path
$python = (Resolve-Path -LiteralPath $PythonPath).Path
$checkPath = Join-Path $repository 'src/main/resources/wps_skills/ppt/windows/ppt_com.ps1'
. ([scriptblock]::Create([IO.File]::ReadAllText($checkPath)))
$availability = Get-PptComAvailability
if (-not $availability.available) {
    if ($availability.reason -eq 'elevated_user_registration') {
        throw '当前是管理员 PowerShell，无法使用这台电脑上仅为当前用户注册的 WPS 演示。请从开始菜单打开普通 PowerShell（不要选择“以管理员身份运行”），再执行同一命令。尚未创建或打开演示文稿。'
    }
    throw $availability.message
}
$entry = Join-Path $repository 'scripts/demo/ppt.py'
if (-not (Test-Path -LiteralPath $entry -PathType Leaf)) { throw 'Ppt demo entry was not found in this repository.' }
if ($OutputDirectory.Contains('"')) { throw 'OutputDirectory cannot contain a quote.' }
$start = New-Object System.Diagnostics.ProcessStartInfo
$start.FileName = $python
$start.WorkingDirectory = $repository
$start.Arguments = '-X utf8 "' + $entry + '" --delay ' + $Delay.ToString([Globalization.CultureInfo]::InvariantCulture)
if ($OutputDirectory) { $start.Arguments += ' --output-dir "' + $OutputDirectory.TrimEnd('\') + '"' }
$start.UseShellExecute = $false
$start.CreateNoWindow = $true
$start.RedirectStandardOutput = $true
$start.RedirectStandardError = $true
$start.StandardOutputEncoding = [Text.UTF8Encoding]::new($false)
$start.StandardErrorEncoding = [Text.UTF8Encoding]::new($false)
$process = New-Object System.Diagnostics.Process
$process.StartInfo = $start
$started = $false
try {
    $started = $process.Start()
    # Drain stderr concurrently so a full error pipe cannot deadlock stdout.
    $errorRead = $process.StandardError.ReadToEndAsync()
    while ($null -ne ($line = $process.StandardOutput.ReadLine())) { Write-Host $line }
    $process.WaitForExit()
    $errorText = $errorRead.GetAwaiter().GetResult()
    if ($process.ExitCode -ne 0) {
        throw ('Ppt demo failed (exit ' + $process.ExitCode + '): ' + $errorText.Trim())
    }
    if ($errorText) { Write-Warning $errorText.Trim() }
}
finally {
    if ($started -and -not $process.HasExited) {
        # Stop only this launcher-owned client if the caller interrupts the demo.
        # Its Session Host observes channel loss and performs normal cleanup.
        $process.Kill()
        [void]$process.WaitForExit(5000)
    }
    $process.Dispose()
}
