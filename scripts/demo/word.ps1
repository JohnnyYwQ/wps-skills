[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$PythonPath,
    [string]$OutputDirectory = '',
    [ValidateRange(0,10)][double]$Delay = 1.5
)
$repository = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$resource = Join-Path $repository 'src/main/resources/wps_skills/word/windows/demo_launcher.ps1'
$launch = [scriptblock]::Create([IO.File]::ReadAllText($resource))
& $launch -RepositoryRoot $repository -PythonPath $PythonPath -OutputDirectory $OutputDirectory -Delay $Delay
