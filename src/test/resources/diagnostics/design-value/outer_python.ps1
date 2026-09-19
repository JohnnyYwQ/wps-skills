param([string]$Python,[string]$Script,[string]$Config)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
& $Python $Script --config $Config
exit $LASTEXITCODE
