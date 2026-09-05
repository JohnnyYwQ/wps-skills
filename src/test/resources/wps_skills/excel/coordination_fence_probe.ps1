param([string]$CommonPath, [string]$File, [ValidateSet('hold','probe')][string]$Mode)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
class DocumentLeaseConflictException : System.Exception {
    DocumentLeaseConflictException([string]$message) : base($message) {}
}
class DocumentQuarantinedException : System.Exception {
    DocumentQuarantinedException([string]$message) : base($message) {}
}
. $CommonPath
$script:PreparationId = 'fixture-preparation'
$script:PreparedIdentity = Get-StableFileIdentity -Path $File
$locator = Get-NormalizedFileLocator -Path $File
try {
    Add-CoordinationFence -Identity ('locator-' + $locator)
    Invoke-AcquireCoordinationGuard -Arguments ([pscustomobject]@{ coordinationIdentity = $script:PreparedIdentity }) | Out-Null
    if ($Mode -eq 'probe') { throw 'A replacement file escaped the quarantined locator fence.' }
    Set-CoordinationInFlight -InFlight $true
    # Simulate replacement while an operation is marked in flight. This fixture
    # never starts WPS or COM and only changes its caller-created scratch file.
    [IO.File]::WriteAllText($File + '.replacement', 'replacement')
    Move-Item -LiteralPath ($File + '.replacement') -Destination $File -Force
    [Console]::Out.WriteLine('READY')
    [Console]::Out.Flush()
    [void][Console]::In.ReadLine()
    throw 'The fixture must be terminated by its owning test process.'
}
catch [DocumentQuarantinedException] {
    if ($Mode -ne 'probe') { throw }
    [Console]::Out.WriteLine('QUARANTINED')
}
