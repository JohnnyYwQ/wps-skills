param([Parameter(Mandatory=$true)][string]$ActionsPath, [Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
. $ActionsPath
New-Item -ItemType Directory -Path $OutputDirectory -ErrorAction Stop | Out-Null
$script:target=Join-Path $OutputDirectory 'report.pdf'
function Assert-BoundDocument { param($DocumentId) }
function Sync-ContentRevision { return 'fingerprint' }
function Get-ContentFingerprint { return 'fingerprint' }
function Get-DocumentFingerprint { return 'fingerprint' }
function Get-DocumentPersistenceState { return 'unsaved' }
$script:Revision='revision-1'
$script:Document=[pscustomobject]@{ReadOnly=$false}
$script:Document | Add-Member -MemberType ScriptMethod -Name ExportAsFixedFormat -Value {
    param($temporary, $format)
    [IO.File]::WriteAllText($temporary,'%PDF-test-only')
    # Simulate another writer winning the output name during WPS export.
    [IO.File]::WriteAllText($script:target,'KEEP COMPETING OUTPUT')
}
$rejected=$false
try {
    Invoke-ExportPdf @{documentId='test';operationArguments=@{outputPath=$script:target;overwritePolicy='failIfExists'}} | Out-Null
}
catch {
    if ($_.Exception.Data['WpsCode'] -ne 'OUTPUT_ALREADY_EXISTS') { throw }
    $rejected=$true
}
if (-not $rejected) { throw 'Expected a definite publication conflict' }
if ([IO.File]::ReadAllText($script:target) -ne 'KEEP COMPETING OUTPUT') { throw 'Competing output was overwritten' }
if (@(Get-ChildItem -LiteralPath $OutputDirectory -Force -File).Count -ne 1) { throw 'Temporary PDF was not cleaned up' }
'PDF publication race: passed (no overwrite, definite collision, temporary removed)'
