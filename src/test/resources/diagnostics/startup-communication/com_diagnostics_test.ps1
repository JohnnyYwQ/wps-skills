param([Parameter(Mandatory=$true)][string]$Helper, [Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
if (Test-Path -LiteralPath $OutputDirectory) { throw 'Use a fresh diagnostic test directory' }
[void][IO.Directory]::CreateDirectory($OutputDirectory)
$env:WPS_COM_DIAGNOSTIC_DIR=$OutputDirectory
. $Helper
$checks=New-Object Collections.ArrayList

# Missing ProgID is intentionally unique. No real WPS activation or document work.
$missing='WpsSkills.Diagnostics.Missing.' + [guid]::NewGuid().ToString('N')
foreach ($operation in @('attach','activate')) {
    $failed=$false
    try {
        if ($operation -eq 'attach') { $null=Get-DiagnosticActiveApplication -ProgId $missing }
        else { $null=New-DiagnosticApplication -ProgId $missing }
    } catch { $failed=$true }
    if (-not $failed) { throw ('Missing ProgID unexpectedly succeeded: ' + $operation) }
    [void]$checks.Add(('missing_' + $operation))
}

# Exercise error preservation through the same stage wrapper, without changing
# account permissions or registrations on the test machine.
$denied=$false
try {
    $null=Invoke-DiagnosticComStage -Stage 'document.bind_create' -Body {
        throw [Runtime.InteropServices.COMException]::new('Injected access denied', -2147024891)
    }
} catch { $denied=$true }
if (-not $denied) { throw 'The diagnostic wrapper swallowed the exception' }
$value=Invoke-DiagnosticComStage -Stage 'fixture.return_value' -Body { [ordered]@{marker='preserved'; count=2} }
if ($value.marker -cne 'preserved' -or $value.count -ne 2) { throw 'Wrapper changed the returned value' }

function Get-ExcelComAvailability { return @{available=$false; reason='fixture_unavailable'; message='Expected fixture rejection'} }
$availability=Get-DiagnosticExcelComAvailability
if ($availability.available) { throw 'Wrapper changed rejected registration to success' }
$records=@(Get-Content -LiteralPath (Join-Path $OutputDirectory ('com-' + $PID + '.jsonl')) -Encoding UTF8 | ForEach-Object { ConvertFrom-Json -InputObject $_ })
foreach ($stage in @('com.attach','com.activate','document.bind_create','com.registration')) {
    $matches=@($records | Where-Object { $_.stage -eq $stage -and $_.event -eq 'stage.finished' -and $_.outcome -eq 'failed' })
    if ($matches.Count -ne 1) { throw ('Missing exact failure record: ' + $stage) }
}
$binding=@($records | Where-Object { $_.stage -eq 'document.bind_create' -and $_.event -eq 'stage.finished' })[0]
if (@($binding.errors | Where-Object { $_.hresult -eq '0x80070005' }).Count -ne 1) { throw 'Original access-denied HRESULT was not retained' }
[void]$checks.Add('original_hresult_and_throw')
[void]$checks.Add('successful_value_preserved')
[void]$checks.Add('registration_rejection_preserved')
@{passed=$true; checks=@($checks.ToArray()); noWpsDocumentOperations=$true} | ConvertTo-Json -Compress
