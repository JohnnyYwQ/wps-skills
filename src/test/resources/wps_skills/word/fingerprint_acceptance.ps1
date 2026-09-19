param(
    [Parameter(Mandatory = $true)][string]$BridgePath,
    [Parameter(Mandatory = $true)][string]$Fixtures,
    [Parameter(Mandatory = $true)][string]$Output
)
$ErrorActionPreference = 'Stop'
if (Test-Path $Output) { throw 'Use a new output directory.' }
[void](New-Item -ItemType Directory $Output)
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($BridgePath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
# Load production functions, without the bridge protocol loop or acquisition.
$needed = @('Get-DocumentFingerprint', 'Get-SectionSnapshots', 'Get-StorySnapshot', 'Normalize-WordText', 'Normalize-StoryText')
foreach ($definition in $ast.EndBlock.Statements) {
    if ($definition -is [Management.Automation.Language.FunctionDefinitionAst] -and $definition.Name -in $needed) { Invoke-Expression $definition.Extent.Text }
}
$base = Split-Path $BridgePath
. (Join-Path $base '../../windows/bridge_common.ps1')
. (Join-Path $base 'word_story_xml.ps1')
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('KWPS.Application')
$documents = $app.Documents
$report = [ordered]@{status='running';cases=@();sessionId=(Get-Process -Id $PID).SessionId}
$fixturePrefix = 'fingerprint-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
function Check($Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
try {
    foreach ($kind in @('plain', 'multi', 'rich', 'populated')) {
        $path = Join-Path $Output ($fixturePrefix+'-'+$kind+'.docx')
        Copy-Item (Join-Path $Fixtures ($kind+'.docx')) $path
        $script:Document = $documents.Open($path)
        $row = [ordered]@{kind=$kind;path=$path;savedAtOpen=[bool]$script:Document.Saved}
        Check $row.savedAtOpen 'Fixture must open saved.'
        $before = Get-DocumentFingerprint
        $again = Get-DocumentFingerprint
        Check ($before -match '^[a-f0-9]{64}$') 'Fingerprint is not SHA-256.'
        Check ($before -eq $again) 'Read-only repeated fingerprints differ.'
        Check ([bool]$script:Document.Saved) 'Fingerprint dirtied the document.'
        $snapshots = @(Get-SectionSnapshots)
        Check ([bool]$script:Document.Saved) 'Section observation dirtied the document.'
        $row.fingerprint = $before
        $row.savedAfterRead = [bool]$script:Document.Saved
        $row.sections = $snapshots
        if ($kind -eq 'plain') {
            foreach ($s in $snapshots[0].headerFooter.stories) { Check ($s.text -ceq '') 'Absent story text is not empty.' }
        }
        elseif ($kind -eq 'multi') {
            Check ($snapshots.Count -eq 3) 'Expected three sections.'
            foreach ($s in $snapshots) {
                foreach ($story in $s.headerFooter.stories) {
                    $type = @{primary='default';firstPage='first';evenPages='even'}[$story.variant]
                    $expected = 's'+[Math]::Min(($s.index+1),2)+'_'+$story.area+'_'+$type
                    Check ($story.text -ceq $expected) ('Wrong section/variant text: '+$expected)
                    Check ($story.linkToPrevious -eq ($s.index -eq 2)) 'Wrong inheritance metadata.'
                }
            }
        }
        elseif ($kind -eq 'rich') {
            # Compare the replacement reader to the original native text semantics.
            $range = $script:Document.Sections.Item(1).Headers.Item(1).Range
            try { $native = Normalize-StoryText -Text ([string]$range.Text) } finally { Release-ComReference $range }
            $observed = @($snapshots[0].headerFooter.stories | Where-Object {$_.area -eq 'header' -and $_.variant -eq 'primary'})[0].text
            Check ($native -ceq $observed) ('XML/native text mismatch: <'+$native+'> vs <'+$observed+'>')
            $row.nativeText = $native
            $row.xmlText = $observed
        }
        elseif ($kind -eq 'populated') {
            $range = $script:Document.StoryRanges.Item(7)
            try { $range.Text = 'FINGERPRINT_UNSAVED_CHANGE' } finally { Release-ComReference $range }
            Check (-not [bool]$script:Document.Saved) 'Intentional edit should remain unsaved.'
            $changed = Get-DocumentFingerprint
            Check ($changed -ne $before) 'Header edit was missed by fingerprint.'
            Check ((Get-DocumentFingerprint) -eq $changed) 'Changed fingerprint is unstable.'
            Check (-not [bool]$script:Document.Saved) 'Read erased unsaved state.'
            $after = @(Get-SectionSnapshots)
            $text = @($after[0].headerFooter.stories | Where-Object {$_.area -eq 'header' -and $_.variant -eq 'primary'})[0].text
            Check ($text -ceq 'FINGERPRINT_UNSAVED_CHANGE') 'Live edit was not observed.'
            $row.changedFingerprint = $changed
            $row.liveText = $text
            $row.savedAfterEditRead = [bool]$script:Document.Saved
        }
        $row.status = 'passed'
        $report.cases += $row
        Release-ComReference $script:Document
        $script:Document = $null
    }
    $report.status = 'passed'
}
catch { $report.status='failed'; $report.error=[string]$_; throw }
finally {
    [IO.File]::WriteAllText((Join-Path $Output 'report.json'),($report|ConvertTo-Json -Depth 20),[Text.UTF8Encoding]::new($false))
    Release-ComReference $script:Document
    Release-ComReference $documents
    Release-ComReference $app
}
