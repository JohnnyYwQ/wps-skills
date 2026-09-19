param([Parameter(Mandatory=$true)][string]$BridgePath)
$ErrorActionPreference='Stop'
class ContentVerificationException : System.Exception {
    ContentVerificationException([string]$message) : base($message) {}
}
$helper=Join-Path (Split-Path $BridgePath) 'word_font_names.ps1'
if(Test-Path $helper){. $helper}
# Exercise the production assertion, without starting the bridge input loop.
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($BridgePath,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Bridge parse error'}
$definition=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Assert-TextFormatPatch'},$true)
Invoke-Expression $definition.Extent.Text
function Get-TextFormatSnapshot {param($Range) return $Range.Snapshot}
function Get-ObjectPropertyNames {param($Value) return @($Value.Keys)}
function Release-ComReference {param($Value)}
function Assert-Check($Condition,$Label){if(-not $Condition){throw $Label}}
$checks=0
foreach($field in @('fontFamily','westernFontFamily','eastAsiaFontFamily')){
 $range=[pscustomobject]@{Snapshot=@{$field='微软雅黑'};Font=[pscustomobject]@{NameOther='微软雅黑'}}
 Assert-TextFormatPatch -Range $range -Patch @{$field='Microsoft YaHei'}
 $checks++
}
# Different names must still be rejected with useful expected/observed context.
foreach($field in @('fontFamily','westernFontFamily','eastAsiaFontFamily')){
 $range=[pscustomobject]@{Snapshot=@{$field='Arial'};Font=[pscustomobject]@{NameOther='Arial'}}
 $message=''
 try {Assert-TextFormatPatch -Range $range -Patch @{$field='Microsoft YaHei'}} catch [ContentVerificationException] {$message=$_.Exception.Message}
 Assert-Check ($message.Contains('Microsoft YaHei') -and $message.Contains('Arial') -and $message.Contains($field)) ('Missing mismatch diagnostic: '+$field)
 $checks++
}
$range=[pscustomobject]@{Snapshot=@{westernFontFamily='微软雅黑'};Font=[pscustomobject]@{NameOther='Arial'}}
$message=''
try {Assert-TextFormatPatch -Range $range -Patch @{westernFontFamily='Microsoft YaHei'}} catch [ContentVerificationException] {$message=$_.Exception.Message}
Assert-Check ($message.Contains('NameOther') -and $message.Contains('Arial')) 'Western fallback must be checked separately'
$checks++
Assert-Check (-not (Test-WordFontNameEquivalent 'DefinitelyNotInstalled_123' 'Arial')) 'Unknown name must not use font fallback'
$checks++
# Deterministic catalog ambiguity and lookup failure checks; no invented aliases.
function Get-InstalledWordFontFamilies {
 [pscustomobject]@{Names=@('Family A','Alias A','Shared')}
 [pscustomobject]@{Names=@('Family B','Alias B','Shared')}
}
$script:WordFontNameIndex=$null
Assert-Check (Test-WordFontNameEquivalent 'Family A' 'Alias A') 'Same installed family'
Assert-Check (-not (Test-WordFontNameEquivalent 'Family A' 'Alias B')) 'Different families'
Assert-Check (-not (Test-WordFontNameEquivalent 'Family A' 'Shared')) 'Ambiguous name'
$checks+=3
function Get-InstalledWordFontFamilies {throw 'Catalog unavailable'}
$script:WordFontNameIndex=$null
Assert-Check (Test-WordFontNameEquivalent 'Arial' 'ARIAL') 'Exact name must not require catalog'
Assert-Check (-not (Test-WordFontNameEquivalent 'Family A' 'Alias A')) 'Unavailable catalog fails closed'
$checks+=2
@{passed=$checks}|ConvertTo-Json -Compress
