param([Parameter(Mandatory = $true)][string]$BridgePath)
$ErrorActionPreference = 'Stop'

# Load the real verification functions without starting the bridge protocol or WPS.
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile(
    $BridgePath, [ref]$tokens, [ref]$errors
)
if ($errors.Count) { throw ($errors | Out-String) }
$definitions = $ast.FindAll({
    param($node)
    ($node -is [Management.Automation.Language.TypeDefinitionAst] -and
        $node.Name -eq 'ContentVerificationException') -or
    ($node -is [Management.Automation.Language.FunctionDefinitionAst] -and
        ($node.Name -match 'LineSpacing' -or
         $node.Name -in @('Get-ParagraphFormatSnapshot', 'Assert-ParagraphFormatPatch')))
}, $false)
foreach ($definition in $definitions) {
    Invoke-Expression $definition.Extent.Text
}
$commonPath = Join-Path (Split-Path $BridgePath) '../../windows/bridge_common.ps1'
$common = [Management.Automation.Language.Parser]::ParseFile(
    $commonPath, [ref]$tokens, [ref]$errors
)
if ($errors.Count) { throw ($errors | Out-String) }
foreach ($definition in $common.FindAll({
    param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -in @('Release-ComReference', 'Get-ObjectPropertyNames')
}, $false)) {
    Invoke-Expression $definition.Extent.Text
}

$cases = @(
    @{name='single'; rule=0; spacing=12; expected=@{kind='single'}; accept=$true},
    @{name='oneAndHalf'; rule=1; spacing=18; expected=@{kind='oneAndHalf'}; accept=$true},
    @{name='double'; rule=2; spacing=24; expected=@{kind='double'}; accept=$true},
    @{name='multiple-one-as-single'; rule=0; spacing=12; expected=@{kind='multiple';value=1}; accept=$true},
    @{name='multiple-half-as-named'; rule=1; spacing=18; expected=@{kind='multiple';value=1.5}; accept=$true},
    @{name='multiple-two-as-double'; rule=2; spacing=24; expected=@{kind='multiple';value=2}; accept=$true},
    @{name='single-as-multiple'; rule=5; spacing=12; expected=@{kind='single'}; accept=$true},
    @{name='named-half-as-multiple'; rule=5; spacing=18; expected=@{kind='oneAndHalf'}; accept=$true},
    @{name='double-as-multiple'; rule=5; spacing=24; expected=@{kind='double'}; accept=$true},
    @{name='arbitrary-multiple'; rule=5; spacing=15; expected=@{kind='multiple';value=1.25}; accept=$true},
    @{name='multiple-within-tolerance'; rule=5; spacing=15.06; expected=@{kind='multiple';value=1.25}; accept=$true},
    @{name='wrong-multiple'; rule=1; spacing=18; expected=@{kind='multiple';value=1.25}; accept=$false},
    @{name='multiple-outside-tolerance'; rule=5; spacing=15.24; expected=@{kind='multiple';value=1.25}; accept=$false},
    @{name='wrong-named-multiple'; rule=2; spacing=24; expected=@{kind='oneAndHalf'}; accept=$false},
    @{name='exact'; rule=4; spacing=18; expected=@{kind='exact';points=18}; accept=$true},
    @{name='exact-within-tolerance'; rule=4; spacing=18.04; expected=@{kind='exact';points=18}; accept=$true},
    @{name='wrong-exact-points'; rule=4; spacing=18.1; expected=@{kind='exact';points=18}; accept=$false},
    @{name='at-least'; rule=3; spacing=18; expected=@{kind='atLeast';points=18}; accept=$true},
    @{name='wrong-at-least-points'; rule=3; spacing=19; expected=@{kind='atLeast';points=18}; accept=$false},
    @{name='exact-is-not-at-least'; rule=4; spacing=18; expected=@{kind='atLeast';points=18}; accept=$false},
    @{name='at-least-is-not-exact'; rule=3; spacing=18; expected=@{kind='exact';points=18}; accept=$false},
    @{name='exact-is-not-a-multiple'; rule=4; spacing=18; expected=@{kind='multiple';value=1.5}; accept=$false},
    @{name='multiple-is-not-exact'; rule=1; spacing=18; expected=@{kind='exact';points=18}; accept=$false},
    @{name='mixed-rule'; rule=9999999; spacing=9999999; expected=@{kind='multiple';value=1.5}; accept=$false}
)
$results = @()
foreach ($case in $cases) {
    $range = [pscustomobject]@{ParagraphFormat = [pscustomobject]@{
        LineSpacingRule=$case.rule; LineSpacing=$case.spacing; Alignment=0;
        SpaceBefore=0; SpaceAfter=0; LeftIndent=0; RightIndent=0; FirstLineIndent=0
    }}
    $accepted = $true
    $message = $null
    try {
        $patch = @{lineSpacing=$case.expected} | ConvertTo-Json | ConvertFrom-Json
        Assert-ParagraphFormatPatch -Range $range -Patch $patch
    }
    catch {
        if ($_.Exception.GetType().Name -ne 'ContentVerificationException') { throw }
        $accepted = $false
        $message = $_.Exception.Message
    }
    $results += @{name=$case.name;passed=($accepted -eq $case.accept);accepted=$accepted;expected=$case.accept;message=$message}
}
$failed = @($results | Where-Object {-not $_.passed})
@{cases=$results.Count;failures=$failed.Count;results=$results} | ConvertTo-Json -Depth 8
if ($failed.Count) { exit 1 }
