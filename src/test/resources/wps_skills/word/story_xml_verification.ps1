param([Parameter(Mandatory = $true)][string]$BridgePath)
$ErrorActionPreference = 'Stop'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($BridgePath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
foreach ($definition in $ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -in @('Normalize-WordText', 'Normalize-StoryText')}, $false)) { Invoke-Expression $definition.Extent.Text }
. (Join-Path (Split-Path $BridgePath) 'word_story_xml.ps1')
$w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
$cases = 0
function Package([string]$Body, [string]$Relations = '', [string]$Parts = '') {
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage" xmlns:w="'+$w+'" xmlns:r="'+$r+'"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>'+$Body+'</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/_rels/document.xml.rels"><pkg:xmlData><Relationships xmlns="'+$rel+'">'+$Relations+'</Relationships></pkg:xmlData></pkg:part>'+$Parts+'</pkg:package>'
}
function Check($Value, $Expected, [string]$Name) {
    $script:cases++
    if ($Value -cne $Expected) { throw "$Name expected <$Expected>, got <$Value>" }
}
function Reject([string]$Xml, [int]$Count, [string]$Name) {
    $script:cases++
    $rejected = $false
    try { $null = Get-WordXmlStoryTextMap -WordOpenXml $Xml -ExpectedSectionCount $Count } catch { $rejected = $true }
    if (-not $rejected) { throw "$Name was accepted" }
}
$empty = Package '<w:p/><w:sectPr/>'
$map = Get-WordXmlStoryTextMap $empty 1
Check $map.Count 6 'six variants'
foreach ($value in $map.Values) { Check $value '' 'absent story stays empty' }
$relations = '<Relationship Id="h" Type="'+$r+'/header" Target="header-custom.xml"/>'
$part = '<pkg:part pkg:name="/word/header-custom.xml"><pkg:xmlData><w:hdr><w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t></w:r></w:p><w:p><w:r><w:t>D</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>E</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>F</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:fldSimple w:instr="QUOTE VALUE"><w:r><w:t>VALUE</w:t></w:r></w:fldSimple><w:del><w:r><w:delText>OLD</w:delText></w:r></w:del><w:ins><w:r><w:t>NEW</w:t></w:r></w:ins><w:r><w:rPr><w:vanish/></w:rPr><w:t>HIDDEN</w:t></w:r></w:p></w:hdr></pkg:xmlData></pkg:part>'
$reference = '<w:headerReference w:type="default" r:id="h"/>'
$xml = Package ('<w:p><w:pPr><w:sectPr>'+$reference+'</w:sectPr></w:pPr></w:p><w:sectPr/>') $relations $part
$map = Get-WordXmlStoryTextMap $xml 2
Check $map['0/header/primary'] "A`tB`nC`nD`nE`tF`nVALUENEW" 'visible text and table normalization'
Check $map['1/header/primary'] $map['0/header/primary'] 'inherit previous section'
Check $map['1/header/firstPage'] '' 'variant remains independent'
$clear = '<pkg:part pkg:name="/word/clear.xml"><pkg:xmlData><w:hdr><w:p/></w:hdr></pkg:xmlData></pkg:part>'
$rels2 = $relations + '<Relationship Id="clear" Type="'+$r+'/header" Target="clear.xml"/>'
$body = '<w:p><w:pPr><w:sectPr>'+$reference+'</w:sectPr></w:pPr></w:p><w:sectPr><w:headerReference w:type="default" r:id="clear"/></w:sectPr>'
$map = Get-WordXmlStoryTextMap (Package $body $rels2 ($part+$clear)) 2
Check $map['1/header/primary'] '' 'explicit blank must not inherit'
$stylePart = '<pkg:part pkg:name="/word/styles.xml"><pkg:xmlData><w:styles><w:style w:styleId="Hide" w:type="character"><w:rPr><w:vanish/></w:rPr></w:style></w:styles></pkg:xmlData></pkg:part>'
$styled = '<pkg:part pkg:name="/word/header-custom.xml"><pkg:xmlData><w:hdr><w:p><w:r><w:rPr><w:rStyle w:val="Hide"/></w:rPr><w:t>HIDDEN</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Hide"/><w:vanish w:val="0"/></w:rPr><w:t>VISIBLE</w:t></w:r></w:p></w:hdr></pkg:xmlData></pkg:part>'
$map = Get-WordXmlStoryTextMap (Package ('<w:sectPr>'+$reference+'</w:sectPr>') $relations ($styled+$stylePart)) 1
Check $map['0/header/primary'] 'VISIBLE' 'hidden style and explicit override'
Reject $xml 1 'section count mismatch'
Reject (Package ('<w:sectPr>'+$reference+'</w:sectPr>') $relations '') 1 'missing referenced part'
Reject (Package ('<w:sectPr>'+$reference+'</w:sectPr>') ($relations.Replace('Target="header-custom.xml"','Target="https://example.invalid/header.xml" TargetMode="External"')) $part) 1 'external relationship'
Reject ('<!DOCTYPE x [<!ENTITY y "z">]>'+$empty) 1 'DTD'
@{cases=$cases;failures=0}|ConvertTo-Json -Compress
