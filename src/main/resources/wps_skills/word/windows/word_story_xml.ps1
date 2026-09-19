# Read live header/footer text without materializing HeaderFooter.Range in WPS.
# The caller owns one WordOpenXML snapshot; no disk fallback or Saved assignment.
function Get-WordXmlToggle {
    param($Node, [System.Xml.XmlNamespaceManager]$Namespaces)
    if ($null -eq $Node) { return $null }
    $value = $Node.GetAttribute('val', $Namespaces.LookupNamespace('w'))
    return ($value -notin @('0', 'false', 'off'))
}

function Get-WordXmlStyleHidden {
    param([string]$StyleId, [hashtable]$Context)
    $seen = @{}
    while ($StyleId -and $Context.styles.ContainsKey($StyleId)) {
        if ($seen.ContainsKey($StyleId)) { throw 'Cyclic Word style inheritance.' }
        $seen[$StyleId] = $true
        $style = $Context.styles[$StyleId]
        $hidden = Get-WordXmlToggle -Node ($style.SelectSingleNode('w:rPr/w:vanish', $Context.ns)) -Namespaces $Context.ns
        if ($null -ne $hidden) { return $hidden }
        $base = $style.SelectSingleNode('w:basedOn', $Context.ns)
        $StyleId = if ($null -eq $base) { '' } else { $base.GetAttribute('val', $Context.w) }
    }
    return $null
}

function Test-WordXmlHiddenRun {
    param([System.Xml.XmlNode]$Node, [hashtable]$Context)
    $hidden = Get-WordXmlToggle -Node ($Node.SelectSingleNode('w:rPr/w:vanish', $Context.ns)) -Namespaces $Context.ns
    if ($null -ne $hidden) { return $hidden }
    $runStyle = $Node.SelectSingleNode('w:rPr/w:rStyle', $Context.ns)
    if ($null -ne $runStyle) {
        $hidden = Get-WordXmlStyleHidden -StyleId ($runStyle.GetAttribute('val', $Context.w)) -Context $Context
        if ($null -ne $hidden) { return $hidden }
    }
    $paragraph = $Node.ParentNode
    while ($null -ne $paragraph -and -not ($paragraph.LocalName -eq 'p' -and $paragraph.NamespaceURI -eq $Context.w)) {
        $paragraph = $paragraph.ParentNode
    }
    $styleId = $Context.defaultParagraphStyle
    if ($null -ne $paragraph) {
        $style = $paragraph.SelectSingleNode('w:pPr/w:pStyle', $Context.ns)
        if ($null -ne $style) { $styleId = $style.GetAttribute('val', $Context.w) }
    }
    $hidden = Get-WordXmlStyleHidden -StyleId $styleId -Context $Context
    if ($null -ne $hidden) { return $hidden }
    return $Context.defaultHidden
}

function Add-WordXmlStoryText {
    param([System.Xml.XmlNode]$Node, [Text.StringBuilder]$Text, [hashtable]$Context)
    if ($Node -isnot [System.Xml.XmlElement]) { return }
    if ($Node.NamespaceURI -eq $Context.w) {
        switch ($Node.LocalName) {
            { $_ -in @('del', 'moveFrom', 'instrText', 'delText', 'rPr', 'pPr', 'sectPr', 'tblPr', 'trPr', 'tcPr', 'drawing', 'pict', 'object') } { return }
            'r' { if (Test-WordXmlHiddenRun -Node $Node -Context $Context) { return } }
            't' { [void]$Text.Append($Node.InnerText); return }
            'tab' { [void]$Text.Append("`t"); return }
            'ptab' { [void]$Text.Append("`t"); return }
            'br' {
                $kind = $Node.GetAttribute('type', $Context.w)
                if ($kind -eq 'page') { [void]$Text.Append([char]12) }
                elseif ($kind -ne 'column') { [void]$Text.Append([char]11) }
                return
            }
            'cr' { [void]$Text.Append([char]11); return }
            'noBreakHyphen' { [void]$Text.Append([char]0x2011); return }
            'softHyphen' { [void]$Text.Append([char]0x00ad); return }
            'sym' { [void]$Text.Append([char][Convert]::ToInt32($Node.GetAttribute('char', $Context.w), 16)); return }
        }
    }
    # AlternateContent carries two representations of the same content.
    if ($Node.NamespaceURI -eq 'http://schemas.openxmlformats.org/markup-compatibility/2006' -and $Node.LocalName -eq 'AlternateContent') {
        $choice = $Node.SelectSingleNode('*[local-name()="Choice"]')
        if ($null -eq $choice) { $choice = $Node.SelectSingleNode('*[local-name()="Fallback"]') }
        if ($null -ne $choice) { Add-WordXmlStoryText -Node $choice -Text $Text -Context $Context }
        return
    }
    foreach ($child in $Node.ChildNodes) { Add-WordXmlStoryText -Node $child -Text $Text -Context $Context }
    if ($Node.NamespaceURI -eq $Context.w) {
        switch ($Node.LocalName) {
            'p' { [void]$Text.Append("`r") }
            'tc' { [void]$Text.Append([char]7) }
            'tr' { [void]$Text.Append("`r`a") }
        }
    }
}

function Get-WordXmlStoryTextMap {
    param(
        [Parameter(Mandatory = $true)][string]$WordOpenXml,
        [Parameter(Mandatory = $true)][int]$ExpectedSectionCount
    )
    $settings = [Xml.XmlReaderSettings]::new()
    $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $stringReader = [IO.StringReader]::new($WordOpenXml)
    $reader = [Xml.XmlReader]::Create($stringReader, $settings)
    $xml = [Xml.XmlDocument]::new()
    $xml.PreserveWhitespace = $true
    $xml.XmlResolver = $null
    try { $xml.Load($reader) } finally { $reader.Dispose(); $stringReader.Dispose() }
    $ns = [Xml.XmlNamespaceManager]::new($xml.NameTable)
    $ns.AddNamespace('pkg', 'http://schemas.microsoft.com/office/2006/xmlPackage')
    $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $ns.AddNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships')
    $parts = @{}
    foreach ($part in $xml.SelectNodes('/pkg:package/pkg:part', $ns)) {
        $name = $part.GetAttribute('name', $ns.LookupNamespace('pkg'))
        if ($parts.ContainsKey($name)) { throw 'Duplicate Word XML package part.' }
        $parts[$name] = $part.SelectSingleNode('pkg:xmlData/*', $ns)
    }
    $document = $parts['/word/document.xml']
    if ($null -eq $document) { throw 'Live Word XML has no document part.' }
    $sections = @($document.SelectNodes('w:body//w:sectPr[not(ancestor::w:sectPrChange)]', $ns))
    if ($sections.Count -ne $ExpectedSectionCount) { throw 'Live Word XML section count does not match the document.' }
    $relationships = @{}
    $rels = $parts['/word/_rels/document.xml.rels']
    if ($null -ne $rels) {
        foreach ($rel in $rels.SelectNodes('rel:Relationship', $ns)) {
            $id = $rel.GetAttribute('Id')
            if ($relationships.ContainsKey($id)) { throw 'Duplicate Word XML relationship.' }
            $relationships[$id] = $rel
        }
    }
    $context = @{ ns = $ns; w = $ns.LookupNamespace('w'); styles = @{}; defaultHidden = $false; defaultParagraphStyle = '' }
    $styles = $parts['/word/styles.xml']
    if ($null -ne $styles) {
        foreach ($style in $styles.SelectNodes('w:style', $ns)) {
            $id = $style.GetAttribute('styleId', $context.w)
            $context.styles[$id] = $style
            if ($style.GetAttribute('type', $context.w) -eq 'paragraph' -and $style.GetAttribute('default', $context.w) -in @('1', 'true', 'on')) {
                $context.defaultParagraphStyle = $id
            }
        }
        $hidden = Get-WordXmlToggle -Node ($styles.SelectSingleNode('w:docDefaults/w:rPrDefault/w:rPr/w:vanish', $ns)) -Namespaces $ns
        if ($null -ne $hidden) { $context.defaultHidden = $hidden }
    }
    $result = @{}
    $previous = @{}
    $textCache = @{}
    for ($index = 0; $index -lt $sections.Count; $index++) {
        foreach ($area in @('header', 'footer')) {
            foreach ($variant in @('primary', 'firstPage', 'evenPages')) {
                $type = @{ primary = 'default'; firstPage = 'first'; evenPages = 'even' }[$variant]
                $key = "$area/$variant"
                $references = @($sections[$index].SelectNodes("w:${area}Reference[@w:type='$type']", $ns))
                if ($references.Count -gt 1) { throw 'Duplicate Word header/footer reference.' }
                $text = ''
                if ($references.Count -eq 1) {
                    $id = $references[0].GetAttribute('id', $ns.LookupNamespace('r'))
                    $rel = $relationships[$id]
                    if ($null -eq $rel -or $rel.GetAttribute('TargetMode') -eq 'External' -or $rel.GetAttribute('Type') -ne ($ns.LookupNamespace('r') + '/' + $area)) {
                        throw 'Invalid live Word header/footer relationship.'
                    }
                    $uri = [Uri]::new([Uri]'http://word-package/word/document.xml', $rel.GetAttribute('Target'))
                    if ($uri.Host -ne 'word-package' -or $uri.Query -or $uri.Fragment) { throw 'Invalid live Word header/footer target.' }
                    $name = [Uri]::UnescapeDataString($uri.AbsolutePath)
                    $story = $parts[$name]
                    $tag = if ($area -eq 'header') { 'hdr' } else { 'ftr' }
                    if ($null -eq $story -or $story.NamespaceURI -ne $context.w -or $story.LocalName -ne $tag) { throw 'Missing live Word header/footer part.' }
                    if (-not $textCache.ContainsKey($name)) {
                        $builder = [Text.StringBuilder]::new()
                        Add-WordXmlStoryText -Node $story -Text $builder -Context $context
                        $observed = Normalize-StoryText -Text $builder.ToString()
                        if ($observed.Length -gt 32768) { throw 'Header or footer text exceeds the supported limit.' }
                        $textCache[$name] = $observed
                    }
                    $text = $textCache[$name]
                }
                elseif ($previous.ContainsKey($key)) { $text = $previous[$key] }
                $previous[$key] = $text
                $result["$index/$key"] = $text
            }
        }
    }
    return $result
}
