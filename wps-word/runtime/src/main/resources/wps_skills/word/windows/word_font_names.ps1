# Word font readback equivalence. Only names reported for one installed family
# are aliases; never construct a FontFamily from an arbitrary requested name,
# because font resolution may silently fall back to a different installed font.
$script:WordFontNameIndex = $null

function Get-InstalledWordFontFamilies {
    Add-Type -AssemblyName PresentationCore
    foreach ($family in [Windows.Media.Fonts]::SystemFontFamilies) {
        [pscustomobject]@{ Names = @($family.FamilyNames.Values) }
    }
}

function Get-WordFontNameIndex {
    if ($null -ne $script:WordFontNameIndex) {
        return $script:WordFontNameIndex
    }
    $index = [Collections.Generic.Dictionary[string,int]]::new(
        [StringComparer]::OrdinalIgnoreCase
    )
    try {
        $identity = 0
        foreach ($family in (Get-InstalledWordFontFamilies)) {
            foreach ($name in $family.Names) {
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                if ($index.ContainsKey($name) -and $index[$name] -ne $identity) {
                    # An alias shared by different families is not proof of identity.
                    $index[$name] = -1
                }
                else {
                    $index[$name] = $identity
                }
            }
            $identity++
        }
    }
    catch {
        # Discard partial catalogs and preserve strict readback verification.
        $index.Clear()
    }
    $script:WordFontNameIndex = $index
    return $index
}

function Test-WordFontNameEquivalent {
    param([string]$Requested, [string]$Observed)

    if ([string]::IsNullOrWhiteSpace($Requested) -or
        [string]::IsNullOrWhiteSpace($Observed)) { return $false }
    if ([StringComparer]::OrdinalIgnoreCase.Equals($Requested, $Observed)) {
        return $true
    }
    $index = Get-WordFontNameIndex
    return ($index.ContainsKey($Requested) -and $index.ContainsKey($Observed) -and
        $index[$Requested] -ge 0 -and $index[$Requested] -eq $index[$Observed])
}

function Assert-WordFontName {
    param([string]$Property, [string]$Requested, [string]$Observed)

    if (-not (Test-WordFontNameEquivalent -Requested $Requested -Observed $Observed)) {
        throw [ContentVerificationException]::new(
            ('The inserted font did not read back equivalently ({0}): requested "{1}", observed "{2}".' -f
                $Property, $Requested, $Observed)
        )
    }
}
