# Common editing operations stay application-local and use the bound Workbook only.
$script:ExcelCommonOperations = @('getWorksheetInfo', 'addWorksheet', 'renameWorksheet', 'copyWorksheet',
    'moveWorksheet', 'deleteWorksheet', 'insertRows', 'deleteRows', 'insertColumns', 'deleteColumns',
    'clearRange', 'mergeRange', 'unmergeRange', 'sortRange', 'filterRange', 'clearFilter',
    'setRowHeight', 'setColumnWidth', 'autoFitColumns', 'replaceInRange', 'copyRange', 'findInRange')

function Get-ExcelAlignmentName {
    param([int]$Value, [bool]$Horizontal)
    if ($Horizontal) {
        switch ($Value) { 1 { return 'general' }; -4131 { return 'left' }; -4108 { return 'center' }; -4152 { return 'right' } }
    } else {
        switch ($Value) { -4160 { return 'top' }; -4108 { return 'center' }; -4107 { return 'bottom' } }
    }
    return 'other'
}

function Set-ExcelCellStyle {
    param($Cell, $Patch)
    $font = $null; $interior = $null
    try {
        $font = $Cell.Font; $interior = $Cell.Interior
        foreach ($field in $Patch.PSObject.Properties) {
            $script:ActionMayHaveEffect = $true
            switch ($field.Name) {
                'italic' { $font.Italic = [bool]$field.Value }
                'fontSize' { $font.Size = [double]$field.Value }
                'fontColor' { $font.Color = [int]$field.Value }
                'fillColor' { $interior.Color = [int]$field.Value }
                'wrapText' { $Cell.WrapText = [bool]$field.Value }
                'horizontalAlignment' { $Cell.HorizontalAlignment = @{ general=1; left=-4131; center=-4108; right=-4152 }[[string]$field.Value] }
                'verticalAlignment' { $Cell.VerticalAlignment = @{ top=-4160; center=-4108; bottom=-4107 }[[string]$field.Value] }
            }
        }
    } finally { Release-ExcelReference $interior; Release-ExcelReference $font }
}

function Get-ExcelColumnName {
    param([int]$Index)
    $name = ''
    while ($Index -gt 0) { $Index--; $name = [string][char](65 + ($Index % 26)) + $name; $Index = [int][Math]::Floor($Index / 26) }
    return $name
}

function Get-ExcelInventory {
    $sheets = $null; $items = New-Object Collections.ArrayList
    try {
        $sheets = $script:Document.Worksheets
        for ($i=1; $i -le [int]$sheets.Count; $i++) {
            $item = $null
            try { $item = $sheets.Item($i); [void]$items.Add([ordered]@{name=[string]$item.Name; visible=([int]$item.Visible -eq -1)}) }
            finally { Release-ExcelReference $item }
        }
        return ,$items.ToArray()
    } finally { Release-ExcelReference $sheets }
}

function Get-ExcelSheetObservation {
    param($Sheet)
    $used = $null
    try {
        $used = $Sheet.UsedRange
        $address = [string]$used.Address($false, $false, 1)
        if ([long]$used.Count -gt 1000) { throw [ExcelActionException]::new('RANGE_UNSUPPORTED', 'Worksheet structure edits require a used range of at most 1000 cells.') }
        $snapshot = Get-ExcelSnapshot -Sheet $Sheet -Range $used -Address $address
        $inventory = Get-ExcelInventory
        $state = [ordered]@{inventory=$inventory; snapshot=$snapshot; protected=[bool]$Sheet.ProtectContents; structureProtected=[bool]$script:Document.ProtectStructure}
        $token = Get-CoordinationHash -Identity ($script:TokenSalt + ':' + ($state | ConvertTo-Json -Depth 20 -Compress))
        return @{ info=[ordered]@{name=[string]$Sheet.Name; index=[int]$Sheet.Index; visible=([int]$Sheet.Visible -eq -1); usedAddress=$address; token=$token}; snapshot=$snapshot; inventory=$inventory; row=[int]$used.Row; column=[int]$used.Column }
    } finally { Release-ExcelReference $used }
}

function Assert-ExcelToken {
    param([string]$Expected, [string]$Actual)
    if ($Expected -cne $Actual) { throw [ExcelActionException]::new('STALE_RANGE', 'The observed content or structure changed; read again before deciding an edit.') }
}

function Assert-ExcelSheetEditable {
    param($Sheet)
    if ([bool]$script:Document.ReadOnly) { throw [ExcelActionException]::new('DOCUMENT_READ_ONLY', 'The workbook is read-only.') }
    if ([bool]$Sheet.ProtectContents -or [bool]$script:Document.ProtectStructure) { throw [ExcelActionException]::new('RANGE_UNSUPPORTED', 'Protected content or workbook structure cannot be edited.') }
}

function Invoke-ExcelSheetAction {
    param([string]$Operation, $Params)
    $sheet = Get-ExcelWorksheet -Name $Params.sheet
    $before = Get-ExcelSheetObservation $sheet
    if ($Operation -eq 'getWorksheetInfo') {
        $again = Get-ExcelSheetObservation $sheet
        Assert-ExcelToken $before.info.token $again.info.token
        return $again.info
    }
    Assert-ExcelSheetEditable $sheet
    Assert-ExcelToken $Params.expectedToken $before.info.token
    $sheets=$null; $other=$null; $range=$null; $axis=$null; $created=$null
    $resultName = [string]$sheet.Name
    try {
        $sheets = $script:Document.Worksheets
        if ($Operation -in @('addWorksheet', 'renameWorksheet', 'copyWorksheet')) {
            if (@($before.inventory | Where-Object { $_.name -ieq $Params.name }).Count -gt 0) { throw [ExcelActionException]::new('INVALID_PARAMS', 'The requested worksheet name already exists.') }
        }
        switch ($Operation) {
            'addWorksheet' {
                $script:ActionMayHaveEffect=$true
                $created = $sheets.Add([Type]::Missing, $sheet)
                $created.Name = [string]$Params.name
                $resultName = [string]$Params.name
            }
            'renameWorksheet' {
                $script:ActionMayHaveEffect=$true; $sheet.Name=[string]$Params.name
                $resultName=[string]$Params.name
            }
            'copyWorksheet' {
                $script:ActionMayHaveEffect=$true
                $sheet.Copy([Type]::Missing, $sheet) | Out-Null
                $created = $sheets.Item([int]$sheet.Index + 1)
                $created.Name = [string]$Params.name
                $resultName=[string]$Params.name
            }
            'moveWorksheet' {
                if ([int]$Params.index -gt [int]$sheets.Count) { throw [ExcelActionException]::new('INVALID_PARAMS', 'Worksheet index exceeds worksheet count.') }
                $other=$sheets.Item([int]$Params.index)
                if ([int]$sheet.Index -ne [int]$Params.index) {
                    $script:ActionMayHaveEffect=$true
                    if ([int]$sheet.Index -gt [int]$Params.index) { $sheet.Move($other) | Out-Null }
                    else { $sheet.Move([Type]::Missing, $other) | Out-Null }
                }
            }
            'deleteWorksheet' {
                if ($before.info.visible -and @($before.inventory | Where-Object { $_.visible }).Count -le 1) { throw [ExcelActionException]::new('RANGE_UNSUPPORTED', 'Cannot delete the last visible worksheet.') }
                $alerts = $script:Application.DisplayAlerts
                try { $script:Application.DisplayAlerts=$false; $script:ActionMayHaveEffect=$true; $sheet.Delete() | Out-Null }
                finally { $script:Application.DisplayAlerts=$alerts }
            }
            default {
                $isRows = $Operation.EndsWith('Rows'); $insert = $Operation.StartsWith('insert')
                $start=[int]$Params.start; $count=[int]$Params.count
                $max = if ($isRows) { 1048576 } else { 16384 }
                $rows=$before.snapshot.cells.Count; $cols=$before.snapshot.cells[0].Count
                $endUsed = if ($isRows) { $before.row+$rows-1 } else { $before.column+$cols-1 }
                if ($start+$count-1 -gt $max -or ($insert -and $endUsed+$count -gt $max) -or
                    ($insert -and (($isRows -and ($rows+$count)*$cols -gt 1000) -or (-not $isRows -and ($cols+$count)*$rows -gt 1000)))) {
                    throw [ExcelActionException]::new('RANGE_UNSUPPORTED', 'The structural edit would exceed worksheet or observation bounds.')
                }
                $used=$null
                try { $used=$sheet.UsedRange; Assert-ExcelEditableRectangle $sheet $used }
                finally { Release-ExcelReference $used }
                $address = if ($isRows) { 'A'+$start+':A'+($start+$count-1) } else { (Get-ExcelColumnName $start)+'1:'+(Get-ExcelColumnName ($start+$count-1))+'1' }
                $range=$sheet.Range($address)
                # Keep the COM Range itself: emitting it from an if expression
                # enumerates multiple rows/columns into a PowerShell Object[].
                if ($isRows) { $axis=$range.EntireRow } else { $axis=$range.EntireColumn }
                $script:ActionMayHaveEffect=$true
                if ($insert) { $axis.Insert() | Out-Null } else { $axis.Delete() | Out-Null }
            }
        }
        # Structure changes invalidate cached worksheet locators and all previous tokens.
        Release-ApplicationResources
        if ($Operation -eq 'deleteWorksheet') {
            if (@((Get-ExcelInventory) | Where-Object { $_.name -ceq $Params.sheet }).Count -ne 0) { throw 'Worksheet deletion could not be verified.' }
            return [ordered]@{deleted=[string]$Params.sheet}
        }
        $afterSheet=Get-ExcelWorksheet $resultName
        $after=Get-ExcelSheetObservation $afterSheet
        if ($after.info.name -cne $resultName -or ($Operation -eq 'moveWorksheet' -and $after.info.index -ne $Params.index)) { throw 'Worksheet change could not be verified.' }
        if ($Operation -eq 'copyWorksheet') {
            if (($before.snapshot.cells | ConvertTo-Json -Depth 16 -Compress) -cne ($after.snapshot.cells | ConvertTo-Json -Depth 16 -Compress)) { throw 'Worksheet copy differs from the observed source.' }
        }
        if ($Operation -match '^(insert|delete)(Rows|Columns)$') {
            # Verify surviving constants at their expected shifted coordinates. Formula references may be rewritten by WPS.
            for ($r=0; $r -lt $before.snapshot.cells.Count; $r++) {
                for ($c=0; $c -lt $before.snapshot.cells[$r].Count; $c++) {
                    $source=$before.snapshot.cells[$r][$c]; $rr=$before.row+$r; $cc=$before.column+$c
                    $pos=if ($isRows) { $rr } else { $cc }
                    if (-not $insert -and $pos -ge $start -and $pos -lt $start+$count) { continue }
                    $shift=0
                    if ($pos -ge $start) { $shift=if ($insert) {$count} else {-$count} }
                    if ($isRows) {$rr+=$shift} else {$cc+=$shift}
                    $check=$null
                    try {
                        $check=$afterSheet.Range((Get-ExcelColumnName $cc)+$rr)
                        if ($null -eq $source.formula -and $null -eq $source.errorCode) {
                            if ($check.Value2 -cne $source.value) { throw 'Shifted cell value differs from the observed source.' }
                        } elseif ($null -ne $source.formula -and -not [bool]$check.HasFormula) { throw 'Shifted formula was lost.' }
                    } finally { Release-ExcelReference $check }
                }
            }
        }
        return $after.info
    } finally { Release-ExcelReference $axis; Release-ExcelReference $range; Release-ExcelReference $created; Release-ExcelReference $other; Release-ExcelReference $sheets }
}

function Test-ExcelTextMatch {
    param([string]$Value, [string]$Text, [bool]$MatchCase, [bool]$WholeCell)
    $comparison=if ($MatchCase) {[StringComparison]::Ordinal} else {[StringComparison]::OrdinalIgnoreCase}
    if ($WholeCell) { return [string]::Equals($Value,$Text,$comparison) }
    return $Value.IndexOf($Text,$comparison) -ge 0
}

function Invoke-ExcelCommonAction {
    param([string]$Operation, $Params)
    if ($Operation -in @('getWorksheetInfo','addWorksheet','renameWorksheet','copyWorksheet','moveWorksheet','deleteWorksheet','insertRows','deleteRows','insertColumns','deleteColumns')) {
        return Invoke-ExcelSheetAction $Operation $Params
    }
    $sheet=Get-ExcelWorksheet $Params.sheet
    $range=$null; $target=$null; $key=$null; $axis=$null; $filter=$null; $filterRange=$null
    try {
        $range=Get-ExcelRectangle $sheet $Params.address
        $before=Get-ExcelSnapshot $sheet $range $Params.address
        if ($Operation -eq 'findInRange') {
            $matches=New-Object Collections.ArrayList
            for ($r=0; $r -lt $before.cells.Count; $r++) {
                for ($c=0; $c -lt $before.cells[$r].Count; $c++) {
                    $cell=$before.cells[$r][$c]
                    $text=if ($Params.lookIn -eq 'formulas') {[string]$cell.formula} else {[string]$cell.value}
                    if (Test-ExcelTextMatch $text $Params.text $Params.matchCase $Params.wholeCell) {
                        [void]$matches.Add([ordered]@{address=(Get-ExcelColumnName ([int]$range.Column+$c))+([int]$range.Row+$r); text=$text})
                    }
                }
            }
            $again=Get-ExcelSnapshot $sheet $range $Params.address
            Assert-ExcelToken $before.token $again.token
            return [ordered]@{sheet=$Params.sheet; address=$Params.address; token=$again.token; matches=$matches.ToArray()}
        }
        Assert-ExcelSheetEditable $sheet
        Assert-ExcelToken $Params.expectedToken $before.token
        if ($Operation -ne 'unmergeRange') { Assert-ExcelEditableRectangle $sheet $range }
        if ($Operation -eq 'copyRange') {
            $targetSheet=Get-ExcelWorksheet $Params.targetSheet
            $target=Get-ExcelRectangle $targetSheet $Params.targetAddress
            Assert-ExcelEditableRectangle $targetSheet $target
            $observed=Get-ExcelSnapshot $targetSheet $target $Params.targetAddress
            Assert-ExcelToken $Params.targetToken $observed.token
            $values=New-Object Collections.ArrayList
            foreach ($row in $before.cells) {
                $items=New-Object Collections.ArrayList
                foreach ($cell in $row) {
                    if ($null -ne $cell.errorCode) { throw [ExcelActionException]::new('RANGE_UNSUPPORTED','Copy values cannot copy error cells.') }
                    [void]$items.Add($cell.value)
                }
                [void]$values.Add($items.ToArray())
            }
            $copied=Invoke-ExcelRegionAction 'write_literal_rectangle' ([pscustomobject]@{sheet=$Params.targetSheet; address=$Params.targetAddress; expectedToken=$Params.targetToken; values=$values.ToArray()})
            for ($r=0; $r -lt $before.cells.Count; $r++) { for ($c=0; $c -lt $before.cells[$r].Count; $c++) {
                $actual=$copied.cells[$r][$c]
                if ($actual.value -cne $before.cells[$r][$c].value -or $null -ne $actual.formula -or $null -ne $actual.errorCode -or
                    ($actual.value -is [bool]) -ne ($before.cells[$r][$c].value -is [bool])) {throw 'Copied values could not be verified.'}
            } }
            return $copied
        }
        if ($Operation -eq 'sortRange') {
            foreach ($row in $before.cells) { foreach ($cell in $row) {
                if ($null -ne $cell.formula -or $null -ne $cell.errorCode) { throw [ExcelActionException]::new('RANGE_UNSUPPORTED','Sorting currently supports constant values without formulas or error cells.') }
            } }
        }
        if ($Operation -eq 'mergeRange') {
            for ($r=0; $r -lt $before.cells.Count; $r++) { for ($c=0; $c -lt $before.cells[$r].Count; $c++) {
                if (($r -ne 0 -or $c -ne 0) -and ($null -ne $before.cells[$r][$c].value -or $null -ne $before.cells[$r][$c].formula -or $null -ne $before.cells[$r][$c].errorCode)) {
                    throw [ExcelActionException]::new('RANGE_UNSUPPORTED','Merging would discard nonempty cells.')
                }
            } }
        }
        if ($Operation -eq 'unmergeRange') {
            foreach ($cell in $range.Cells) {
                $area=$null
                try {
                    $area=$cell.MergeArea
                    if ([int]$area.Row -lt [int]$range.Row -or [int]$area.Column -lt [int]$range.Column -or
                        [int]$area.Row+[int]$area.Rows.Count -gt [int]$range.Row+$before.cells.Count -or
                        [int]$area.Column+[int]$area.Columns.Count -gt [int]$range.Column+$before.cells[0].Count) {
                        throw [ExcelActionException]::new('RANGE_UNSUPPORTED','The rectangle must contain every complete merged area.')
                    }
                } finally { Release-ExcelReference $area; Release-ExcelReference $cell }
            }
        }
        if ($Operation -in @('filterRange','clearFilter') -and [bool]$sheet.AutoFilterMode) {
            $filter=$sheet.AutoFilter; $filterRange=$filter.Range
            if ([string]$filterRange.Address($false,$false,1) -cne [string]$range.Address($false,$false,1)) { throw [ExcelActionException]::new('RANGE_UNSUPPORTED','An existing filter belongs to another rectangle.') }
        }
        switch ($Operation) {
            'clearRange' {
                $script:ActionMayHaveEffect=$true
                switch ($Params.mode) { 'contents' { $range.ClearContents() | Out-Null }; 'formats' { $range.ClearFormats() | Out-Null }; 'all' { $range.Clear() | Out-Null } }
            }
            'mergeRange' { $script:ActionMayHaveEffect=$true; $range.Merge() | Out-Null }
            'unmergeRange' { $script:ActionMayHaveEffect=$true; $range.UnMerge() | Out-Null }
            'sortRange' {
                $key=$sheet.Range((Get-ExcelColumnName ([int]$range.Column+[int]$Params.column-1))+[int]$range.Row)
                $order=if ($Params.order -eq 'ascending') {1} else {2}; $header=if ($Params.header) {1} else {2}
                $script:ActionMayHaveEffect=$true
                $range.Sort($key,$order,[Type]::Missing,[Type]::Missing,[Type]::Missing,[Type]::Missing,[Type]::Missing,$header) | Out-Null
            }
            'filterRange' {
                $criteria='='+([string]$Params.value).Replace('~','~~').Replace('*','~*').Replace('?','~?')
                $script:ActionMayHaveEffect=$true
                $range.AutoFilter([int]$Params.column,$criteria) | Out-Null
            }
            'clearFilter' { $script:ActionMayHaveEffect=$true; if ([bool]$sheet.FilterMode) {$sheet.ShowAllData() | Out-Null}; $sheet.AutoFilterMode=$false }
            'setRowHeight' { $script:ActionMayHaveEffect=$true; $range.RowHeight=[double]$Params.height }
            'setColumnWidth' { $script:ActionMayHaveEffect=$true; $range.ColumnWidth=[double]$Params.width }
            'autoFitColumns' { $axis=$range.Columns; $script:ActionMayHaveEffect=$true; $axis.AutoFit() | Out-Null }
            'replaceInRange' {
                # Regex is used only after escaping the caller's literal text; replacement is literal too.
                $options=if ($Params.matchCase) {[Text.RegularExpressions.RegexOptions]::None} else {[Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant}
                $pattern=[regex]::Escape([string]$Params.text)
                if ($Params.wholeCell) {$pattern='\A'+$pattern+'\z'}
                $regex=[regex]::new($pattern,$options)
                $replacement=([string]$Params.replacement).Replace('$','$$')
                for ($r=0; $r -lt $before.cells.Count; $r++) { for ($c=0; $c -lt $before.cells[$r].Count; $c++) {
                    $old=$before.cells[$r][$c]
                    if ($null -ne $old.formula -or $old.value -isnot [string]) { continue }
                    $value=$regex.Replace([string]$old.value,$replacement)
                    if ($value.Length -gt 4096) { throw [ExcelActionException]::new('RANGE_UNSUPPORTED','Replacement would exceed the literal write limit.') }
                } }
                for ($r=0; $r -lt $before.cells.Count; $r++) { for ($c=0; $c -lt $before.cells[$r].Count; $c++) {
                    $old=$before.cells[$r][$c]
                    if ($null -ne $old.formula -or $old.value -isnot [string]) { continue }
                    $value=$regex.Replace([string]$old.value,$replacement)
                    if ($value -ceq $old.value) {continue}
                    $cell=$null
                    try {
                        $cell=$sheet.Range((Get-ExcelColumnName ([int]$range.Column+$c))+([int]$range.Row+$r))
                        $script:ActionMayHaveEffect=$true
                        [void]$cell.GetType().InvokeMember('Value2',[Reflection.BindingFlags]::SetProperty,$null,$cell,[object[]]@("'"+$value))
                    } finally { Release-ExcelReference $cell }
                } }
            }
        }
        $after=Get-ExcelSnapshot $sheet $range $Params.address
        if ($Operation -eq 'replaceInRange') {
            for ($r=0; $r -lt $before.cells.Count; $r++) { for ($c=0; $c -lt $before.cells[$r].Count; $c++) {
                $old=$before.cells[$r][$c]; $actual=$after.cells[$r][$c]
                $wanted=if ($null -eq $old.formula -and $old.value -is [string]) {$regex.Replace([string]$old.value,$replacement)} else {$old.value}
                if ($actual.value -cne $wanted -or $actual.formula -cne $old.formula -or $actual.errorCode -ne $old.errorCode) {throw 'Replacement verification failed.'}
            } }
        }
        if ($Operation -eq 'sortRange') {
            $first=if ($Params.header) {1} else {0}
            if ($first -eq 1 -and ($before.cells[0] | ConvertTo-Json -Depth 16 -Compress) -cne ($after.cells[0] | ConvertTo-Json -Depth 16 -Compress)) {throw 'Sort changed the header.'}
            $oldRows=New-Object Collections.ArrayList; $newRows=New-Object Collections.ArrayList
            for ($r=$first; $r -lt $before.cells.Count; $r++) {
                [void]$oldRows.Add((@($before.cells[$r] | ForEach-Object {$_.value}) | ConvertTo-Json -Compress))
                [void]$newRows.Add((@($after.cells[$r] | ForEach-Object {$_.value}) | ConvertTo-Json -Compress))
            }
            if ((($oldRows | Sort-Object) -join "`n") -cne (($newRows | Sort-Object) -join "`n")) {throw 'Sort did not preserve complete rows.'}
            # WPS defines locale-dependent text ordering. For numeric keys, independently prove order too.
            for ($r=$first+1; $r -lt $after.cells.Count; $r++) {
                $prev=$after.cells[$r-1][[int]$Params.column-1].value; $next=$after.cells[$r][[int]$Params.column-1].value
                if ($null -ne $prev -and $null -ne $next -and $prev -isnot [string] -and $next -isnot [string]) {
                    if (($Params.order -eq 'ascending' -and [double]$prev -gt [double]$next) -or ($Params.order -eq 'descending' -and [double]$prev -lt [double]$next)) {throw 'Numeric sort order verification failed.'}
                }
            }
        }
        if ($Operation -eq 'setRowHeight' -and [Math]::Abs([double]$after.cells[0][0].rowHeight-[double]$Params.height) -gt 0.8) {throw 'Row height verification failed.'}
        if ($Operation -eq 'setColumnWidth' -and [Math]::Abs([double]$after.cells[0][0].columnWidth-[double]$Params.width) -gt 0.2) {throw 'Column width verification failed.'}
        if ($Operation -eq 'filterRange' -and -not [bool]$sheet.AutoFilterMode) {throw 'Filter verification failed.'}
        if ($Operation -eq 'clearFilter' -and ([bool]$sheet.FilterMode -or [bool]$sheet.AutoFilterMode)) {throw 'Filter removal verification failed.'}
        return $after
    } finally { Release-ExcelReference $filterRange; Release-ExcelReference $filter; Release-ExcelReference $axis; Release-ExcelReference $key; Release-ExcelReference $target; Release-ExcelReference $range }
}
