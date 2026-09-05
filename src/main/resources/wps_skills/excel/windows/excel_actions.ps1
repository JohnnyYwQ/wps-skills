# Application-local Excel operations. Only the exact private bound Workbook is used.

function Release-ExcelReference {
    param($Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value) } catch {}
    }
}

$script:WorksheetBindings = @{}

function Release-ApplicationResources {
    foreach ($binding in $script:WorksheetBindings.Values) {
        Release-ExcelReference -Value $binding.sheet
    }
    $script:WorksheetBindings.Clear()
}

function Get-ExcelWorksheet {
    param([string]$Name)
    if ($script:WorksheetBindings.ContainsKey($Name)) {
        $sheet = $script:WorksheetBindings[$Name].sheet
        try {
            if ([string]$sheet.Name -cne $Name) { throw 'The worksheet was renamed.' }
        }
        catch { throw [ExcelActionException]::new('WORKSHEET_NOT_FOUND', 'The previously observed worksheet is no longer available under this name.') }
        return ,$sheet
    }
    $sheets = $null; $sheet = $null
    try {
        $sheets = $script:Document.Worksheets
        try { $sheet = $sheets.Item($Name) }
        catch { throw [ExcelActionException]::new('WORKSHEET_NOT_FOUND', 'The requested worksheet does not exist.') }
        if ([string]$sheet.Name -cne $Name) {
            throw [ExcelActionException]::new('WORKSHEET_NOT_FOUND', 'Use the exact worksheet name returned by listWorksheets.')
        }
        # Keep the exact COM worksheet alive. WPS can allocate a different COM
        # wrapper when a released worksheet is fetched again, changing IUnknown.
        $script:WorksheetBindings[$Name] = @{ sheet = $sheet; identity = [guid]::NewGuid().ToString('N') }
        $result = $sheet; $sheet = $null
        return ,$result
    }
    finally { Release-ExcelReference -Value $sheet; Release-ExcelReference -Value $sheets }
}

function Get-ExcelRectangle {
    param($Sheet, [string]$Address)
    if ($Address -cnotmatch '^[A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}[1-9][0-9]{0,6})?$') {
        throw [ExcelActionException]::new('RANGE_UNSUPPORTED', 'Expected one canonical A1 rectangle.')
    }
    $range = $null
    $areas = $null
    try {
        $range = $Sheet.Range($Address)
        $areas = $range.Areas
        if ([long]$range.Count -gt 1000 -or [int]$areas.Count -ne 1) {
            throw [ExcelActionException]::new('RANGE_UNSUPPORTED', 'The region must contain at most 1000 cells in one area.')
        }
        $result = $range; $range = $null
        return ,$result
    }
    finally { Release-ExcelReference -Value $areas; Release-ExcelReference -Value $range }
}

function Get-ExcelSnapshot {
    param($Sheet, $Range, [string]$Address)
    $rows = $null; $columns = $null; $rangeCells = $null; $functions = $null
    try {
        $rows = $Range.Rows; $columns = $Range.Columns; $rangeCells = $Range.Cells
        $functions = $script:Application.WorksheetFunction
        $rowCount = [int]$rows.Count; $columnCount = [int]$columns.Count
        $matrix = New-Object Collections.ArrayList
        for ($r = 1; $r -le $rowCount; $r++) {
            $row = New-Object Collections.ArrayList
            for ($c = 1; $c -le $columnCount; $c++) {
                $cell = $null; $font = $null; $interior = $null; $entireRow = $null; $entireColumn = $null
                try {
                    $cell = $rangeCells.GetType().InvokeMember('Item', [Reflection.BindingFlags]::GetProperty, $null, $rangeCells, [object[]]@($r, $c))
                    $font = $cell.Font; $interior = $cell.Interior
                    $entireRow=$cell.EntireRow; $entireColumn=$cell.EntireColumn
                    $value = $cell.Value2
                    $errorCode = $null
                    # VT_ERROR is distinct from a legitimate negative numeric cell.
                    if ($value -is [Runtime.InteropServices.ErrorWrapper]) {
                        $errorCode = [int]$value.ErrorCode; $value = $null
                    }
                    elseif ($null -ne $value -and [bool]$functions.IsError($cell)) {
                        $errorCode = [int]$value; $value = $null
                    }
                    elseif ($null -ne $value -and $value -isnot [string] -and $value -isnot [bool] -and
                            $value -isnot [double] -and $value -isnot [int] -and $value -isnot [decimal]) {
                        throw [ExcelActionException]::new('RANGE_UNSUPPORTED', 'WPS returned an unsupported cell value type.')
                    }
                    $formula = if ([bool]$cell.HasFormula) { [string]$cell.Formula } else { $null }
                    [void]$row.Add([ordered]@{
                        value = $value; formula = $formula; text = [string]$cell.Text; errorCode = $errorCode
                        numberFormat = [string]$cell.NumberFormat; bold = [bool]$font.Bold
                        italic = [bool]$font.Italic; fontSize = [double]$font.Size
                        fontColor = [int]$font.Color; fillColor = [int]$interior.Color
                        wrapText = [bool]$cell.WrapText
                        horizontalAlignment = Get-ExcelAlignmentName -Value ([int]$cell.HorizontalAlignment) -Horizontal $true
                        verticalAlignment = Get-ExcelAlignmentName -Value ([int]$cell.VerticalAlignment) -Horizontal $false
                        merged = [bool]$cell.MergeCells
                        rowHidden = [bool]$entireRow.Hidden; columnHidden = [bool]$entireColumn.Hidden
                        rowHeight = [double]$cell.RowHeight; columnWidth = [double]$cell.ColumnWidth
                    })
                }
                finally { Release-ExcelReference -Value $entireRow; Release-ExcelReference -Value $entireColumn; Release-ExcelReference -Value $interior; Release-ExcelReference -Value $font; Release-ExcelReference -Value $cell }
            }
            [void]$matrix.Add($row.ToArray())
        }
        $identity = $script:WorksheetBindings[[string]$Sheet.Name].identity
        $snapshot = [ordered]@{ sheet = [string]$Sheet.Name; address = $Address; cells = $matrix.ToArray() }
        $textLength = 0
        foreach ($row in $matrix) {
            foreach ($cell in $row) {
                foreach ($value in $cell.Values) {
                    if ($value -is [string]) { $textLength += $value.Length }
                }
            }
        }
        if ($textLength -gt 262144) { throw [ExcelActionException]::new('RANGE_UNSUPPORTED', 'Region text exceeds 262144 UTF-16 units; read smaller rectangles.') }
        $json = $snapshot | ConvertTo-Json -Compress -Depth 16
        $token = Get-CoordinationHash -Identity ($script:TokenSalt + ':' + $identity + ':' + $json)
        $snapshot['token'] = $token
        return $snapshot
    }
    finally {
        Release-ExcelReference -Value $functions; Release-ExcelReference -Value $rangeCells; Release-ExcelReference -Value $columns; Release-ExcelReference -Value $rows
    }
}

function Assert-ExcelEditableRectangle {
    param($Sheet, $Range)
    if ([bool]$script:Document.ReadOnly) { throw [ExcelActionException]::new('DOCUMENT_READ_ONLY', 'The workbook is read-only.') }
    if ([bool]$Sheet.ProtectContents) { throw [ExcelActionException]::new('RANGE_UNSUPPORTED', 'Protected worksheets are outside this milestone.') }
    $merge = $Range.MergeCells; $array = $Range.HasArray
    if ($null -eq $merge -or [bool]$merge -or $null -eq $array -or [bool]$array) {
        throw [ExcelActionException]::new('RANGE_UNSUPPORTED', 'Merged and array-formula cells are outside this milestone.')
    }
}

function Invoke-ExcelRegionAction {
    param([string]$Operation, $Params)
    $sheet = $null; $range = $null; $cells = $null
    try {
        $sheet = Get-ExcelWorksheet -Name ([string]$Params.sheet)
        $range = Get-ExcelRectangle -Sheet $sheet -Address ([string]$Params.address)
        $before = Get-ExcelSnapshot -Sheet $sheet -Range $range -Address ([string]$Params.address)
        if ($Operation -eq 'read_cell_rectangle') {
            # Two observations reject a visibly changing region instead of returning a mixed read.
            $again = Get-ExcelSnapshot -Sheet $sheet -Range $range -Address ([string]$Params.address)
            if ($before.token -cne $again.token) { throw [ExcelActionException]::new('STALE_RANGE', 'The region changed during reading.') }
            return $again
        }
        Assert-ExcelEditableRectangle -Sheet $sheet -Range $range
        if ([string]$Params.expectedToken -cne $before.token) {
            throw [ExcelActionException]::new('STALE_RANGE', 'Read this exact region again before editing it.')
        }
        $cells = $range.Cells
        if ($Operation -eq 'calculate_rectangle') {
            $script:ActionMayHaveEffect = $true
            $range.Calculate() | Out-Null
        }
        else {
            for ($r = 0; $r -lt $before.cells.Count; $r++) {
                for ($c = 0; $c -lt $before.cells[$r].Count; $c++) {
                    $cell = $null; $font = $null; $interior = $null; $entireRow = $null; $entireColumn = $null
                    try {
                        $cell = $cells.GetType().InvokeMember('Item', [Reflection.BindingFlags]::GetProperty, $null, $cells, [object[]]@(($r + 1), ($c + 1)))
                        switch ($Operation) {
                            'write_literal_rectangle' {
                                $value = $Params.values[$r][$c]
                                $script:ActionMayHaveEffect = $true
                                if ($null -eq $value) { $cell.ClearContents() | Out-Null }
                                else {
                                    $literal = if ($value -is [string]) { "'" + $value } elseif ($value -is [bool]) { [bool]$value } else { [double]$value }
                                    [void]$cell.GetType().InvokeMember('Value2', [Reflection.BindingFlags]::SetProperty, $null, $cell, [object[]]@($literal))
                                }
                            }
                            'write_formula_rectangle' {
                                $script:ActionMayHaveEffect = $true
                                $cell.Formula = [string]$Params.formulas[$r][$c]
                            }
                            'format_rectangle' {
                                Set-ExcelCellStyle -Cell $cell -Patch $Params.format
                                if ($null -ne $Params.format.PSObject.Properties['numberFormat']) {
                                    $script:ActionMayHaveEffect = $true
                                    $cell.NumberFormat = [string]$Params.format.numberFormat
                                }
                                if ($null -ne $Params.format.PSObject.Properties['bold']) {
                                    $font = $cell.Font
                                    $script:ActionMayHaveEffect = $true
                                    $font.Bold = [bool]$Params.format.bold
                                }
                            }
                        }
                    }
                    finally { Release-ExcelReference -Value $entireRow; Release-ExcelReference -Value $entireColumn; Release-ExcelReference -Value $interior; Release-ExcelReference -Value $font; Release-ExcelReference -Value $cell }
                }
            }
        }
        # Runtime also validates exact shape, literal values, formulas and requested format.
        return Get-ExcelSnapshot -Sheet $sheet -Range $range -Address ([string]$Params.address)
    }
    finally { Release-ExcelReference -Value $cells; Release-ExcelReference -Value $range }
}

function Invoke-ExcelAction {
    param([string]$Operation, $Params)
    switch ($Operation) {
        'inspect_workbook' {
            $sheets = $null
            try {
                $sheets = $script:Document.Worksheets
                return [ordered]@{ name = [string]$script:Document.Name; documentState = Get-WorkbookState
                    date1904 = [bool]$script:Document.Date1904; worksheetCount = [int]$sheets.Count
                    window = Get-ExcelWorkbookWindowInfo }
            }
            finally { Release-ExcelReference -Value $sheets }
        }
        'list_worksheets' {
            $sheets = $null
            try {
                $sheets = $script:Document.Worksheets
                $total = [int]$sheets.Count
                $end = [Math]::Min($total, [int]$Params.offset + [int]$Params.limit)
                $items = New-Object Collections.ArrayList
                for ($i = [int]$Params.offset + 1; $i -le $end; $i++) {
                    $sheet = $null
                    try {
                        $sheet = $sheets.Item($i)
                        [void]$items.Add([ordered]@{ name = [string]$sheet.Name; index = $i })
                    }
                    finally { Release-ExcelReference -Value $sheet }
                }
                $next = if ($end -lt $total) { $end } else { $null }
                return [ordered]@{ worksheets = $items.ToArray(); total = $total; nextOffset = $next }
            }
            finally { Release-ExcelReference -Value $sheets }
        }
        'save_existing_workbook' {
            if ([string]::IsNullOrEmpty($script:AuthorizedPath)) { throw [ExcelActionException]::new('PERSISTENCE_LOCATOR_REQUIRED','Use saveAs for the first save.') }
            if ([bool]$script:Document.ReadOnly) { throw [ExcelActionException]::new('DOCUMENT_READ_ONLY', 'The workbook is read-only.') }
            Assert-BoundWorkbook -DocumentId $script:DocumentId
            $script:ActionMayHaveEffect = $true
            $script:Document.Save() | Out-Null
            if ((Get-NormalizedFileLocator -Path $script:PreparedCanonicalPath) -cne $script:PreparedLocator) {
                throw [ExcelActionException]::new('DOCUMENT_BINDING_UNAVAILABLE', 'Save changed the authorized workbook locator.')
            }
            $savedIdentity = Get-StableFileIdentity -Path $script:PreparedCanonicalPath
            # The normalized locator fence has been held since establish. Retain
            # the old file lease while claiming each replacement file identity.
            Add-CoordinationFence -Identity $savedIdentity
            $script:BoundFileIdentity = $savedIdentity
            Assert-BoundWorkbook -DocumentId $script:DocumentId
            $file = Get-Item -LiteralPath $script:PreparedCanonicalPath
            if (-not [bool]$script:Document.Saved -or $file.Length -lt 1 -or [int]$script:Document.FileFormat -ne 51) {
                throw [ExcelActionException]::new('OUTPUT_VERIFICATION_FAILED', 'The .xlsx artifact could not be verified as saved.')
            }
            return [ordered]@{
                artifact = [ordered]@{ path = $script:AuthorizedPath; format = 'xlsx'; sizeBytes = [long]$file.Length }
                documentState = Get-WorkbookState
            }
        }
        default { return Invoke-ExcelRegionAction -Operation $Operation -Params $Params }
    }
}
