function Invoke-PrepareNewDocument {
    if ($null -ne $script:Document -or $null -ne $script:CoordinationMutex) { throw 'The bridge is already bound or coordinated.' }
    $availability=Get-ExcelComAvailability
    if (-not $availability.available) { throw [ExcelActionException]::new('EXCEL_CAPABILITY_UNAVAILABLE',$availability.message) }
    $script:PreparationId='preparation-'+[guid]::NewGuid().ToString('N')
    $script:PreparedIdentity='new-'+[guid]::NewGuid().ToString('N')
    $script:PreparedPath=$null;$script:PreparedCanonicalPath=$null;$script:PreparedLocator=$null
    return [ordered]@{preparationId=$script:PreparationId;coordinationIdentity=$script:PreparedIdentity}
}

function Invoke-AcquireNewDocument {
    param($Arguments)
    if ($null -ne $script:Document -or $null -eq $script:CoordinationMutex -or [string]$Arguments.preparationId -cne $script:PreparationId -or $script:PreparedPath) { throw 'Invalid new-document preparation.' }
    $created=$false
    try { $script:Application=[Runtime.InteropServices.Marshal]::GetActiveObject('KET.Application') }
    catch { $script:Application=New-Object -ComObject 'KET.Application';$created=$true }
    $script:Documents=$script:Application.Workbooks
    $script:ActionMayHaveEffect=$true
    $script:Document=$script:Documents.Add(-4167)
    if ([int]$script:Document.Worksheets.Count -ne 1) { throw 'New workbook must have one worksheet.' }
    if (-not [string]::IsNullOrEmpty([string]$script:Document.Path)) { throw 'New document unexpectedly has a file path.' }
    $script:DocumentId='excel-'+[guid]::NewGuid().ToString('N')
    $script:UnsavedName=[string]$script:Document.Name
    Show-ExcelWorkbookWindow -ApplicationCreated $created
    return [ordered]@{documentId=$script:DocumentId;documentState=(Get-WorkbookState)}
}

function Get-ExcelPersistenceObservation {
    $items=New-Object Collections.ArrayList
    $sheets=$script:Document.Worksheets
    try {
        if ([int]$sheets.Count -gt 20) { throw [ExcelActionException]::new('RANGE_UNSUPPORTED','Persistence verification supports at most 20 worksheets.') }
        for($i=1;$i -le [int]$sheets.Count;$i++) {
            $sheet=$sheets.Item($i)
            try { [void]$items.Add((Get-ExcelSheetObservation $sheet)) }
            finally { Release-ExcelReference $sheet }
        }
    } finally { Release-ExcelReference $sheets }
    return Get-CoordinationHash -Identity (ConvertTo-Json -InputObject $items.ToArray() -Depth 30 -Compress)
}

function Invoke-ExcelPersistence {
    param([string]$Operation,$Params)
    $before=Get-ExcelPersistenceObservation
    $stateBefore=Get-WorkbookState
    if ($Operation -eq 'saveAs') {
        if ([bool]$script:Document.ReadOnly) { throw [ExcelActionException]::new('DOCUMENT_READ_ONLY','The document is read-only.') }
        Assert-SaveAsNames $Params.outputPath
        $reservation=$null;$started=$false
        try {
            $reservation=New-OutputReservation -Path $Params.outputPath -Extension '.xlsx' -ReserveFile
            $alerts=$script:Application.DisplayAlerts
            try {
                $script:Application.DisplayAlerts=0
                $script:ActionMayHaveEffect=$true;$started=$true
                $script:Document.SaveAs($reservation.path,51)|Out-Null
            } finally { $script:Application.DisplayAlerts=$alerts }
            $size=Complete-SaveAsBinding $reservation $Params.outputPath
            if ((Get-ExcelPersistenceObservation) -cne $before) { throw [PersistenceActionException]::new('OUTPUT_VERIFICATION_FAILED','Observed content changed during Save As.') }
            if ([int]$script:Document.FileFormat -ne 51) { throw 'Expected ordinary XLSX output.' }
            return [ordered]@{artifact=[ordered]@{path=[string]$Params.outputPath;format='xlsx';sizeBytes=$size};documentState=(Get-WorkbookState);replacedExisting=$false}
        } finally {
            if ($null -ne $reservation) {
                if (-not $started -and [IO.File]::Exists($reservation.path) -and (Get-StableFileIdentity $reservation.path) -eq $reservation.identity) { [IO.File]::Delete($reservation.path) }
                $reservation.handle.Dispose()
            }
        }
    }
    $format=if($Operation -eq 'exportSlideImage'){'png'}else{'pdf'}
    $reservation=$null;$temporary=$null
    try {
        $reservation=New-OutputReservation -Path $Params.outputPath -Extension ('.'+$format)
        $temporary=Join-Path ([IO.Path]::GetDirectoryName($reservation.path)) ('.wps-export-'+[guid]::NewGuid().ToString('N')+'.'+$format)
        $script:ActionMayHaveEffect=$true
        $script:Document.ExportAsFixedFormat(0,$temporary)|Out-Null
        if (-not (Test-OutputSignature $temporary $format)) { throw [PersistenceActionException]::new('OUTPUT_VERIFICATION_FAILED','Native export did not produce the requested file format.') }
        if ((Get-ExcelPersistenceObservation) -cne $before -or ((Get-WorkbookState)|ConvertTo-Json -Compress) -cne ($stateBefore|ConvertTo-Json -Compress)) {
            throw [PersistenceActionException]::new('DOCUMENT_CHANGED_DURING_ACTION','Observed document content or persistence state changed during export.')
        }
        Assert-BoundWorkbook -DocumentId $script:DocumentId
        # File.Move has no overwrite behavior; an external race cannot replace an existing output.
        [IO.File]::Move($temporary,$reservation.path)
        return [ordered]@{artifact=[ordered]@{path=[string]$Params.outputPath;format=$format;sizeBytes=[long](Get-Item $reservation.path).Length};documentStateBefore=$stateBefore;documentStateAfter=(Get-WorkbookState);replacedExisting=$false}
    } finally {
        if ($temporary -and [IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
        if ($null -ne $reservation) { $reservation.handle.Dispose() }
    }
}
