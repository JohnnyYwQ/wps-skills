function Invoke-PrepareNewDocument {
    if ($null -ne $script:Document -or $null -ne $script:CoordinationMutex) { throw 'The bridge is already bound or coordinated.' }
    $availability=Get-PptComAvailability
    if (-not $availability.available) { throw [PptActionException]::new('PPT_CAPABILITY_UNAVAILABLE',$availability.message) }
    $script:PreparationId='preparation-'+[guid]::NewGuid().ToString('N')
    $script:PreparedIdentity='new-'+[guid]::NewGuid().ToString('N')
    $script:PreparedPath=$null;$script:PreparedCanonicalPath=$null;$script:PreparedLocator=$null
    return [ordered]@{preparationId=$script:PreparationId;coordinationIdentity=$script:PreparedIdentity}
}

function Invoke-AcquireNewDocument {
    param($Arguments)
    if ($null -ne $script:Document -or $null -eq $script:CoordinationMutex -or [string]$Arguments.preparationId -cne $script:PreparationId -or $script:PreparedPath) { throw 'Invalid new-document preparation.' }
    $created=$false
    try { $script:Application=[Runtime.InteropServices.Marshal]::GetActiveObject('KWPP.Application') }
    catch { $script:Application=New-Object -ComObject 'KWPP.Application';$created=$true }
    $script:Documents=$script:Application.Presentations
    $script:ActionMayHaveEffect=$true
    $script:Document=$script:Documents.Add(-1)
    if ([int]$script:Document.Slides.Count -ne 0) { throw 'New presentation must have no slides.' }
    if (-not [string]::IsNullOrEmpty([string]$script:Document.Path)) { throw 'New document unexpectedly has a file path.' }
    $script:DocumentId='ppt-'+[guid]::NewGuid().ToString('N')
    $script:UnsavedName=[string]$script:Document.Name
    Show-PptPresentationWindow -ApplicationCreated $created
    return [ordered]@{documentId=$script:DocumentId;documentState=(Get-PresentationState)}
}

function Get-PptPersistenceObservation {
    $items=New-Object Collections.ArrayList
    $slides=$script:Document.Slides
    if ([int]$slides.Count -gt 200) { throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','Persistence verification supports at most 200 slides.') }
    for($i=1;$i -le [int]$slides.Count;$i++) {
        $slide=$slides.Item($i)
        [void]$items.Add((Get-PptSlideSnapshot $slide))
        Release-PptReference $slide
    }
    Release-PptReference $slides
    return Get-CoordinationHash -Identity (ConvertTo-Json -InputObject $items.ToArray() -Depth 30 -Compress)
}

function Invoke-PptPersistence {
    param([string]$Operation,$Params)
    $before=Get-PptPersistenceObservation
    $stateBefore=Get-PresentationState
    if ($Operation -eq 'saveAs') {
        if ([bool]$script:Document.ReadOnly) { throw [PptActionException]::new('DOCUMENT_READ_ONLY','The document is read-only.') }
        Assert-SaveAsNames $Params.outputPath
        $reservation=$null;$started=$false
        try {
            $reservation=New-OutputReservation -Path $Params.outputPath -Extension '.pptx' -ReserveFile
            $alerts=$script:Application.DisplayAlerts
            try {
                $script:Application.DisplayAlerts=1
                $script:ActionMayHaveEffect=$true;$started=$true
                $script:Document.SaveAs($reservation.path,24)|Out-Null
            } finally { $script:Application.DisplayAlerts=$alerts }
            $size=Complete-SaveAsBinding $reservation $Params.outputPath
            if ((Get-PptPersistenceObservation) -cne $before) { throw [PersistenceActionException]::new('OUTPUT_VERIFICATION_FAILED','Observed content changed during Save As.') }
            Assert-PptxPackage $reservation.path
            return [ordered]@{artifact=[ordered]@{path=[string]$Params.outputPath;format='pptx';sizeBytes=$size};documentState=(Get-PresentationState);replacedExisting=$false}
        } finally {
            if ($null -ne $reservation) {
                if (-not $started -and [IO.File]::Exists($reservation.path) -and (Get-StableFileIdentity $reservation.path) -eq $reservation.identity) { [IO.File]::Delete($reservation.path) }
                $reservation.handle.Dispose()
            }
        }
    }
    $slide=$null
    if ($Operation -eq 'exportSlideImage') { $slide=Get-PptSlide $script:Document $Params.slideId }
    $format=if($Operation -eq 'exportSlideImage'){'png'}else{'pdf'}
    $reservation=$null;$temporary=$null
    try {
        $reservation=New-OutputReservation -Path $Params.outputPath -Extension ('.'+$format)
        $temporary=Join-Path ([IO.Path]::GetDirectoryName($reservation.path)) ('.wps-export-'+[guid]::NewGuid().ToString('N')+'.'+$format)
        $script:ActionMayHaveEffect=$true
        if ($Operation -eq 'exportSlideImage') {
            $slide.Export($temporary,'PNG',[int]$Params.width,[int]$Params.height)|Out-Null
            Release-PptReference $slide
        # WPS declares ExportAsFixedFormat but its IDispatch returns MEMBERNOTFOUND.
        # Native SaveCopyAs(PDF) leaves this exact Presentation and Saved state unchanged.
        } else { $script:Document.SaveCopyAs($temporary,32)|Out-Null }
        if (-not (Test-OutputSignature $temporary $format)) { throw [PersistenceActionException]::new('OUTPUT_VERIFICATION_FAILED','Native export did not produce the requested file format.') }
        if ($format -eq 'png') {
            Add-Type -AssemblyName System.Drawing
            $bitmap=[Drawing.Image]::FromFile($temporary)
            try { if($bitmap.Width -ne $Params.width -or $bitmap.Height -ne $Params.height){throw [PersistenceActionException]::new('OUTPUT_VERIFICATION_FAILED','Exported pixel dimensions differ.')} }
            finally {$bitmap.Dispose()}
        }
        if ((Get-PptPersistenceObservation) -cne $before -or ((Get-PresentationState)|ConvertTo-Json -Compress) -cne ($stateBefore|ConvertTo-Json -Compress)) {
            throw [PersistenceActionException]::new('DOCUMENT_CHANGED_DURING_ACTION','Observed document content or persistence state changed during export.')
        }
        Assert-BoundPresentation -DocumentId $script:DocumentId
        # File.Move has no overwrite behavior; an external race cannot replace an existing output.
        [IO.File]::Move($temporary,$reservation.path)
        return [ordered]@{artifact=[ordered]@{path=[string]$Params.outputPath;format=$format;sizeBytes=[long](Get-Item $reservation.path).Length};documentStateBefore=$stateBefore;documentStateAfter=(Get-PresentationState);replacedExisting=$false}
    } finally {
        if ($temporary -and [IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
        if ($null -ne $reservation) { $reservation.handle.Dispose() }
    }
}
