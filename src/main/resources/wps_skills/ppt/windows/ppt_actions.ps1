# PPT operations receive the exact bound Presentation. No Selection or active-document routing.
$script:PptOperations = @('getPresentationInfo','listSlides','getSlideInfo','addSlide','duplicateSlide',
    'moveSlide','deleteSlide','addTextBox','addShape','setShapeText','formatText','setShapeGeometry','deleteShape','save')

# WPS can return the same RCW for nested collection reads. Defer releases to
# the operation boundary so helpers cannot detach a caller's live reference.
$script:PptTemporaryReferences = [Collections.Generic.HashSet[object]]::new()
function Release-PptReference {
    param($Value)
    if($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void]$script:PptTemporaryReferences.Add($Value)
    }
}
function Release-PptOperationReferences {
    foreach($value in $script:PptTemporaryReferences){
        if(-not [object]::ReferenceEquals($value,$script:Document) -and
           -not [object]::ReferenceEquals($value,$script:Documents) -and
           -not [object]::ReferenceEquals($value,$script:Application)) {Release-ComReference -Value $value}
    }
    $script:PptTemporaryReferences.Clear()
}
function Release-ApplicationResources { Release-PptOperationReferences }

function Assert-PptxPackage {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip=$null; $reader=$null; $stream=$null
    try {
        $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        $zip=[IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Read,$false)
        $entry=$zip.GetEntry('[Content_Types].xml')
        if ($null -eq $entry -or $entry.Length -gt 1048576) { throw 'Invalid presentation content types.' }
        $reader=[IO.StreamReader]::new($entry.Open())
        $settings=[Xml.XmlReaderSettings]::new(); $settings.DtdProcessing=[Xml.DtdProcessing]::Prohibit
        $xmlReader=[Xml.XmlReader]::Create($reader,$settings)
        try { $xml=[Xml.XmlDocument]::new(); $xml.Load($xmlReader) } finally { $xmlReader.Dispose() }
        $main=@($xml.DocumentElement.ChildNodes | Where-Object { $_.PartName -eq '/ppt/presentation.xml' })
        if ($main.Count -ne 1 -or $main[0].ContentType -cne 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml' -or $null -ne $zip.GetEntry('ppt/vbaProject.bin')) {
            throw 'Only ordinary macro-free .pptx packages are supported.'
        }
    } catch { throw [PptActionException]::new('DOCUMENT_OPEN_FAILED', $_.Exception.Message) }
    finally { if($reader){$reader.Dispose()}; if($zip){$zip.Dispose()}; if($stream){$stream.Dispose()} }
}

function Get-PptSlide {
    param($Document,[int]$SlideId)
    $slides=$null
    try {
        $slides=$Document.Slides
        # WPS 12.1 returns null from FindBySlideID despite its type declaration.
        # Enumerate only this exact Presentation's slides and compare stable IDs.
        for($i=1;$i -le [int]$slides.Count;$i++) {
            $slide=$slides.Item($i)
            if([int]$slide.SlideID -eq $SlideId){return ,$slide}
            Release-PptReference $slide
        }
        throw [PptActionException]::new('SLIDE_NOT_FOUND','The observed slide ID is no longer present.')
    } finally { Release-PptReference $slides }
}

function Get-PptShape {
    param($Slide,[int]$ShapeId)
    $shapes=$null
    try {
        $shapes=$Slide.Shapes
        for($i=1;$i -le [int]$shapes.Count;$i++) {
            $shape=$shapes.Item($i)
            if([int]$shape.Id -eq $ShapeId){return ,$shape}
            Release-PptReference $shape
        }
        throw [PptActionException]::new('SHAPE_NOT_FOUND','The observed top-level shape ID is no longer present.')
    } finally { Release-PptReference $shapes }
}

function Get-PptSlideSummary {
    param($Slide)
    $shapes=$null
    try {
        $shapes=$Slide.Shapes
        return [ordered]@{id=[int]$Slide.SlideID;index=[int]$Slide.SlideIndex;name=[string]$Slide.Name;layout=[int]$Slide.Layout;shapeCount=[int]$shapes.Count}
    } finally {Release-PptReference $shapes}
}

function Get-PptSlidesSnapshot {
    param($Document)
    $slides=$null; $slide=$null
    try {
        $slides=$Document.Slides
        if([int]$slides.Count -gt 200){throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','At most 200 slides are supported by structure operations.')}
        $items=@(for($i=1;$i -le [int]$slides.Count;$i++) {
            $slide=$slides.Item($i)
            Get-PptSlideSummary $slide
            Release-PptReference $slide; $slide=$null
        })
        return [ordered]@{slides=$items;token=(Get-CoordinationHash -Identity ($script:TokenSalt+':slides:'+ (ConvertTo-Json -InputObject $items -Depth 8 -Compress)))}
    }finally{Release-PptReference $slide;Release-PptReference $slides}
}

function Get-PptTriState { param([int]$Value) if($Value -eq -1){return $true};if($Value -eq 0){return $false};return $null }

function Get-PptShapeSnapshot {
    param($Shape)
    $frame=$null;$range=$null;$font=$null;$color=$null
    try {
        $text=$null;$style=$null
        if([int]$Shape.HasTextFrame -eq -1) {
            $frame=$Shape.TextFrame; $range=$frame.TextRange
            $text=([string]$range.Text).Replace("`r`n","`n").Replace("`r","`n").Replace([string][char]11,"`n")
            if($text.Length -gt 10000){throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','Shape text exceeds 10000 UTF-16 units.')}
            $font=$range.Font; $color=$font.Color
            $size=if([double]$font.Size -gt 0){[double]$font.Size}else{$null}
            $rgb=if([int]$color.RGB -ge 0){[int]$color.RGB}else{$null}
            $latin=if([string]::IsNullOrEmpty([string]$font.NameAscii)){$null}else{[string]$font.NameAscii}
            $eastAsian=if([string]::IsNullOrEmpty([string]$font.NameFarEast)){$null}else{[string]$font.NameFarEast}
            $style=[ordered]@{latinName=$latin;eastAsianName=$eastAsian;size=$size;bold=(Get-PptTriState $font.Bold);italic=(Get-PptTriState $font.Italic);color=$rgb}
        }
        return [ordered]@{id=[int]$Shape.Id;name=[string]$Shape.Name;type=[int]$Shape.Type;
            left=[double]$Shape.Left;top=[double]$Shape.Top;width=[double]$Shape.Width;height=[double]$Shape.Height;
            rotation=[double]$Shape.Rotation;zOrder=[int]$Shape.ZOrderPosition;autoShapeType=$(if([int]$Shape.Type -eq 1){[int]$Shape.AutoShapeType}else{$null});text=$text;font=$style}
    }finally{Release-PptReference $color;Release-PptReference $font;Release-PptReference $range;Release-PptReference $frame}
}

function Get-PptSlideSnapshot {
    param($Slide)
    $shapes=$null;$shape=$null
    try {
        $shapes=$Slide.Shapes
        if([int]$shapes.Count -gt 100){throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','At most 100 top-level shapes can be observed per slide.')}
        $items=@(for($i=1;$i -le [int]$shapes.Count;$i++) {
            $shape=$shapes.Item($i);Get-PptShapeSnapshot $shape
            Release-PptReference $shape;$shape=$null
        })
        $snapshot=[ordered]@{slide=(Get-PptSlideSummary $Slide);shapes=$items}
        $snapshot.token=Get-CoordinationHash -Identity ($script:TokenSalt+':slide:'+ (ConvertTo-Json -InputObject $snapshot -Depth 12 -Compress))
        return $snapshot
    }finally{Release-PptReference $shape;Release-PptReference $shapes}
}

function Assert-PptToken {
    param($Snapshot,[string]$Expected)
    if($Snapshot.token -cne $Expected){throw [PptActionException]::new('STALE_CONTENT','Observed PPT fields changed. Read again and reconsider the requested edit.')}
}

function Invoke-PptAction {
    param($Document,[string]$Operation,$Params)
    $slide=$null;$shape=$null;$slides=$null;$shapes=$null;$range=$null;$frame=$null;$font=$null;$duplicate=$null;$page=$null
    try {
        switch($Operation) {
            'getPresentationInfo' {
                $page=$Document.PageSetup; $slides=$Document.Slides
                return [ordered]@{name=[string]$Document.Name;documentState=(Get-PresentationState);slideCount=[int]$slides.Count;
                    width=[double]$page.SlideWidth;height=[double]$page.SlideHeight;window=(Get-PptPresentationWindowInfo)}
            }
            'listSlides' {return Get-PptSlidesSnapshot $Document}
            'getSlideInfo' {$slide=Get-PptSlide $Document $Params.slideId;return Get-PptSlideSnapshot $slide}
        }
        if([int]$Document.ReadOnly -ne 0){throw [PptActionException]::new('DOCUMENT_READ_ONLY','The presentation is read-only.')}
        if($Operation -eq 'save') {
            if ([string]::IsNullOrEmpty($script:AuthorizedPath)) { throw [PptActionException]::new('PERSISTENCE_LOCATOR_REQUIRED','Use saveAs for the first save.') }
            Assert-BoundPresentation -DocumentId $script:DocumentId
            $script:ActionMayHaveEffect=$true
            $Document.Save()|Out-Null
            if((Get-NormalizedFileLocator -Path $script:PreparedCanonicalPath) -cne $script:PreparedLocator){throw [PptActionException]::new('DOCUMENT_BINDING_UNAVAILABLE','Save changed the authorized locator.')}
            $identity=Get-StableFileIdentity -Path $script:PreparedCanonicalPath
            Add-CoordinationFence -Identity $identity
            $script:BoundFileIdentity=$identity
            Assert-BoundPresentation -DocumentId $script:DocumentId
            $file=Get-Item -LiteralPath $script:PreparedCanonicalPath
            if([int]$Document.Saved -ne -1 -or $file.Length -lt 1){throw [PptActionException]::new('OUTPUT_VERIFICATION_FAILED','Saved state or output file could not be verified.')}
            Assert-PptxPackage -Path $script:PreparedCanonicalPath
            return [ordered]@{artifact=[ordered]@{path=$script:AuthorizedPath;format='pptx';sizeBytes=[long]$file.Length};documentState=(Get-PresentationState)}
        }
        $slides=$Document.Slides
        if($Operation -in @('addSlide','moveSlide')) {
            $before=Get-PptSlidesSnapshot $Document
            Assert-PptToken $before $Params.expectedToken
            $max=[int]$slides.Count;if($Operation -eq 'addSlide'){$max++}
            if([int]$Params.position -gt $max -or ($Operation -eq 'addSlide' -and $max -gt 200)){throw [PptActionException]::new('INVALID_PARAMS','Position or resulting slide count exceeds supported bounds.')}
            if($Operation -eq 'moveSlide'){$slide=Get-PptSlide $Document $Params.slideId}
            $script:ActionMayHaveEffect=$true
            if($Operation -eq 'addSlide'){$slide=$slides.Add([int]$Params.position,12)}else{$slide.MoveTo([int]$Params.position)|Out-Null}
            $after=Get-PptSlidesSnapshot $Document
            $expected=@($before.slides|ForEach-Object{$_.id})
            if($Operation -eq 'addSlide') {
                $id=[int]$slide.SlideID
                if($after.slides.Count -ne $before.slides.Count+1 -or $after.slides[[int]$Params.position-1].id -ne $id -or $id -in $expected){throw [PptActionException]::new('PPT_VERIFICATION_FAILED','Slide insertion did not match the requested structure.')}
            }
            return $after
        }
        $slide=Get-PptSlide $Document $Params.slideId
        $before=Get-PptSlideSnapshot $slide
        Assert-PptToken $before $Params.expectedToken
        if($Operation -in @('duplicateSlide','deleteSlide')) {
            $structure=Get-PptSlidesSnapshot $Document
            if($Operation -eq 'duplicateSlide' -and $structure.slides.Count -ge 200){throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','Duplicate would exceed 200 slides.')}
            $script:ActionMayHaveEffect=$true
            if($Operation -eq 'duplicateSlide'){$duplicate=$slide.Duplicate()}else{$slide.Delete()|Out-Null}
            $after=Get-PptSlidesSnapshot $Document
            $expectedCount=if($Operation -eq 'duplicateSlide'){$structure.slides.Count+1}else{$structure.slides.Count-1}
            if($after.slides.Count -ne $expectedCount){throw [PptActionException]::new('PPT_VERIFICATION_FAILED','Unexpected slide count after edit.')}
            if($Operation -eq 'duplicateSlide'){
                $new=@($after.slides|Where-Object{$_.id -notin @($structure.slides|ForEach-Object{$_.id})})
                if($new.Count -ne 1 -or $new[0].index -ne $before.slide.index+1){throw [PptActionException]::new('PPT_VERIFICATION_FAILED','Duplicate did not appear immediately after source.')}
            }
            return $after
        }
        $shapes=$slide.Shapes
        if($Operation -in @('addShape','addTextBox')) {
            if($before.shapes.Count -ge 100){throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','Adding a shape would exceed 100 shapes.')}
            $script:ActionMayHaveEffect=$true
            if($Operation -eq 'addTextBox'){
                $shape=$shapes.AddTextbox(1,[single]$Params.left,[single]$Params.top,[single]$Params.width,[single]$Params.height)
                $frame=$shape.TextFrame;$range=$frame.TextRange;$range.Text=([string]$Params.text).Replace("`n","`r")
            }else{
                $kind=if($Params.kind -eq 'rectangle'){1}else{9}
                $shape=$shapes.AddShape($kind,[single]$Params.left,[single]$Params.top,[single]$Params.width,[single]$Params.height)
            }
        }else{
            $shape=Get-PptShape $slide $Params.shapeId
            if($Operation -in @('setShapeText','formatText')){
                if([int]$shape.HasTextFrame -ne -1){throw [PptActionException]::new('TEXT_UNSUPPORTED','The shape has no editable top-level text frame.')}
                $frame=$shape.TextFrame;$range=$frame.TextRange
                if($Operation -eq 'formatText' -and [string]::IsNullOrEmpty([string]$range.Text)){throw [PptActionException]::new('TEXT_UNSUPPORTED','Uniform font verification requires nonempty text.')}
            }
            $script:ActionMayHaveEffect=$true
            switch($Operation){
                'deleteShape' {$shape.Delete()|Out-Null}
                'setShapeText' {$range.Text=([string]$Params.text).Replace("`n","`r")}
                'setShapeGeometry' {foreach($p in $Params.geometry.PSObject.Properties){$shape.($p.Name)=[single]$p.Value}}
                'formatText' {
                    $font=$range.Font
                    foreach($p in $Params.format.PSObject.Properties){
                        switch($p.Name){
                            'bold' {$font.Bold=if($p.Value){-1}else{0}}
                            'italic' {$font.Italic=if($p.Value){-1}else{0}}
                            'color' {$color=$font.Color;try{$color.RGB=[int]$p.Value}finally{Release-PptReference $color}}
                            'size' {$font.Size=[single]$p.Value}
                            'latinName' {$font.NameAscii=[string]$p.Value}
                            'eastAsianName' {$font.NameFarEast=[string]$p.Value}
                        }
                    }
                }
                default {throw [PptActionException]::new('PPT_CAPABILITY_UNAVAILABLE','Unknown shape edit.')}
            }
        }
        $after=Get-PptSlideSnapshot $slide
        if($Operation -in @('addShape','addTextBox')){
            $new=@($after.shapes|Where-Object{$_.id -notin @($before.shapes|ForEach-Object{$_.id})})
            if($new.Count -ne 1 -or $after.shapes.Count -ne $before.shapes.Count+1){throw [PptActionException]::new('PPT_VERIFICATION_FAILED','Exactly one new shape was not observed.')}
            foreach($key in @('left','top','width','height')){if([math]::Abs($new[0][$key]-[double]$Params.$key) -gt 0.1){throw [PptActionException]::new('PPT_VERIFICATION_FAILED','New shape geometry differs from request.')}}
            if($Operation -eq 'addTextBox' -and $new[0].text -cne $Params.text){throw [PptActionException]::new('PPT_VERIFICATION_FAILED','New text differs from request.')}
        }
        return $after
    }finally{
        foreach($reference in @($page,$duplicate,$font,$range,$frame,$shape,$shapes,$slide,$slides)){Release-PptReference $reference}
    }
}
