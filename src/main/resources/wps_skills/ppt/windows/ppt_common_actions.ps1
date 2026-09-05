# Additional application-local PPT operations. References remain owned by the
# current operation and are released by the same bridge boundary as base Actions.
$script:PptCommonOperations=@('getShapeStyle','formatShape','formatParagraph','setTextBoxLayout',
    'renameShape','setShapeOrder','alignShapes','distributeShapes','getSlideSettings','setSlideSettings',
    'getSlideNotes','setSlideNotes','findText','replaceText','addImage','addTable','readTable','writeTable')

function ConvertTo-PptText {
    param($Value)
    return ([string]$Value).Replace("`r`n","`n").Replace("`r","`n").Replace([string][char]11,"`n")
}
function Get-PptColor {
    param($Color)
    try { $rgb=[int]$Color.RGB;if($rgb -ge 0 -and $rgb -le 16777215){return $rgb};return $null }
    finally{Release-PptReference $Color}
}
function New-PptObservedToken {
    param([string]$Scope,$Data)
    return Get-CoordinationHash -Identity ($script:TokenSalt+':'+$Scope+':'+(ConvertTo-Json -InputObject $Data -Depth 20 -Compress))
}
function Get-PptShapeStyle {
    param([int]$SlideId,$Shape)
    $fill=$null;$line=$null;$frame=$null;$range=$null;$paragraphs=$null;$paragraph=$null;$format=$null;$bullet=$null
    try {
        if([int]$Shape.Type -in @(6,19)){throw [PptActionException]::new('SHAPE_UNSUPPORTED','Style Actions support individual top-level shapes, not groups or tables.')}
        $fill=$Shape.Fill;$line=$Shape.Line
        $appearance=[ordered]@{fillVisible=(Get-PptTriState $fill.Visible);fillType=[int]$fill.Type;
            fillColor=$(if([int]$fill.Visible -eq -1 -and [int]$fill.Type -eq 1){Get-PptColor $fill.ForeColor}else{$null});fillTransparency=[double]$fill.Transparency;
            lineVisible=(Get-PptTriState $line.Visible);lineColor=$(if([int]$line.Visible -eq -1){Get-PptColor $line.ForeColor}else{$null});lineWidth=[double]$line.Weight}
        $layout=$null;$items=@()
        if([int]$Shape.HasTextFrame -eq -1){
            $frame=$Shape.TextFrame;$range=$frame.TextRange
            $layout=[ordered]@{marginLeft=[double]$frame.MarginLeft;marginRight=[double]$frame.MarginRight;
                marginTop=[double]$frame.MarginTop;marginBottom=[double]$frame.MarginBottom;
                verticalAnchor=[int]$frame.VerticalAnchor;wordWrap=(Get-PptTriState $frame.WordWrap)}
            $paragraphs=$range.Paragraphs()
            if([int]$paragraphs.Count -gt 100){throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','Style observation supports at most 100 paragraphs.')}
            $items=@(for($i=1;$i -le [int]$paragraphs.Count;$i++){
                $paragraph=$range.Paragraphs($i,1);$format=$paragraph.ParagraphFormat;$bullet=$format.Bullet
                [ordered]@{alignment=[int]$format.Alignment;spaceBefore=[double]$format.SpaceBefore;spaceAfter=[double]$format.SpaceAfter;
                    spaceBeforeInLines=(Get-PptTriState $format.LineRuleBefore);spaceAfterInLines=(Get-PptTriState $format.LineRuleAfter);
                    bulletVisible=(Get-PptTriState $bullet.Visible);bulletType=[int]$bullet.Type}
                Release-PptReference $bullet;Release-PptReference $format;Release-PptReference $paragraph
            })
        }
        $data=[ordered]@{slideId=$SlideId;shape=(Get-PptShapeSnapshot $Shape);appearance=$appearance;textBox=$layout;paragraphs=$items}
        $data.token=New-PptObservedToken 'shape-style' $data
        return $data
    }finally{foreach($ref in @($bullet,$format,$paragraph,$paragraphs,$range,$frame,$line,$fill)){Release-PptReference $ref}}
}
function Get-PptSlideSettings {
    param($Slide)
    $transition=$null;$background=$null;$fill=$null
    try{
        $transition=$Slide.SlideShowTransition;$background=$Slide.Background;$fill=$background.Fill
        $data=[ordered]@{slideId=[int]$Slide.SlideID;name=[string]$Slide.Name;hidden=([int]$transition.Hidden -eq -1);
            followMasterBackground=([int]$Slide.FollowMasterBackground -eq -1);backgroundType=[int]$fill.Type;
            backgroundColor=(Get-PptColor $fill.ForeColor)}
        $data.token=New-PptObservedToken 'slide-settings' $data;return $data
    }finally{foreach($ref in @($fill,$background,$transition)){Release-PptReference $ref}}
}
function Get-PptNotesBody {
    param($Slide)
    $notes=$null;$shapes=$null;$shape=$null;$placeholder=$null;$body=$null
    try{
        $notes=$Slide.NotesPage;$shapes=$notes.Shapes
        for($i=1;$i -le [int]$shapes.Count;$i++){
            $shape=$shapes.Item($i)
            if([int]$shape.Type -eq 14){
                $placeholder=$shape.PlaceholderFormat
                if([int]$placeholder.Type -eq 2 -and [int]$shape.HasTextFrame -eq -1){
                    if($null -ne $body){throw [PptActionException]::new('NOTES_UNSUPPORTED','More than one notes body placeholder was found.')}
                    $body=$shape
                }
            }
            Release-PptReference $placeholder;Release-PptReference $shape
        }
        if($null -eq $body){throw [PptActionException]::new('NOTES_UNSUPPORTED','The slide has no unambiguous notes body placeholder.')}
        return ,$body
    }finally{foreach($ref in @($placeholder,$shape,$shapes,$notes)){Release-PptReference $ref}}
}
function Get-PptNotesSnapshot {
    param($Slide,$Body)
    $frame=$null;$range=$null
    try{
        $frame=$Body.TextFrame;$range=$frame.TextRange;$text=ConvertTo-PptText $range.Text
        if($text.Length -gt 10000){throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','Notes body exceeds 10000 UTF-16 units.')}
        $data=[ordered]@{slideId=[int]$Slide.SlideID;text=$text}
        $data.token=New-PptObservedToken ('notes-'+[int]$Body.Id) $data;return $data
    }finally{Release-PptReference $range;Release-PptReference $frame}
}
function Get-PptTableSnapshot {
    param([int]$SlideId,$Shape)
    $table=$null;$rows=$null;$columns=$null;$cell=$null;$cellShape=$null;$frame=$null;$range=$null
    try{
        if([int]$Shape.HasTable -ne -1){throw [PptActionException]::new('TABLE_UNSUPPORTED','The top-level shape is not a native table.')}
        $table=$Shape.Table;$rows=$table.Rows;$columns=$table.Columns
        $nr=[int]$rows.Count;$nc=[int]$columns.Count
        if($nr -lt 1 -or $nc -lt 1 -or $nr -gt 20 -or $nc -gt 10 -or $nr*$nc -gt 100){throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','Table exceeds the 20-row, 10-column or 100-cell limit.')}
        $values=[Collections.Generic.List[object]]::new();$units=0
        for($r=1;$r -le $nr;$r++){
            $row=[Collections.Generic.List[object]]::new()
            for($c=1;$c -le $nc;$c++){
                $cell=$table.Cell($r,$c);$cellShape=$cell.Shape;$frame=$cellShape.TextFrame;$range=$frame.TextRange
                $value=ConvertTo-PptText $range.Text;$units+=$value.Length
                if($value.Length -gt 2000 -or $units -gt 20000){throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','Table text exceeds the per-cell or total text limit.')}
                $row.Add($value)
                foreach($ref in @($range,$frame,$cellShape,$cell)){Release-PptReference $ref}
            }
            $values.Add($row.ToArray())
        }
        $data=[ordered]@{slideId=$SlideId;shapeId=[int]$Shape.Id;rows=$nr;columns=$nc;values=$values.ToArray()}
        $data.token=New-PptObservedToken 'table' $data;return $data
    }finally{foreach($ref in @($range,$frame,$cellShape,$cell,$columns,$rows,$table)){Release-PptReference $ref}}
}
function Set-PptTableValues {
    param($Shape,$Values)
    $table=$null;$cell=$null;$shapeCell=$null;$frame=$null;$range=$null
    try{
        $table=$Shape.Table
        for($r=0;$r -lt $Values.Count;$r++){
            for($c=0;$c -lt $Values[$r].Count;$c++){
                $cell=$table.Cell($r+1,$c+1);$shapeCell=$cell.Shape;$frame=$shapeCell.TextFrame;$range=$frame.TextRange
                $range.Text=([string]$Values[$r][$c]).Replace("`n","`r")
                foreach($ref in @($range,$frame,$shapeCell,$cell)){Release-PptReference $ref}
            }
        }
    }finally{foreach($ref in @($range,$frame,$shapeCell,$cell,$table)){Release-PptReference $ref}}
}
function Get-PptLiteralMatches {
    param([string]$Text,[string]$Find)
    $positions=[Collections.Generic.List[int]]::new();$offset=0
    while($offset -le $Text.Length-$Find.Length){
        $index=$Text.IndexOf($Find,$offset,[StringComparison]::Ordinal)
        if($index -lt 0){break};$positions.Add($index);$offset=$index+$Find.Length
        if($positions.Count -gt 100){throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','More than 100 text matches were found.')}
    }
    return ,$positions.ToArray()
}
function Invoke-PptCommonAction {
    param($Document,[string]$Operation,$Params)
    $slide=$null;$shape=$null;$shapes=$null;$fill=$null;$line=$null;$color=$null;$frame=$null;$range=$null;
    $format=$null;$bullet=$null;$body=$null;$transition=$null;$background=$null;$image=$null;$stream=$null;$segment=$null
    try{
        $slide=Get-PptSlide $Document $Params.slideId
        if($Operation -in @('getShapeStyle','formatShape','formatParagraph','setTextBoxLayout','readTable','writeTable')){$shape=Get-PptShape $slide $Params.shapeId}
        switch($Operation){
            'getShapeStyle' {return Get-PptShapeStyle $Params.slideId $shape}
            'getSlideSettings' {return Get-PptSlideSettings $slide}
            'getSlideNotes' {$body=Get-PptNotesBody $slide;return Get-PptNotesSnapshot $slide $body}
            'readTable' {return Get-PptTableSnapshot $Params.slideId $shape}
            'findText' {
                $before=Get-PptSlideSnapshot $slide;$matches=[Collections.Generic.List[object]]::new()
                foreach($s in $before.shapes){if($null -ne $s.text){foreach($offset in (Get-PptLiteralMatches $s.text $Params.text)){
                    $matches.Add([ordered]@{shapeId=$s.id;start=$offset;length=([string]$Params.text).Length;text=[string]$Params.text})
                    if($matches.Count -gt 100){throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','More than 100 slide text matches were found.')}
                }}}
                return [ordered]@{slideId=$Params.slideId;matches=$matches.ToArray();token=$before.token}
            }
        }
        if([int]$Document.ReadOnly -ne 0){throw [PptActionException]::new('DOCUMENT_READ_ONLY','The presentation is read-only.')}
        if($Operation -in @('formatShape','formatParagraph','setTextBoxLayout')){
            $before=Get-PptShapeStyle $Params.slideId $shape;Assert-PptToken $before $Params.expectedToken
            if($Operation -in @('formatParagraph','setTextBoxLayout') -and $null -eq $before.textBox){throw [PptActionException]::new('TEXT_UNSUPPORTED','The shape has no text frame.')}
            if($Operation -eq 'formatParagraph' -and $before.paragraphs.Count -eq 0){throw [PptActionException]::new('TEXT_UNSUPPORTED','No paragraphs were observed.')}
            if($Operation -eq 'formatShape'){
                $fill=$shape.Fill;$line=$shape.Line
                $script:ActionMayHaveEffect=$true
                # WPS cannot expose a reliable color for hidden fills/borders. Color edits
                # enable visibility; conflicting explicit hidden patches are rejected by the contract.
                if($null -ne $Params.format.fillColor){$fill.Solid()|Out-Null;$color=$fill.ForeColor;$color.RGB=[int]$Params.format.fillColor;Release-PptReference $color}
                if($null -ne $Params.format.fillTransparency){$fill.Visible=-1;$fill.Transparency=[single]$Params.format.fillTransparency}
                if($null -ne $Params.format.lineColor){$line.Visible=-1;$color=$line.ForeColor;$color.RGB=[int]$Params.format.lineColor;Release-PptReference $color}
                if($null -ne $Params.format.lineWidth){$line.Weight=[single]$Params.format.lineWidth}
                if($null -ne $Params.format.fillVisible){$fill.Visible=if($Params.format.fillVisible){-1}else{0}}
                if($null -ne $Params.format.lineVisible){$line.Visible=if($Params.format.lineVisible){-1}else{0}}
            }else{
                $frame=$shape.TextFrame;$range=$frame.TextRange
                if($Operation -eq 'setTextBoxLayout'){
                    $script:ActionMayHaveEffect=$true
                    foreach($p in $Params.layout.PSObject.Properties){
                        if($p.Name -eq 'verticalAnchor'){$frame.VerticalAnchor=@{top=1;middle=3;bottom=4}[[string]$p.Value]}
                        elseif($p.Name -eq 'wordWrap'){$frame.WordWrap=if($p.Value){-1}else{0}}
                        else{$frame.($p.Name)=[single]$p.Value}
                    }
                }else{
                    $format=$range.ParagraphFormat;$script:ActionMayHaveEffect=$true
                    if($null -ne $Params.format.alignment){$format.Alignment=@{left=1;center=2;right=3;justify=4}[[string]$Params.format.alignment]}
                    if($null -ne $Params.format.spaceBefore){$format.LineRuleBefore=0;$format.SpaceBefore=[single]$Params.format.spaceBefore}
                    if($null -ne $Params.format.spaceAfter){$format.LineRuleAfter=0;$format.SpaceAfter=[single]$Params.format.spaceAfter}
                    if($null -ne $Params.format.bulletVisible){$bullet=$format.Bullet;$bullet.Visible=if($Params.format.bulletVisible){-1}else{0}}
                }
            }
            return Get-PptShapeStyle $Params.slideId $shape
        }
        if($Operation -eq 'setSlideSettings'){
            $before=Get-PptSlideSettings $slide;Assert-PptToken $before $Params.expectedToken
            $script:ActionMayHaveEffect=$true
            if($null -ne $Params.settings.name){$slide.Name=[string]$Params.settings.name}
            if($null -ne $Params.settings.hidden){$transition=$slide.SlideShowTransition;$transition.Hidden=if($Params.settings.hidden){-1}else{0}}
            if($null -ne $Params.settings.backgroundColor){
                $slide.FollowMasterBackground=0;$background=$slide.Background;$fill=$background.Fill;$fill.Solid()|Out-Null
                $color=$fill.ForeColor;$color.RGB=[int]$Params.settings.backgroundColor
            }elseif($null -ne $Params.settings.followMasterBackground){$slide.FollowMasterBackground=if($Params.settings.followMasterBackground){-1}else{0}}
            return Get-PptSlideSettings $slide
        }
        if($Operation -eq 'setSlideNotes'){
            $body=Get-PptNotesBody $slide;$before=Get-PptNotesSnapshot $slide $body;Assert-PptToken $before $Params.expectedToken
            $frame=$body.TextFrame;$range=$frame.TextRange;$script:ActionMayHaveEffect=$true;$range.Text=([string]$Params.text).Replace("`n","`r")
            return Get-PptNotesSnapshot $slide $body
        }
        if($Operation -eq 'writeTable'){
            $before=Get-PptTableSnapshot $Params.slideId $shape;Assert-PptToken $before $Params.expectedToken
            if($Params.values.Count -ne $before.rows -or $Params.values[0].Count -ne $before.columns){throw [PptActionException]::new('INVALID_PARAMS','Table write requires exactly the existing row and column counts.')}
            $script:ActionMayHaveEffect=$true;Set-PptTableValues $shape $Params.values
            return Get-PptTableSnapshot $Params.slideId $shape
        }
        $before=Get-PptSlideSnapshot $slide;Assert-PptToken $before $Params.expectedToken
        $shapes=$slide.Shapes
        if($Operation -in @('addImage','addTable')){
            if($before.shapes.Count -ge 100){throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','Insertion would exceed 100 top-level shapes.')}
            if($Operation -eq 'addTable'){
                $script:ActionMayHaveEffect=$true
                $shape=$shapes.AddTable([int]$Params.values.Count,[int]$Params.values[0].Count,[single]$Params.left,[single]$Params.top,[single]$Params.width,[single]$Params.height)
                Set-PptTableValues $shape $Params.values
                $after=Get-PptSlideSnapshot $slide
                if($after.shapes.Count -ne $before.shapes.Count+1 -or [int]$shape.Id -in @($before.shapes|ForEach-Object{$_.id})){throw [PptActionException]::new('PPT_VERIFICATION_FAILED','Exactly one new table was not observed.')}
                return Get-PptTableSnapshot $Params.slideId $shape
            }
            if(-not(Test-Path -LiteralPath $Params.path -PathType Leaf)){throw [PptActionException]::new('IMAGE_NOT_FOUND','The authorized image file does not exist.')}
            try{
                $stream=[IO.File]::Open([string]$Params.path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
                if($stream.Length -gt 20971520){throw 'Image exceeds 20 MiB.'}
                Add-Type -AssemblyName System.Drawing
                $image=[Drawing.Image]::FromStream($stream,$false,$true)
                if($image.Width*$image.Height -gt 40000000 -or $image.RawFormat.Guid -notin @([Drawing.Imaging.ImageFormat]::Png.Guid,[Drawing.Imaging.ImageFormat]::Jpeg.Guid)){throw 'Expected PNG/JPEG with at most 40 million pixels.'}
            }catch{throw [PptActionException]::new('IMAGE_UNSUPPORTED',$_.Exception.Message)}
            $script:ActionMayHaveEffect=$true
            $shape=$shapes.AddPicture([string]$Params.path,0,-1,[single]$Params.left,[single]$Params.top,[single]$Params.width,[single]$Params.height)
            $after=Get-PptSlideSnapshot $slide
            $new=@($after.shapes|Where-Object{$_.id -notin @($before.shapes|ForEach-Object{$_.id})})
            if($new.Count -ne 1 -or $new[0].type -ne 13 -or $after.shapes.Count -ne $before.shapes.Count+1){throw [PptActionException]::new('PPT_VERIFICATION_FAILED','Exactly one embedded picture was not observed.')}
            foreach($key in @('left','top','width','height')){if([math]::Abs($new[0][$key]-[double]$Params.$key) -gt 0.1){throw [PptActionException]::new('PPT_VERIFICATION_FAILED','Picture dimensions differ from request.')}}
            return $after
        }
        if($Operation -in @('alignShapes','distributeShapes')){
            $selected=@($before.shapes|Where-Object{$_.id -in $Params.shapeIds})
            if($selected.Count -ne $Params.shapeIds.Count){throw [PptActionException]::new('SHAPE_NOT_FOUND','One of the requested shape IDs is absent.')}
            if(@($selected|Where-Object{[math]::Abs($_.rotation) -gt 0.01 -or $_.type -eq 6}).Count){throw [PptActionException]::new('SHAPE_UNSUPPORTED','Alignment/distribution supports unrotated, ungrouped top-level shapes.')}
            $axis=if($Operation -eq 'alignShapes'){if($Params.alignment -in @('left','center','right')){'left'}else{'top'}}else{if($Params.direction -eq 'horizontal'){'left'}else{'top'}}
            $size=if($axis -eq 'left'){'width'}else{'height'}
            $minimum=($selected|ForEach-Object{$_[$axis]}|Measure-Object -Minimum).Minimum
            $maximum=($selected|ForEach-Object{$_[$axis]+$_[$size]}|Measure-Object -Maximum).Maximum
            $positions=@{}
            if($Operation -eq 'alignShapes'){
                $factor=if($Params.alignment -in @('left','top')){0}elseif($Params.alignment -in @('center','middle')){0.5}else{1}
                $target=$minimum+($maximum-$minimum)*$factor
                foreach($s in $selected){$positions[[int]$s.id]=$target-$s[$size]*$factor}
            }else{
                $selected=@($selected|Sort-Object @{Expression={$_[$axis]}},@{Expression={$_.id}})
                $span=$selected[-1][$axis]+$selected[-1][$size]-$selected[0][$axis]
                $total=($selected|ForEach-Object{$_[$size]}|Measure-Object -Sum).Sum
                $gap=($span-$total)/($selected.Count-1)
                if($gap -lt 0){throw [PptActionException]::new('INVALID_PARAMS','Selected outer bounds cannot contain nonnegative gaps.')}
                $position=$selected[0][$axis]
                foreach($s in $selected){$positions[[int]$s.id]=$position;$position+=$s[$size]+$gap}
            }
            $script:ActionMayHaveEffect=$true
            foreach($s in $selected){$shape=Get-PptShape $slide $s.id;$shape.$axis=[single]$positions[[int]$s.id];Release-PptReference $shape}
            return Get-PptSlideSnapshot $slide
        }
        $shape=Get-PptShape $slide $Params.shapeId
        switch($Operation){
            'renameShape' {$script:ActionMayHaveEffect=$true;$shape.Name=[string]$Params.name}
            'setShapeOrder' {$script:ActionMayHaveEffect=$true;$shape.ZOrder($(if($Params.position -eq 'front'){0}else{1}))|Out-Null}
            'replaceText' {
                if([int]$shape.HasTextFrame -ne -1){throw [PptActionException]::new('TEXT_UNSUPPORTED','The shape has no top-level text frame.')}
                $frame=$shape.TextFrame;$range=$frame.TextRange;$raw=[string]$range.Text;$text=ConvertTo-PptText $raw
                if($raw.Length -ne $text.Length){throw [PptActionException]::new('TEXT_UNSUPPORTED','Native line breaks cannot be mapped to normalized UTF-16 offsets.')}
                $positions=Get-PptLiteralMatches $text $Params.find;$expected=$text.Replace([string]$Params.find,[string]$Params.replacement)
                if($expected.Length -gt 10000){throw [PptActionException]::new('CONTENT_LIMIT_EXCEEDED','Replacement text would exceed 10000 UTF-16 units.')}
                for($i=$positions.Count-1;$i -ge 0;$i--){
                    $segment=$range.Characters($positions[$i]+1,([string]$Params.find).Length)
                    $script:ActionMayHaveEffect=$true;$segment.Text=([string]$Params.replacement).Replace("`n","`r");Release-PptReference $segment
                }
                $after=Get-PptSlideSnapshot $slide;$actual=@($after.shapes|Where-Object{$_.id -eq $Params.shapeId})
                if($actual[0].text -cne $expected){throw [PptActionException]::new('PPT_VERIFICATION_FAILED','Replacement result differs from complete expected text.')}
                $after.replacements=$positions.Count;return $after
            }
            default {throw [PptActionException]::new('PPT_CAPABILITY_UNAVAILABLE','Unknown common PPT operation.')}
        }
        return Get-PptSlideSnapshot $slide
    }finally{
        if($image){$image.Dispose()};if($stream){$stream.Dispose()}
        foreach($ref in @($segment,$background,$transition,$body,$bullet,$format,$range,$frame,$color,$line,$fill,$shapes,$shape,$slide)){Release-PptReference $ref}
    }
}
