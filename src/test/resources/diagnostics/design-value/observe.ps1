param([ValidateSet('word','excel','ppt')][string]$Application,[string]$Path,[string]$OwnedRoot,[ValidateSet('read','dirty','close')][string]$Mode='read')
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
$canonical=[IO.Path]::GetFullPath($Path)
$owned=[IO.Path]::GetFullPath($OwnedRoot).TrimEnd('\')+'\'
if(-not $canonical.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Observation outside test root'}
$app=$null;$docs=$null;$doc=$null
function Release($value){if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)}}
try {
    $app=[Runtime.InteropServices.Marshal]::GetActiveObject(@{word='KWPS.Application';excel='KET.Application';ppt='KWPP.Application'}[$Application])
    if($Application -eq 'word'){$docs=$app.Documents}elseif($Application -eq 'excel'){$docs=$app.Workbooks}else{$docs=$app.Presentations}
    for($i=1;$i -le $docs.Count;$i++){
        $item=$docs.Item($i)
        if([string]::Equals([string]$item.FullName,$canonical,[StringComparison]::OrdinalIgnoreCase)){$doc=$item;break}
        Release $item
    }
    if($null -eq $doc){$doc=$docs.Open($canonical)}
    if(-not [string]::Equals([string]$doc.FullName,$canonical,[StringComparison]::OrdinalIgnoreCase)){throw 'Wrong observed document'}
    $texts=New-Object Collections.ArrayList
    if($Application -eq 'word'){
        $content=$doc.Content
        if($Mode -eq 'dirty'){$content.InsertAfter('EXTERNAL_DIRTY')}
        [void]$texts.Add([string]$content.Text);Release $content
    } elseif($Application -eq 'excel'){
        $sheets=$doc.Worksheets;$sheet=$sheets.Item('实验')
        if($Mode -eq 'dirty'){$cell=$sheet.Range('D1');$cell.Value2='EXTERNAL_DIRTY';Release $cell}
        for($i=1;$i -le 8;$i++){$cell=$sheet.Range('A'+$i);[void]$texts.Add($cell.Value2);Release $cell}
        Release $sheet;Release $sheets
    } else {
        $slides=$doc.Slides
        for($i=1;$i -le $slides.Count;$i++){
            $slide=$slides.Item($i);$shapes=$slide.Shapes
            for($j=1;$j -le $shapes.Count;$j++){
                $shape=$shapes.Item($j)
                if($shape.HasTextFrame -and $shape.TextFrame.HasText){
                    $frame=$shape.TextFrame;$range=$frame.TextRange
                    if($Mode -eq 'dirty' -and $i -eq 1 -and $j -eq 1){$range.Text=[string]$range.Text+'EXTERNAL_DIRTY'}
                    [void]$texts.Add([string]$range.Text);Release $range;Release $frame
                };Release $shape
            };Release $shapes;Release $slide
        };Release $slides
    }
    $saved=[bool]$doc.Saved
    if($Mode -eq 'close'){
        if(-not $saved){throw 'Preserve unsaved test document'}
        if($Application -eq 'ppt'){$doc.Close()}else{$doc.Close(0)}
    }
    @{path=$canonical;saved=$saved;texts=@($texts.ToArray());mode=$Mode}|ConvertTo-Json -Depth 8 -Compress
} finally {Release $doc;Release $docs;Release $app}
