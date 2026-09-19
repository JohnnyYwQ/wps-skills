param([ValidateSet('word','excel','ppt')][string]$Application,[Parameter(Mandatory=$true)][string]$Path)
# Only invoked after response/receipt/artifact assertions succeed. No activation.
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
$progId=@{word='KWPS.Application';excel='KET.Application';ppt='KWPP.Application'}[$Application]
$app=$null;$documents=$null;$found=$false
try {
    try {$app=[Runtime.InteropServices.Marshal]::GetActiveObject($progId)}
    catch [Runtime.InteropServices.COMException] {
        if ($_.Exception.HResult -eq -2147221021) { 'not_attached_preserved';exit 0 }
        throw
    }
    if($Application -eq 'word'){$documents=$app.Documents}elseif($Application -eq 'excel'){$documents=$app.Workbooks}else{$documents=$app.Presentations}
    for($i=$documents.Count;$i -ge 1;$i--){
        $document=$documents.Item($i)
        try {
            if ([string]::Equals([IO.Path]::GetFullPath([string]$document.FullName),[IO.Path]::GetFullPath($Path),[StringComparison]::OrdinalIgnoreCase)) {
                if (-not $document.Saved){throw 'Saved test document changed after verification; preserve it.'}
                if($Application -eq 'ppt'){$document.Close()}else{$document.Close(0)}
                $found=$true
            }
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
    }
    if($found){'closed_exact_verified_document'}else{'not_found_preserved'}
} finally {
    if($null -ne $documents -and [Runtime.InteropServices.Marshal]::IsComObject($documents)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($documents)}
    if($null -ne $app -and [Runtime.InteropServices.Marshal]::IsComObject($app)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}
}
