param([Parameter(Mandatory=$true)][string]$Path)
# Called only after the saved artifact and successful Task/cleanup were verified.
$ErrorActionPreference = 'Stop'
$app = $null; $books = $null
try {
    try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('KET.Application') }
    catch [Runtime.InteropServices.COMException] {
        if ($_.Exception.HResult -eq -2147221021) { exit 0 }
        throw
    }
    $books = $app.Workbooks
    for ($i = $books.Count; $i -ge 1; $i--) {
        $book = $books.Item($i)
        try {
            if ([string]::Equals([IO.Path]::GetFullPath([string]$book.FullName), [IO.Path]::GetFullPath($Path), [StringComparison]::OrdinalIgnoreCase)) {
                if (-not $book.Saved) { throw 'Test workbook changed after saving; preserve it for inspection.' }
                $book.Close($false)
            }
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    }
} finally {
    if ($null -ne $books) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($books) }
    if ($null -ne $app) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
}
