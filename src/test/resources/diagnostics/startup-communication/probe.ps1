# Test fixture only: same production Python transport, no WPS/COM calls.
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Console]::InputEncoding = New-Object Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
while ($null -ne ($line = [Console]::In.ReadLine())) {
    $request = ConvertFrom-Json -InputObject $line -ErrorAction Stop
    if ($request.operation -notin @('__diagnostic_ready', '__diagnostic_echo')) { throw 'Unexpected probe operation' }
    $response = [ordered]@{requestId=$request.requestId; outcome='succeeded'; data=$request.arguments}
    [Console]::Out.WriteLine((ConvertTo-Json -InputObject $response -Depth 20 -Compress))
    [Console]::Out.Flush()
}
