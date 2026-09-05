param([Parameter(Mandatory=$true)][string]$DocumentPath)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$application = [Runtime.InteropServices.Marshal]::GetActiveObject('KWPS.Application')
$documents = $application.Documents
$matches = @()
for ($i=1; $i -le $documents.Count; $i++) {
    $document = $documents.Item($i)
    if ([StringComparer]::OrdinalIgnoreCase.Equals([string]$document.FullName, $DocumentPath)) {
        $matches += $document
    }
}
if ($matches.Count -ne 1) { throw 'Expected one exact open demo document.' }
$window = $matches[0].ActiveWindow
$handle = [long]$window.Hwnd
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class WordDemoWindowProbe {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
'@
[uint32]$owner = 0
[void][WordDemoWindowProbe]::GetWindowThreadProcessId([IntPtr]$handle, [ref]$owner)
@{hwnd=$handle; processId=$owner} | ConvertTo-Json -Compress
