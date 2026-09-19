# Run in an interactive Windows session. Uses only a new, unsaved private workbook.
param([Parameter(Mandatory=$true)][string]$RuntimeRoot, [string]$LegacyActions, [Parameter(Mandatory=$true)][string]$ReportPath)
$ErrorActionPreference='Stop'
class ExcelActionException : System.Exception {
    [string]$Code
    ExcelActionException([string]$code,[string]$message):base($message){$this.Code=$code}
}
function Get-CoordinationHash { param([string]$Identity)
    $sha=[Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Identity))) } finally {$sha.Dispose()}
}
. "$RuntimeRoot/excel/windows/excel_common_actions.ps1"
if ($LegacyActions) {
    . $LegacyActions
    Set-Item Function:Get-LegacyExcelSnapshot ${function:Get-ExcelSnapshot}
}
. "$RuntimeRoot/excel/windows/excel_actions.ps1"
$script:Application=[Runtime.InteropServices.Marshal]::GetActiveObject('KET.Application')
$script:Document=$script:Application.Workbooks.Add()
$script:TokenSalt='regression'
$report=New-Object Collections.ArrayList
function Check([bool]$Value,[string]$Name) { if(-not $Value){throw "FAILED: $Name"}; [void]$report.Add($Name) }
function Reject([scriptblock]$Operation,[string]$Code,[string]$Name) {
    try { & $Operation; throw "Not rejected: $Name" }
    catch [ExcelActionException] { Check ($_.Exception.Code -eq $Code) $Name }
}
try {
    $sheet=Get-ExcelWorksheet -Name ([string]$script:Document.Worksheets.Item(1).Name)
    $name=[string]$sheet.Name
    $sheet.Range('A1').Value2='mixed'
    $sheet.Range('B1').Formula='=1/0'
    $sheet.Range('C1').Value2=-2146826281
    $sheet.Range('D1').Value2=$true
    $sheet.Range('E1').Value2="'=literal"
    $sheet.Range('A2').Value2=123.456
    $sheet.Range('A2').NumberFormat='0.00'
    $sheet.Range('B2').Formula='=1+2'
    $sheet.Range('B2').Font.Bold=$true
    $sheet.Range('C2').Font.Italic=$true
    $sheet.Range('D2').Font.Size=18
    $sheet.Range('E2').Font.Color=255
    $sheet.Range('F2').Interior.Color=65535
    $sheet.Range('G2').WrapText=$true
    $sheet.Range('H2').HorizontalAlignment=-4108
    $sheet.Range('I2').VerticalAlignment=-4108
    $sheet.Range('A4:C6').Merge()
    $sheet.Range('A4').Value2='merged'
    $sheet.Range('A10').EntireRow.Hidden=$true
    $sheet.Range('J1').EntireColumn.Hidden=$true
    $sheet.Range('A11').EntireRow.RowHeight=31
    $sheet.Range('K1').EntireColumn.ColumnWidth=22
    foreach($address in @('A1','B1','A1:P12','A4:C6','B4:C5')) {
        $range=$sheet.Range($address)
        try {
            $current=Get-ExcelSnapshot $sheet $range $address
            if($LegacyActions) {
                $legacy=Get-LegacyExcelSnapshot $sheet $range $address
                if (($current|ConvertTo-Json -Depth 20 -Compress) -cne ($legacy|ConvertTo-Json -Depth 20 -Compress)) {
                    [IO.File]::WriteAllText(($ReportPath+'.difference.json'),(@{current=$current;legacy=$legacy}|ConvertTo-Json -Depth 20))
                }
                Check (($current|ConvertTo-Json -Depth 20 -Compress) -ceq ($legacy|ConvertTo-Json -Depth 20 -Compress)) "snapshot equality $address"
            }
        } finally {Release-ExcelReference $range}
    }
    $p=[pscustomobject]@{sheet=$name;address='A4:D7'}
    $before=Invoke-ExcelRegionAction read_cell_rectangle $p
    $patch=[pscustomobject]@{bold=$true;italic=$true;fontSize=14;fontColor=255;fillColor=65535;wrapText=$true;horizontalAlignment='center';verticalAlignment='center';numberFormat='0.00'}
    $edit=[pscustomobject]@{sheet=$name;address=$p.address;expectedToken=$before.token;format=$patch}
    $after=Invoke-ExcelRegionAction format_rectangle $edit
    foreach($row in $after.cells){foreach($cell in $row){foreach($prop in $patch.PSObject.Properties){Check ($cell[$prop.Name] -eq $prop.Value) "merged format $($prop.Name)"}}}
    Check ($after.cells[0][0].value -ceq 'merged') 'merged content preserved'
    Reject {Invoke-ExcelRegionAction format_rectangle $edit} 'STALE_RANGE' 'stale token rejected'
    foreach($address in @('A4:B5','B4:C5')) {
        $range=$sheet.Range($address)
        try { Reject {Assert-ExcelEditableRectangle $sheet $range -AllowCompleteMergedAreas} 'RANGE_UNSUPPORTED' "partial merge rejected $address" }
        finally {Release-ExcelReference $range}
    }
    $range=$sheet.Range('A4:C6')
    try {Reject {Assert-ExcelEditableRectangle $sheet $range} 'RANGE_UNSUPPORTED' 'merged writes remain rejected'} finally {Release-ExcelReference $range}
    $sheet.Protect()
    $range=$sheet.Range('M1')
    try {Reject {Assert-ExcelEditableRectangle $sheet $range -AllowCompleteMergedAreas} 'RANGE_UNSUPPORTED' 'protected format rejected'} finally {Release-ExcelReference $range;$sheet.Unprotect()}
    $range=$sheet.Range('M1:M2');$range.FormulaArray='=1+1'
    try {Reject {Assert-ExcelEditableRectangle $sheet $range -AllowCompleteMergedAreas} 'RANGE_UNSUPPORTED' 'array format rejected'} finally {Release-ExcelReference $range}
    [IO.File]::WriteAllText($ReportPath,(@{passed=$true;checks=$report.Count;details=$report}|ConvertTo-Json -Depth 5))
} catch {
    [IO.File]::WriteAllText($ReportPath,(@{passed=$false;checks=$report.Count;error=($_|Out-String)}|ConvertTo-Json -Depth 5));throw
} finally {
    Release-ApplicationResources
    $script:Document.Close($false)
    Release-ExcelReference $script:Document
    Release-ExcelReference $script:Application
}
