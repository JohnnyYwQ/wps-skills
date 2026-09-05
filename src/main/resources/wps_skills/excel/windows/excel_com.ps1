# Read-only COM availability check. Do not bypass COM's elevation boundary.
# https://devblogs.microsoft.com/oldnewthing/20190801-00/?p=102745
function Get-ExcelComAvailability {
    try { $type = [Type]::GetTypeFromProgID('KET.Application') } catch { $type = $null }
    if ($null -ne $type) { return @{available=$true; reason=''; message=''} }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    $elevated = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $userKey=$null; $machineKey=$null
    try {
        $userKey=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\Classes\KET.Application\CLSID')
        $machineKey=[Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('Software\Classes\KET.Application\CLSID')
        $userRegistered=$null -ne $userKey -and -not [string]::IsNullOrEmpty([string]$userKey.GetValue(''))
        $machineRegistered=$null -ne $machineKey -and -not [string]::IsNullOrEmpty([string]$machineKey.GetValue(''))
        if ($elevated -and $userRegistered -and -not $machineRegistered) {
            return @{available=$false; reason='elevated_user_registration'; message='WPS Spreadsheets is registered only for the current user. This elevated process cannot resolve KET.Application. Run the same command from a non-administrator PowerShell window.'}
        }
    } finally {
        if ($userKey) { $userKey.Dispose() }
        if ($machineKey) { $machineKey.Dispose() }
        $identity.Dispose()
    }
    return @{available=$false; reason='registration_unavailable'; message='KET.Application is not available to this process. Check that WPS Spreadsheets is installed and registered for the current user and process architecture.'}
}
