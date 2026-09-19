param([Parameter(Mandatory=$true)][string]$ProgId)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
[Console]::InputEncoding=New-Object Text.UTF8Encoding($false)
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
$identity=[Security.Principal.WindowsIdentity]::GetCurrent()
try {
    $principal=[Security.Principal.WindowsPrincipal]::new($identity)
    $registered=[type]::GetTypeFromProgID($ProgId)
    $request=ConvertFrom-Json -InputObject ([Console]::In.ReadLine())
    [ordered]@{
        powershellVersion=$PSVersionTable.PSVersion.ToString()
        edition=[string]$PSVersionTable.PSEdition
        processBits=([IntPtr]::Size*8)
        elevated=$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        sessionId=[Diagnostics.Process]::GetCurrentProcess().SessionId
        userInteractive=[Environment]::UserInteractive
        apartmentState=[string][Threading.Thread]::CurrentThread.GetApartmentState()
        progId=$ProgId
        registered=($null -ne $registered)
        echo=$request.payload
    } | ConvertTo-Json -Depth 5 -Compress
} finally { $identity.Dispose() }
