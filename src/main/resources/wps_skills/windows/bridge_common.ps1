if ($null -eq ('WpsSkills.NativeFileIdentity' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace WpsSkills {
    [StructLayout(LayoutKind.Sequential)]
    public struct ByHandleFileInformation {
        public uint FileAttributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME CreationTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastAccessTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastWriteTime;
        public uint VolumeSerialNumber;
        public uint FileSizeHigh;
        public uint FileSizeLow;
        public uint NumberOfLinks;
        public uint FileIndexHigh;
        public uint FileIndexLow;
    }

    public static class NativeFileIdentity {
        // https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-getfinalpathnamebyhandlew
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern uint GetFinalPathNameByHandleW(
            SafeFileHandle handle, System.Text.StringBuilder path, uint size, uint flags
        );
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool GetFileInformationByHandle(
            SafeFileHandle handle,
            out ByHandleFileInformation information
        );
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WindowRect {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MonitorInfo {
        public int Size;
        public WindowRect Monitor;
        public WindowRect Work;
        public uint Flags;
    }

    public static class NativeWindowPlacement {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
        [DllImport("user32.dll")]
        public static extern bool IsIconic(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(IntPtr hwnd, int command);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, out WindowRect rect);

        [DllImport("user32.dll")]
        public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);

        [DllImport("user32.dll")]
        public static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo information);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetWindowPos(
            IntPtr hwnd,
            IntPtr insertAfter,
            int x,
            int y,
            int width,
            int height,
            uint flags
        );
    }
}
'@
}

function Write-BridgeRecord {
    param([Parameter(Mandatory = $true)]$Record)

    $json = $Record | ConvertTo-Json -Compress -Depth 24
    $ascii = New-Object Text.StringBuilder
    foreach ($character in $json.ToCharArray()) {
        $codePoint = [int][char]$character
        if ($codePoint -le 0x7f) {
            [void]$ascii.Append($character)
        }
        else {
            [void]$ascii.AppendFormat('\u{0:x4}', $codePoint)
        }
    }
    [Console]::Out.WriteLine($ascii.ToString())
    [Console]::Out.Flush()
}

function New-SuccessRecord {
    param(
        [Parameter(Mandatory = $true)][string]$RequestId,
        [Parameter(Mandatory = $true)]$Data
    )

    return [ordered]@{
        requestId = $RequestId
        outcome = 'succeeded'
        data = $Data
    }
}

function New-FailureRecord {
    param(
        [Parameter(Mandatory = $true)][string]$RequestId,
        [Parameter(Mandatory = $true)][ValidateSet('failed', 'unknown')]
        [string]$Outcome,
        [Parameter(Mandatory = $true)][string]$Code,
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $true)]
        [ValidateSet('unchanged', 'lost', 'unprovable')]
        [string]$BindingDisposition
    )

    return [ordered]@{
        requestId = $RequestId
        outcome = $Outcome
        error = [ordered]@{
            code = $Code
            message = $Message
        }
        bindingDisposition = $BindingDisposition
    }
}

function Get-ObjectPropertyNames {
    param([Parameter(Mandatory = $true)]$Value)

    return @($Value.PSObject.Properties | ForEach-Object { $_.Name })
}

function Test-ExactFields {
    param(
        [Parameter(Mandatory = $true)]$Value,
        [Parameter(Mandatory = $true)][string[]]$Expected
    )

    if ($null -eq $Value -or $Value -isnot [pscustomobject]) {
        return $false
    }
    $actual = @(Get-ObjectPropertyNames -Value $Value | Sort-Object)
    $wanted = @($Expected | Sort-Object)
    return $null -eq (Compare-Object -ReferenceObject $wanted -DifferenceObject $actual)
}

function Get-StableFileIdentity {
    param([Parameter(Mandatory = $true)][string]$Path)

    $stream = $null
    try {
        $stream = [IO.File]::Open(
            $Path,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Read,
            ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
        )
        $information = New-Object WpsSkills.ByHandleFileInformation
        if (-not [WpsSkills.NativeFileIdentity]::GetFileInformationByHandle(
            $stream.SafeFileHandle,
            [ref]$information
        )) {
            throw [ComponentModel.Win32Exception]::new(
                [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            )
        }
        $fileIndex = ([uint64]$information.FileIndexHigh -shl 32) -bor [uint64]$information.FileIndexLow
        return 'file-{0:x8}-{1:x16}' -f [uint32]$information.VolumeSerialNumber, $fileIndex
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

# Supplemental fences share the same mutex/state protocol as the primary file
# identity. They remain held through every in-place file replacement and cleanup.
$script:AdditionalFences = New-Object Collections.ArrayList
$script:WpsCallInFlight = $false

function Get-NormalizedFileLocator {
    param([string]$Path)
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read,
        ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    try {
        $buffer = New-Object Text.StringBuilder 32768
        $length = [WpsSkills.NativeFileIdentity]::GetFinalPathNameByHandleW($stream.SafeFileHandle, $buffer, 32768, 0)
        if ($length -eq 0 -or $length -ge 32768) {
            throw [ComponentModel.Win32Exception]::new([Runtime.InteropServices.Marshal]::GetLastWin32Error())
        }
        return $buffer.ToString().ToUpperInvariant()
    }
    finally { $stream.Dispose() }
}

function Get-CoordinationContext {
    return @{
        PreparedIdentity = $script:PreparedIdentity
        CoordinationMutex = $script:CoordinationMutex
        CoordinationMutexName = $script:CoordinationMutexName
        CoordinationStatePath = $script:CoordinationStatePath
        CoordinationGuardId = $script:CoordinationGuardId
        CoordinationLeaseId = $script:CoordinationLeaseId
    }
}

function Set-CoordinationContext {
    param($Context)
    $script:PreparedIdentity = $Context.PreparedIdentity
    $script:CoordinationMutex = $Context.CoordinationMutex
    $script:CoordinationMutexName = $Context.CoordinationMutexName
    $script:CoordinationStatePath = $Context.CoordinationStatePath
    $script:CoordinationGuardId = $Context.CoordinationGuardId
    $script:CoordinationLeaseId = $Context.CoordinationLeaseId
}

function Add-CoordinationFence {
    param([string]$Identity)
    if ($Identity -eq $script:PreparedIdentity -and $null -ne $script:CoordinationMutex) { return }
    foreach ($fence in $script:AdditionalFences) {
        if ($fence.PreparedIdentity -eq $Identity) { return }
    }
    $primary = Get-CoordinationContext
    try {
        Set-CoordinationContext -Context @{ PreparedIdentity = $Identity }
        Invoke-AcquireCoordinationGuard -Arguments ([pscustomobject]@{ coordinationIdentity = $Identity }) | Out-Null
        Write-CoordinationState -Mode 'guard' -InFlight $script:WpsCallInFlight
    }
    finally {
        # Retain even a partially initialized acquired fence for terminal cleanup.
        if ($null -ne $script:CoordinationMutex) { [void]$script:AdditionalFences.Add((Get-CoordinationContext)) }
        Set-CoordinationContext -Context $primary
    }
}

function Release-AdditionalFences {
    param([bool]$Clean)
    $primary = Get-CoordinationContext
    try {
        foreach ($fence in $script:AdditionalFences) {
            Set-CoordinationContext -Context $fence
            Release-PrimaryCoordinationResources -Clean $Clean
        }
        $script:AdditionalFences.Clear()
    }
    finally { Set-CoordinationContext -Context $primary }
}

function Get-CoordinationHash {
    param([Parameter(Mandatory = $true)][string]$Identity)

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Identity)
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Write-CoordinationState {
    param(
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][bool]$InFlight
    )

    if ($null -eq $script:CoordinationMutex) { return }
    $directory = [IO.Path]::GetDirectoryName($script:CoordinationStatePath)
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    $record = [ordered]@{
        identity = $script:PreparedIdentity
        ownerPid = $PID
        mode = $Mode
        inFlight = $InFlight
        updatedUtc = [DateTime]::UtcNow.ToString('o')
    }
    $temporary = $script:CoordinationStatePath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    try {
        [IO.File]::WriteAllText(
            $temporary,
            ($record | ConvertTo-Json -Compress),
            (New-Object Text.UTF8Encoding($false))
        )
        Move-Item -LiteralPath $temporary -Destination $script:CoordinationStatePath -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporary) {
            Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
        }
    }
}

function Set-CoordinationInFlight {
    param([Parameter(Mandatory = $true)][bool]$InFlight)
    $script:WpsCallInFlight = $InFlight
    $primary = Get-CoordinationContext
    try {
        if ($null -ne $script:CoordinationMutex) {
            $mode = if ([string]::IsNullOrEmpty($script:CoordinationLeaseId)) { 'guard' } else { 'lease' }
            Write-CoordinationState -Mode $mode -InFlight $InFlight
        }
        foreach ($fence in $script:AdditionalFences) {
            Set-CoordinationContext -Context $fence
            Write-CoordinationState -Mode 'guard' -InFlight $InFlight
        }
    }
    finally { Set-CoordinationContext -Context $primary }
}

function Invoke-CoordinatedWpsCall {
    param([Parameter(Mandatory = $true)][scriptblock]$Action)

    if ($null -eq $script:CoordinationMutex) {
        throw 'A WPS call requires an acquired document guard.'
    }
    Set-CoordinationInFlight -InFlight $true
    try {
        return & $Action
    }
    finally {
        Set-CoordinationInFlight -InFlight $false
    }
}

function Release-PrimaryCoordinationResources {
    param([Parameter(Mandatory = $true)][bool]$Clean)

    if ($null -eq $script:CoordinationMutex) { return }
    try {
        if ($Clean -and -not [string]::IsNullOrEmpty($script:CoordinationStatePath)) {
            Remove-Item -LiteralPath $script:CoordinationStatePath -Force -ErrorAction SilentlyContinue
        }
        $script:CoordinationMutex.ReleaseMutex()
    }
    catch {
        if ($Clean) { throw }
    }
    finally {
        $script:CoordinationMutex.Dispose()
        $script:CoordinationMutex = $null
        $script:CoordinationMutexName = $null
        $script:CoordinationStatePath = $null
        $script:CoordinationGuardId = $null
        $script:CoordinationLeaseId = $null
    }
}

function Release-ComReference {
    param($Value)

    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
        }
        catch {
        }
    }
}

function Invoke-AcquireCoordinationGuard {
    param([Parameter(Mandatory = $true)]$Arguments)

    if ($null -ne $script:CoordinationMutex) {
        throw 'The bridge already owns a coordination guard.'
    }
    $identity = [string]$Arguments.coordinationIdentity
    if (
        [string]::IsNullOrEmpty($script:PreparationId) -or
        $identity -ne $script:PreparedIdentity
    ) {
        throw 'The coordination identity does not match the prepared document.'
    }
    $hash = Get-CoordinationHash -Identity $identity
    $mutexName = 'Global\WpsSkills.Document.' + $hash
    $stateDirectory = Join-Path $env:LOCALAPPDATA 'WpsSkills\document-coordination'
    $statePath = Join-Path $stateDirectory ($hash + '.json')
    $mutex = New-Object Threading.Mutex($false, $mutexName)
    $acquired = $false
    $abandoned = $false
    try {
        try {
            $acquired = $mutex.WaitOne(0)
        }
        catch [Threading.AbandonedMutexException] {
            $acquired = $true
            $abandoned = $true
        }
        if (-not $acquired) {
            throw [DocumentLeaseConflictException]::new(
                'Another Task owns the document Lease.'
            )
        }
        if (Test-Path -LiteralPath $statePath) {
            $unsafe = $true
            try {
                $previous = [IO.File]::ReadAllText($statePath) | ConvertFrom-Json -ErrorAction Stop
                $unsafe = (
                    [string]$previous.identity -ne $identity -or
                    [bool]$previous.inFlight -or
                    [string]$previous.mode -eq 'quarantine'
                )
            }
            catch {
                $unsafe = $true
            }
            if ($unsafe) {
                try { $mutex.ReleaseMutex() } catch {}
                throw [DocumentQuarantinedException]::new(
                    'The prior document owner ended without proving WPS quiescence.'
                )
            }
            Remove-Item -LiteralPath $statePath -Force
        }
        elseif ($abandoned) {
            # No persisted in-flight marker means the old owner had not entered WPS.
        }
        $script:CoordinationMutex = $mutex
        $script:CoordinationMutexName = $mutexName
        $script:CoordinationStatePath = $statePath
        $script:CoordinationGuardId = 'guard-' + [guid]::NewGuid().ToString('N')
        Write-CoordinationState -Mode 'guard' -InFlight $false
        $mutex = $null
        return [ordered]@{ guardId = $script:CoordinationGuardId }
    }
    finally {
        if ($null -ne $mutex) { $mutex.Dispose() }
    }
}

function Invoke-CommitDocumentLease {
    param([Parameter(Mandatory = $true)]$Arguments)

    if (
        $null -eq $script:CoordinationMutex -or
        [string]$Arguments.guardId -ne $script:CoordinationGuardId -or
        [string]$Arguments.documentId -ne $script:DocumentId
    ) {
        throw 'The acquisition guard and exact document cannot be committed.'
    }
    $script:CoordinationLeaseId = 'lease-' + [guid]::NewGuid().ToString('N')
    Write-CoordinationState -Mode 'lease' -InFlight $false
    return [ordered]@{ leaseId = $script:CoordinationLeaseId }
}

function Invoke-ReleaseDocumentResources {
    param([Parameter(Mandatory = $true)]$Arguments)

    if ($null -eq $script:CoordinationMutex) {
        return [ordered]@{ state = 'not_acquired' }
    }
    if (
        -not [string]::IsNullOrEmpty($script:CoordinationLeaseId) -and
        [string]$Arguments.leaseId -ne $script:CoordinationLeaseId
    ) {
        throw 'The document Lease reference does not match.'
    }
    if (
        [string]::IsNullOrEmpty($script:CoordinationLeaseId) -and
        [string]$Arguments.guardId -ne $script:CoordinationGuardId
    ) {
        throw 'The acquisition guard reference does not match.'
    }
    if (Get-Command Invoke-DebugCreatedDocumentCleanup -ErrorAction SilentlyContinue) {
        Invoke-DebugCreatedDocumentCleanup
    }
    Release-CoordinationResources -Clean $true
    return [ordered]@{ state = 'released' }
}


function Release-CoordinationResources {
    param([Parameter(Mandatory = $true)][bool]$Clean)
    # Keep the locator fence until all file-identity claims have stopped being used.
    Release-PrimaryCoordinationResources -Clean $Clean
    Release-AdditionalFences -Clean $Clean
}
