# Persistence keeps the exact live document; destination claims join its lease.
class PersistenceActionException : System.Exception {
    [string]$Code
    PersistenceActionException([string]$code, [string]$message) : base($message) { $this.Code = $code }
}
if ($null -eq ('WpsSkills.OutputDirectory' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
namespace WpsSkills {
    public static class OutputDirectory {
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        public static extern SafeFileHandle CreateFileW(string path, uint access, uint share,
            IntPtr security, uint disposition, uint attributes, IntPtr template);
    }
}
'@
}

function New-OutputReservation {
    param([string]$Path, [string]$Extension, [switch]$ReserveFile)
    if (-not [IO.Path]::IsPathRooted($Path) -or -not $Path.EndsWith($Extension,[StringComparison]::OrdinalIgnoreCase)) {
        throw [PersistenceActionException]::new('OUTPUT_PATH_INVALID','Expected an absolute output path of the requested format.')
    }
    $full=[IO.Path]::GetFullPath($Path)
    $parent=[IO.Path]::GetDirectoryName($full)
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        throw [PersistenceActionException]::new('OUTPUT_PARENT_NOT_FOUND','The output parent directory does not exist.')
    }
    if (Test-Path -LiteralPath $full) {
        if ($ReserveFile -and (Test-Path -LiteralPath $full -PathType Leaf) -and -not [string]::IsNullOrEmpty($script:AuthorizedPath) -and
            (Get-StableFileIdentity $full) -eq $script:BoundFileIdentity) {
            throw [PersistenceActionException]::new('OUTPUT_MATCHES_BOUND_DOCUMENT','Use save for the existing backing file.')
        }
        throw [PersistenceActionException]::new('OUTPUT_ALREADY_EXISTS','The output already exists; this Action never overwrites an existing file.')
    }
    # Directory handle resolves junction aliases and prevents renaming this parent
    # during persistence. Its normalized locator matches existing-file acquisition.
    $handle=[WpsSkills.OutputDirectory]::CreateFileW($parent,0,3,[IntPtr]::Zero,3,0x02000000,[IntPtr]::Zero)
    if ($handle.IsInvalid) { $handle.Dispose();throw [PersistenceActionException]::new('OUTPUT_ACCESS_DENIED','Cannot reserve the output directory.') }
    $identity=$null
    try {
        $buffer=New-Object Text.StringBuilder 32768
        $length=[WpsSkills.NativeFileIdentity]::GetFinalPathNameByHandleW($handle,$buffer,32768,0)
        if ($length -eq 0 -or $length -ge 32768) { throw 'Cannot normalize output directory.' }
        $locator=($buffer.ToString().TrimEnd('\')+'\'+[IO.Path]::GetFileName($full)).ToUpperInvariant()
        Add-CoordinationFence -Identity ('locator-'+$locator)
        if (Test-Path -LiteralPath $full) { throw [PersistenceActionException]::new('OUTPUT_ALREADY_EXISTS','The output appeared during reservation.') }
        $identity=$null
        if ($ReserveFile) {
            # CreateNew atomically claims an absent destination. Native SaveAs may
            # replace this task-owned empty file; it never receives a user file.
            $stream=[IO.File]::Open($full,[IO.FileMode]::CreateNew,[IO.FileAccess]::ReadWrite,[IO.FileShare]::Read)
            $stream.Dispose()
            $identity=Get-StableFileIdentity $full
            Add-CoordinationFence -Identity $identity
        }
        return @{path=$full;locator=$locator;handle=$handle;identity=$identity}
    }
    catch {
        if ($identity -and [IO.File]::Exists($full) -and (Get-StableFileIdentity $full) -eq $identity) { [IO.File]::Delete($full) }
        $handle.Dispose();throw
    }
}

function Complete-SaveAsBinding {
    param($Reservation,[string]$AuthorizedPath)
    $actual=[IO.Path]::GetFullPath([string]$script:Document.FullName)
    if ((Get-NormalizedFileLocator $actual) -cne $Reservation.locator -or -not [bool]$script:Document.Saved) {
        throw [PersistenceActionException]::new('OUTPUT_VERIFICATION_FAILED','The exact document did not become the saved output.')
    }
    $identity=Get-StableFileIdentity $actual
    Add-CoordinationFence -Identity $identity
    $file=Get-Item -LiteralPath $actual
    if ($file.Length -lt 1) { throw [PersistenceActionException]::new('OUTPUT_VERIFICATION_FAILED','Saved output is empty.') }
    $script:AuthorizedPath=$AuthorizedPath
    $script:PreparedCanonicalPath=$actual
    $script:PreparedLocator=$Reservation.locator
    $script:BoundFileIdentity=$identity
    return [long]$file.Length
}

function Assert-SaveAsNames {
    param([string]$Path)
    for($i=1;$i -le [int]$script:Documents.Count;$i++) {
        $candidate=$script:Documents.Item($i)
        if (-not [object]::ReferenceEquals($candidate,$script:Document) -and
            [string]::Equals([string]$candidate.Name,[IO.Path]::GetFileName($Path),[StringComparison]::OrdinalIgnoreCase)) {
            throw [PersistenceActionException]::new('OUTPUT_IN_USE','Another open document has this output file name.')
        }
        # RCWs may be shared with the retained Document; release at bridge exit.
    }
}

function Test-OutputSignature {
    param([string]$Path,[string]$Format)
    if (-not [IO.File]::Exists($Path)) { return $false }
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    try {
        $bytes=New-Object byte[] 8
        $n=$stream.Read($bytes,0,8)
        if ($Format -eq 'pdf') { return $n -ge 5 -and [Text.Encoding]::ASCII.GetString($bytes,0,5) -eq '%PDF-' }
        if ($Format -eq 'png') { return $n -eq 8 -and [BitConverter]::ToString($bytes) -eq '89-50-4E-47-0D-0A-1A-0A' }
        return $false
    }
    finally { $stream.Dispose() }
}
