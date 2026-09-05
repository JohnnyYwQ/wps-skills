[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$script:Application = $null
$script:Documents = $null
$script:Document = $null
$script:DocumentId = $null
$script:AuthorizedPath = $null
$script:PreparationId = $null
$script:PreparedIdentity = $null
$script:BoundFileIdentity = $null
$script:PreparedLocator = $null
$script:PreparedPath = $null
$script:PreparedCanonicalPath = $null
$script:CoordinationMutex = $null
$script:CoordinationMutexName = $null
$script:CoordinationStatePath = $null
$script:CoordinationGuardId = $null
$script:CoordinationLeaseId = $null
$script:TokenSalt = [guid]::NewGuid().ToString('N')
$script:ActionMayHaveEffect = $false

class DocumentLeaseConflictException : System.Exception {
    DocumentLeaseConflictException([string]$message) : base($message) {}
}
class DocumentQuarantinedException : System.Exception {
    DocumentQuarantinedException([string]$message) : base($message) {}
}
class PptActionException : System.Exception {
    [string]$Code
    PptActionException([string]$code, [string]$message) : base($message) { $this.Code = $code }
}

. (Join-Path $PSScriptRoot '../../windows/bridge_common.ps1')
. (Join-Path $PSScriptRoot '../../windows/window_presentation.ps1')
. (Join-Path $PSScriptRoot 'ppt_com.ps1')
. (Join-Path $PSScriptRoot 'ppt_actions.ps1')
. (Join-Path $PSScriptRoot '../../windows/output_persistence.ps1')
. (Join-Path $PSScriptRoot 'ppt_persistence.ps1')
. (Join-Path $PSScriptRoot 'ppt_common_actions.ps1')
. (Join-Path $PSScriptRoot 'ppt_window.ps1')

function Invoke-DebugCreatedDocumentCleanup {
    # This slice only opens existing user presentations. Cleanup never closes them.
}

function Invoke-PrepareExistingDocument {
    param($Arguments)
    if ($null -ne $script:Document -or $null -ne $script:CoordinationMutex) {
        throw 'The bridge is already bound or coordinated.'
    }
    $path = [string]$Arguments.path
    if (-not [IO.Path]::IsPathRooted($path) -or -not $path.EndsWith('.pptx', [StringComparison]::OrdinalIgnoreCase)) {
        throw [PptActionException]::new('INVALID_PARAMS', 'Expected an absolute .pptx path.')
    }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw [IO.FileNotFoundException]::new('The presentation does not exist.', $path)
    }
    $availability = Get-PptComAvailability
    if (-not $availability.available) {
        throw [PptActionException]::new('PPT_CAPABILITY_UNAVAILABLE', $availability.message)
    }
    Assert-PptxPackage -Path $path
    $canonical = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $path).Path)
    if ((Get-Item -LiteralPath $canonical).Length -lt 1) { throw 'The presentation is empty.' }
    $script:PreparationId = 'preparation-' + [guid]::NewGuid().ToString('N')
    $script:PreparedIdentity = Get-StableFileIdentity -Path $canonical
    $script:PreparedLocator = Get-NormalizedFileLocator -Path $canonical
    $script:PreparedPath = $path
    $script:PreparedCanonicalPath = $canonical
    return [ordered]@{ preparationId = $script:PreparationId; coordinationIdentity = $script:PreparedIdentity }
}

function Get-PresentationState {
    $state = if ([string]::IsNullOrEmpty([string]$script:Document.Path)) { 'unsaved' } elseif ([bool]$script:Document.Saved) { 'saved' } else { 'modified' }
    return [ordered]@{ persistenceState = $state; readOnly = [bool]$script:Document.ReadOnly }
}

function Test-BoundDocumentLive {
    if ($null -eq $script:Document) { return $false }
    $sheets = $null
    try {
        $name = [string]$script:Document.Name
        $sheets = $script:Document.Slides
        return -not [string]::IsNullOrEmpty($name) -and [int]$sheets.Count -ge 0
    }
    catch { return $false }
    finally { Release-PptReference -Value $sheets }
}

function Assert-BoundPresentation {
    param([string]$DocumentId)
    if ($DocumentId -cne $script:DocumentId -or -not (Test-BoundDocumentLive)) {
        throw [PptActionException]::new('DOCUMENT_CLOSED', 'The exact bound presentation is no longer available.')
    }
    if ([string]::IsNullOrEmpty($script:AuthorizedPath)) {
        if (-not [string]::IsNullOrEmpty([string]$script:Document.Path) -or [string]$script:Document.Name -cne $script:UnsavedName) {
            throw [PptActionException]::new('DOCUMENT_BINDING_UNAVAILABLE','The new document was saved outside this Session.')
        }
        return
    }
    # Follow only the locator committed by this Session’s explicit persistence Action.
    $actual = [IO.Path]::GetFullPath([string]$script:Document.FullName)
    if (-not [string]::Equals($actual, $script:PreparedCanonicalPath, [StringComparison]::OrdinalIgnoreCase) -or
        (Get-StableFileIdentity -Path $actual) -ne $script:BoundFileIdentity) {
        throw [PptActionException]::new('DOCUMENT_BINDING_UNAVAILABLE', 'The presentation backing identity changed.')
    }
}

function Invoke-AcquireExistingDocument {
    param($Arguments)
    if ($null -ne $script:Document -or $null -eq $script:CoordinationMutex -or
        [string]$Arguments.preparationId -cne $script:PreparationId) { throw 'Invalid prepared acquisition.' }
    if ((Get-StableFileIdentity -Path $script:PreparedCanonicalPath) -ne $script:PreparedIdentity) {
        throw 'Presentation identity changed after preparation.'
    }
    if ((Get-NormalizedFileLocator -Path $script:PreparedCanonicalPath) -cne $script:PreparedLocator) { throw 'Presentation locator changed after preparation.' }
    $applicationCreated = $false
    try { $script:Application = [Runtime.InteropServices.Marshal]::GetActiveObject('KWPP.Application') }
    catch {
        $script:Application = New-Object -ComObject 'KWPP.Application'
        $applicationCreated = $true
    }
    $script:Documents = $script:Application.Presentations
    $matches = New-Object Collections.ArrayList
    $nameConflict = $false
    $requestedName = [IO.Path]::GetFileName($script:PreparedCanonicalPath)
    for ($index = 1; $index -le [int]$script:Documents.Count; $index++) {
        $candidate = $script:Documents.Item($index)
        try {
            if ([string]::Equals([string]$candidate.Name, $requestedName, [StringComparison]::OrdinalIgnoreCase)) { $nameConflict = $true }
            if (-not [string]::IsNullOrEmpty([string]$candidate.Path)) {
                if ((Get-StableFileIdentity -Path ([string]$candidate.FullName)) -eq $script:PreparedIdentity) {
                    [void]$matches.Add($candidate)
                    $candidate = $null
                }
            }
        }
        finally { Release-PptReference -Value $candidate }
    }
    if ($matches.Count -gt 1) {
        foreach ($match in $matches) { Release-PptReference -Value $match }
        throw 'More than one live presentation claims the requested file.'
    }
    if ($matches.Count -eq 1) { $script:Document = $matches[0] }
    else {
        if ($nameConflict) {
            throw [PptActionException]::new('DOCUMENT_OPEN_FAILED', 'WPS already has another presentation with this file name open. Use a different file name or close that presentation explicitly before opening this file.')
        }
        # Open the exact existing presentation, writable, with its own window.
        $script:ActionMayHaveEffect = $true
        $script:Document = $script:Documents.Open($script:PreparedCanonicalPath, 0, 0, -1)
    }
    $actual = [IO.Path]::GetFullPath([string]$script:Document.FullName)
    if ((Get-StableFileIdentity -Path $actual) -ne $script:PreparedIdentity) { throw 'WPS returned another presentation.' }

    Show-PptPresentationWindow -ApplicationCreated $applicationCreated
    $script:DocumentId = 'presentation-' + [guid]::NewGuid().ToString('N')
    $script:AuthorizedPath = $script:PreparedPath
    $script:BoundFileIdentity = $script:PreparedIdentity
    return [ordered]@{
        documentId = $script:DocumentId
        artifact = [ordered]@{ path = $script:AuthorizedPath; format = 'pptx'; sizeBytes = [long](Get-Item -LiteralPath $actual).Length }
        documentState = Get-PresentationState
    }
}

function Invoke-BridgeOperation {
    param([string]$RequestId, [string]$Operation, $Arguments)
    $script:ActionMayHaveEffect = $false
    try {
        switch ($Operation) {
            'prepare_new_document' { $data = Invoke-PrepareNewDocument }
            'acquire_new_document' { $data = Invoke-CoordinatedWpsCall { Invoke-AcquireNewDocument -Arguments $Arguments } }
            'prepare_existing_document' { $data = Invoke-PrepareExistingDocument -Arguments $Arguments }
            'acquire_coordination_guard' {
                if ($script:PreparedLocator) { Add-CoordinationFence -Identity ('locator-' + $script:PreparedLocator) }
                $data = Invoke-AcquireCoordinationGuard -Arguments $Arguments
            }
            'commit_document_lease' { $data = Invoke-CommitDocumentLease -Arguments $Arguments }
            'release_document_resources' { $data = Invoke-ReleaseDocumentResources -Arguments $Arguments }
            'acquire_existing_document' {
                $data = Invoke-CoordinatedWpsCall { Invoke-AcquireExistingDocument -Arguments $Arguments }
            }
            'probe_bound_document' {
                $data = Invoke-CoordinatedWpsCall {
                    [ordered]@{ live = ([string]$Arguments.documentId -ceq $script:DocumentId -and (Test-BoundDocumentLive)) }
                }
            }
            { $_ -in @('saveAs','exportPdf','exportSlideImage') } {
                $data=Invoke-CoordinatedWpsCall {
                    Assert-BoundPresentation -DocumentId ([string]$Arguments.documentId)
                    Invoke-PptPersistence -Operation $Operation -Params $Arguments.operationArguments
                }
            }
            default {
                if ($Operation -notin $script:PptOperations -and $Operation -notin $script:PptCommonOperations) { throw [PptActionException]::new('PPT_CAPABILITY_UNAVAILABLE', 'Unknown PPT bridge operation.') }
                $data = Invoke-CoordinatedWpsCall {
                    Assert-BoundPresentation -DocumentId ([string]$Arguments.documentId)
                    if($Operation -in $script:PptCommonOperations){Invoke-PptCommonAction -Document $script:Document -Operation $Operation -Params $Arguments.operationArguments}
                    else{Invoke-PptAction -Document $script:Document -Operation $Operation -Params $Arguments.operationArguments}
                }
            }
        }
        return New-SuccessRecord -RequestId $RequestId -Data $data
    }
    catch {
        $code = 'PPT_READ_FAILED'
        $outcome = 'failed'
        $binding = 'unchanged'
        if ($_.Exception -is [PersistenceActionException]) { $code = $_.Exception.Code }
        elseif ($_.Exception -is [PptActionException]) { $code = $_.Exception.Code }
        elseif ($_.Exception -is [DocumentLeaseConflictException]) { $code = 'DOCUMENT_LEASE_CONFLICT' }
        elseif ($_.Exception -is [DocumentQuarantinedException]) { $code = 'DOCUMENT_QUARANTINED' }
        elseif ($_.Exception -is [IO.FileNotFoundException]) { $code = 'DOCUMENT_NOT_FOUND' }
        elseif ($_.Exception -is [UnauthorizedAccessException]) { $code = 'DOCUMENT_ACCESS_DENIED' }
        elseif ($Operation -eq 'prepare_existing_document') { $code = 'DOCUMENT_OPEN_FAILED' }
        if ($Operation -eq 'acquire_coordination_guard' -and $code -in @('DOCUMENT_LEASE_CONFLICT', 'DOCUMENT_QUARANTINED')) {
            Release-AdditionalFences -Clean $true
        }
        if ($script:ActionMayHaveEffect) {
            $outcome = 'unknown'
            if ($code -eq 'PPT_READ_FAILED') { $code = 'PPT_WRITE_FAILED' }
        }
        if ($Operation -eq 'acquire_existing_document' -and $code -eq 'DOCUMENT_OPEN_FAILED' -and -not $script:ActionMayHaveEffect) {
            return New-FailureRecord -RequestId $RequestId -Outcome 'failed' -Code $code -Message $_.Exception.Message -BindingDisposition 'unchanged'
        }
        if ($Operation -eq 'saveAs' -and $script:ActionMayHaveEffect) { $code='OUTPUT_WRITE_FAILED'; $binding='unprovable'; $outcome='unknown' }
        if ($code -eq 'DOCUMENT_CLOSED') { $binding = 'lost'; $outcome = 'failed' }
        elseif ($code -eq 'DOCUMENT_BINDING_UNAVAILABLE' -or
                $Operation -in @('acquire_existing_document', 'acquire_new_document', 'commit_document_lease', 'release_document_resources') -or
                ($Operation -eq 'acquire_coordination_guard' -and $code -notin @('DOCUMENT_LEASE_CONFLICT', 'DOCUMENT_QUARANTINED'))) {
            $binding = 'unprovable'; $outcome = 'unknown'
        }
        return New-FailureRecord -RequestId $RequestId -Outcome $outcome -Code $code -Message $_.Exception.Message -BindingDisposition $binding
    }
    finally { Release-PptOperationReferences }
}

. (Join-Path $PSScriptRoot '../../windows/bridge_loop.ps1')
