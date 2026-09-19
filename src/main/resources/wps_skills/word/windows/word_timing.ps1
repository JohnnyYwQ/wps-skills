# Word retains its trace component name; phase recording is shared.
. (Join-Path $PSScriptRoot '../../windows/application_timing.ps1')
function Write-WordBridgeTiming {
    param([string]$RequestId, [string]$TraceId, [string]$Operation,
          [string]$Phase, [double]$DurationMs, [string]$Outcome = '')
    Write-ApplicationBridgeTiming @PSBoundParameters -Component 'word_bridge_timing'
}
