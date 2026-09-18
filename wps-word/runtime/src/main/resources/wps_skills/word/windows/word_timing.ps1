# Diagnostics never enter the JSON protocol or change an operation outcome.
function Write-WordBridgeTiming {
    param([string]$RequestId, [string]$TraceId, [string]$Operation,
          [string]$Phase, [double]$DurationMs, [string]$Outcome = '')
    try {
        $root = $env:WPS_TRACE_DIR
        if ([string]::IsNullOrWhiteSpace($root)) {
            $root = Join-Path $env:LOCALAPPDATA 'wps-skills/logs'
        }
        $directory = Join-Path $root 'bridges'
        [void][IO.Directory]::CreateDirectory($directory)
        $record = [ordered]@{
            ts = [DateTime]::UtcNow.ToString('o')
            component = 'word_bridge_timing'
            event = 'phase.finished'
            bridgeRequestId = $RequestId
            traceId = $TraceId
            operation = $Operation
            name = $Phase
            durationMs = $DurationMs
            outcome = $Outcome
            pid = $PID
        }
        $line = ConvertTo-Json -InputObject $record -Depth 4 -Compress
        [IO.File]::AppendAllText((Join-Path $directory ('bridge-' + $PID + '.jsonl')),
            $line + [Environment]::NewLine, (New-Object Text.UTF8Encoding($false)))
    }
    catch { }
}
