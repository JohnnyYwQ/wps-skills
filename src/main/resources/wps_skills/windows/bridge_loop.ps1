try {
    while ($null -ne ($line = [Console]::In.ReadLine())) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        $requestId = 'unparsed'
        try {
            $request = $line | ConvertFrom-Json -ErrorAction Stop
            if (-not (Test-ExactFields -Value $request -Expected @('requestId', 'traceId', 'operation', 'arguments'))) {
                throw 'A bridge request has invalid fields.'
            }
            $requestId = [string]$request.requestId
            if ([string]::IsNullOrEmpty($requestId)) {
                throw 'A bridge requestId must be non-empty.'
            }
            if ([string]::IsNullOrEmpty([string]$request.traceId)) {
                throw 'A bridge traceId must be non-empty.'
            }
            if ([string]::IsNullOrEmpty([string]$request.operation)) {
                throw 'A bridge operation must be non-empty.'
            }
            if ($request.arguments -isnot [pscustomobject]) {
                throw 'Bridge arguments must be an object.'
            }
            $response = Invoke-BridgeOperation `
                -RequestId $requestId `
                -Operation ([string]$request.operation) `
                -Arguments $request.arguments
        }
        catch {
            $response = New-FailureRecord `
                -RequestId $requestId `
                -Outcome failed `
                -Code 'INVALID_PARAMS' `
                -Message $_.Exception.Message `
                -BindingDisposition unchanged
        }
        Write-BridgeRecord -Record $response
    }
}
finally {
    try {
        Release-CoordinationResources -Clean $true
    }
    catch {
    }
    if (Get-Command Release-ApplicationResources -ErrorAction SilentlyContinue) { Release-ApplicationResources }
    Release-ComReference -Value $script:Document
    Release-ComReference -Value $script:Documents
    Release-ComReference -Value $script:Application
    $script:Document = $null
    $script:Documents = $null
    $script:Application = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
