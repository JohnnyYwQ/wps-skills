# Test-package-only fault and observation hooks. Never copied into production distributions.
$script:DesignConfig = $null
$script:DesignFired = $false
if ($env:WPS_DESIGN_SPEC) { $script:DesignConfig = [IO.File]::ReadAllText($env:WPS_DESIGN_SPEC) | ConvertFrom-Json }
function Write-DesignEvent {
    param([string]$Event, $Data)
    if (-not $script:DesignConfig) { return }
    $record=@{event=$Event;data=$Data;pid=$PID;utc=[DateTime]::UtcNow.ToString('o')}
    [IO.File]::AppendAllText((Join-Path $script:DesignConfig.caseDir 'native-events.jsonl'),($record|ConvertTo-Json -Depth 15 -Compress)+"`n",(New-Object Text.UTF8Encoding($false)))
}
function Invoke-DesignHook {
    param([string]$Stage,[string]$Operation)
    if (-not $script:DesignConfig) { return }
    if($Stage -eq 'before') { $script:DesignCurrentOperation=$Operation; Write-DesignEvent 'operation' @{operation=$Operation} }
    if($script:DesignFired -or $Operation -ne $script:DesignConfig.operation -or $Stage -ne $script:DesignConfig.stage) { return }
    $script:DesignFired=$true
    Write-DesignEvent 'injection' @{stage=$Stage;operation=$Operation;mode=$script:DesignConfig.mode}
    if($script:DesignConfig.mode -in @('focus','active_candidate')) {
        $path=[IO.Path]::GetFullPath([string]$script:DesignConfig.otherPath)
        $root=[IO.Path]::GetFullPath([string]$script:DesignConfig.ownedRoot).TrimEnd('\')+'\'
        if(-not $path.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)) {throw 'Test document outside owned root'}
        if($script:DesignConfig.app -eq 'excel') {
            $other=$script:Application.Workbooks.Open($path)
            [void]$other.Activate()
            if($script:DesignConfig.mode -eq 'active_candidate') {
                # Change only the worksheet target choice. Preserve the current observation identity.
                $script:WorksheetBindings['实验'].sheet=$script:Application.ActiveWorkbook.Worksheets.Item('实验')
            }
            $active=[string]$script:Application.ActiveWorkbook.FullName
        } elseif($script:DesignConfig.app -eq 'word') {
            $other=$script:Application.Documents.Open($path)
            [void]$other.Activate();$active=[string]$script:Application.ActiveDocument.FullName
        } else {
            $other=$script:Application.Presentations.Open($path)
            [void]$other.Windows.Item(1).Activate();$active=[string]$script:Application.ActivePresentation.FullName
        }
        if(-not [string]::Equals($active,$path,[StringComparison]::OrdinalIgnoreCase)){throw 'Focus switch did not happen'}
        Write-DesignEvent 'focus_confirmed' @{path=$active}
    } elseif($script:DesignConfig.mode -eq 'drop_response') {
        [IO.File]::WriteAllText((Join-Path $script:DesignConfig.caseDir 'ready.json'),'{"phase":"effect_without_response"}')
        exit 71
    } elseif($script:DesignConfig.mode -eq 'pause') {
        $info=@{phase=$Stage;pid=$PID;coordinationPath=$script:CoordinationStatePath;operation=$Operation}
        [IO.File]::WriteAllText((Join-Path $script:DesignConfig.caseDir 'ready.json'),($info|ConvertTo-Json -Compress))
        $watch=[Diagnostics.Stopwatch]::StartNew()
        while(-not (Test-Path (Join-Path $script:DesignConfig.caseDir 'release'))) {
            if($watch.Elapsed.TotalSeconds -gt 45) {throw 'Design barrier was not released'}
            Start-Sleep -Milliseconds 20
        }
    }
}
