# Test-package-only observations. Exceptions and the production fallback policy
# are preserved; no registry mutation, application shutdown, or hidden replay.
function Get-DiagnosticExceptionChain {
    param([Exception]$Exception)
    $items = New-Object Collections.ArrayList
    while ($null -ne $Exception) {
        [void]$items.Add([ordered]@{
            type=$Exception.GetType().FullName
            message=$Exception.Message
            hresult=('0x{0:X8}' -f $Exception.HResult)
        })
        $Exception = $Exception.InnerException
    }
    return $items.ToArray()
}

function Write-DiagnosticComEvent {
    param([string]$Stage, [string]$Event, [string]$Outcome='', [double]$DurationMs=0,
          [string]$SpanId='', $Data=$null, $Errors=@())
    try {
        if ([string]::IsNullOrEmpty($env:WPS_COM_DIAGNOSTIC_DIR)) { return }
        [void][IO.Directory]::CreateDirectory($env:WPS_COM_DIAGNOSTIC_DIR)
        $record = [ordered]@{
            event=$Event; stage=$Stage; outcome=$Outcome; durationMs=$DurationMs
            spanId=$SpanId; pid=$PID; utc=[DateTime]::UtcNow.ToString('o')
            monotonicTicks=[Diagnostics.Stopwatch]::GetTimestamp()
            clockFrequency=[Diagnostics.Stopwatch]::Frequency
            data=$Data; errors=@($Errors)
        }
        $line = ConvertTo-Json -InputObject $record -Depth 12 -Compress
        [IO.File]::AppendAllText((Join-Path $env:WPS_COM_DIAGNOSTIC_DIR ('com-' + $PID + '.jsonl')),
            $line + [Environment]::NewLine, (New-Object Text.UTF8Encoding($false)))
    } catch {
        # Missing diagnostics will fail the test's evidence check, but must not
        # replace a COM error or prevent the existing cleanup path.
        [Console]::Error.WriteLine('[WPS-COM-DIAGNOSTIC-WRITE] ' + $_.Exception.Message)
    }
}

function Get-DiagnosticRegistrationDetails {
    param([string]$ProgId)
    $entries = New-Object Collections.ArrayList
    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)) {
        foreach ($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryHive]::LocalMachine)) {
            $base=$null; $key=$null; $server=$null
            $item=[ordered]@{hive=[string]$hive; view=[string]$view; clsid=$null; localServer=$null; error=$null}
            try {
                $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
                $key=$base.OpenSubKey(('Software\Classes\' + $ProgId + '\CLSID'))
                if ($null -ne $key) {
                    $item.clsid=[string]$key.GetValue('')
                    if ($item.clsid) {
                        $server=$base.OpenSubKey(('Software\Classes\CLSID\' + $item.clsid + '\LocalServer32'))
                        if ($null -ne $server) { $item.localServer=[string]$server.GetValue('') }
                    }
                }
            } catch { $item.error=@(Get-DiagnosticExceptionChain $_.Exception) }
            finally {
                if ($null -ne $server) { $server.Dispose() }
                if ($null -ne $key) { $key.Dispose() }
                if ($null -ne $base) { $base.Dispose() }
            }
            [void]$entries.Add($item)
        }
    }
    return [ordered]@{progId=$ProgId; registrations=@($entries.ToArray())}
}

function Write-DiagnosticComEnvironment {
    $identity=$null; $process=$null
    try {
        $identity=[Security.Principal.WindowsIdentity]::GetCurrent()
        $principal=[Security.Principal.WindowsPrincipal]::new($identity)
        $process=[Diagnostics.Process]::GetCurrentProcess()
        $data=[ordered]@{
            powershellVersion=$PSVersionTable.PSVersion.ToString()
            powershellEdition=[string]$PSVersionTable.PSEdition
            executable=$process.MainModule.FileName
            processBits=([IntPtr]::Size * 8)
            os64Bit=[Environment]::Is64BitOperatingSystem
            clrVersion=[Environment]::Version.ToString()
            osVersion=[Environment]::OSVersion.VersionString
            sessionId=$process.SessionId
            elevated=$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            apartmentState=[string][Threading.Thread]::CurrentThread.GetApartmentState()
            registration=(Get-DiagnosticRegistrationDetails 'KET.Application')
        }
        Write-DiagnosticComEvent -Stage 'environment' -Event 'stage.finished' -Outcome 'succeeded' -Data $data
    } catch {
        Write-DiagnosticComEvent -Stage 'environment' -Event 'stage.finished' -Outcome 'failed' -Errors @(Get-DiagnosticExceptionChain $_.Exception)
    } finally {
        if ($null -ne $identity) { $identity.Dispose() }
        if ($null -ne $process) { $process.Dispose() }
    }
}

function Invoke-DiagnosticComStage {
    param([string]$Stage, [scriptblock]$Body, $Details=$null)
    $span=[guid]::NewGuid().ToString('N')
    Write-DiagnosticComEvent -Stage $Stage -Event 'stage.started' -SpanId $span -Data $Details
    $timer=[Diagnostics.Stopwatch]::StartNew()
    try {
        $value = & $Body
        Write-DiagnosticComEvent -Stage $Stage -Event 'stage.finished' -SpanId $span -Outcome 'succeeded' -DurationMs $timer.Elapsed.TotalMilliseconds -Data $Details
        return ,$value
    } catch {
        $errors=@(Get-DiagnosticExceptionChain $_.Exception)
        $outcome='failed'
        if ($Stage -eq 'com.attach' -and @($errors | Where-Object { $_.hresult -eq '0x800401E3' }).Count -gt 0) { $outcome='not_found' }
        Write-DiagnosticComEvent -Stage $Stage -Event 'stage.finished' -SpanId $span -Outcome $outcome -DurationMs $timer.Elapsed.TotalMilliseconds -Errors $errors -Data $Details
        throw
    }
}

function Write-DiagnosticApplicationInfo {
    param($Application, [string]$Connection)
    $data=[ordered]@{connection=$Connection; version=$null; versionError=$null}
    try { $data.version=[string]$Application.Version }
    catch { $data.versionError=@(Get-DiagnosticExceptionChain $_.Exception) }
    Write-DiagnosticComEvent -Stage 'com.application_info' -Event 'stage.finished' -Outcome 'observed' -Data $data
}

function Get-DiagnosticActiveApplication {
    param([string]$ProgId)
    $application = Invoke-DiagnosticComStage -Stage 'com.attach' -Details @{progId=$ProgId} -Body {
        [Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
    }
    Write-DiagnosticApplicationInfo -Application $application -Connection 'existing'
    return ,$application
}

function New-DiagnosticApplication {
    param([string]$ProgId)
    $application = Invoke-DiagnosticComStage -Stage 'com.activate' -Details @{progId=$ProgId} -Body {
        New-Object -ComObject $ProgId
    }
    Write-DiagnosticApplicationInfo -Application $application -Connection 'activation'
    return ,$application
}

function Get-DiagnosticExcelComAvailability {
    $span=[guid]::NewGuid().ToString('N')
    Write-DiagnosticComEvent -Stage 'com.registration' -Event 'stage.started' -SpanId $span
    $timer=[Diagnostics.Stopwatch]::StartNew()
    try {
        $value=Get-ExcelComAvailability
        $outcome=if ($value.available) {'succeeded'} else {'failed'}
        Write-DiagnosticComEvent -Stage 'com.registration' -Event 'stage.finished' -SpanId $span -Outcome $outcome -DurationMs $timer.Elapsed.TotalMilliseconds -Data $value
        return $value
    } catch {
        Write-DiagnosticComEvent -Stage 'com.registration' -Event 'stage.finished' -SpanId $span -Outcome 'failed' -DurationMs $timer.Elapsed.TotalMilliseconds -Errors @(Get-DiagnosticExceptionChain $_.Exception)
        throw
    }
}

function Invoke-AcquireNewDocument {
    param($Arguments)
    return Invoke-DiagnosticComStage -Stage 'document.bind_create' -Body {
        Invoke-OriginalAcquireNewDocument -Arguments $Arguments
    }
}

function Invoke-AcquireExistingDocument {
    param($Arguments)
    return Invoke-DiagnosticComStage -Stage 'document.bind_open' -Body {
        Invoke-OriginalAcquireExistingDocument -Arguments $Arguments
    }
}

Write-DiagnosticComEnvironment
