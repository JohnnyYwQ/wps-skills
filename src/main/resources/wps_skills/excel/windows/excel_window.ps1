# Presentation uses the retained Workbook only. Focus never selects a target.
function Show-ExcelWorkbookWindow {
    param([bool]$ApplicationCreated)
    $windows = $null; $window = $null
    $script:Application.Visible = $true
    $script:Document.Activate() | Out-Null
    try {
        $windows = $script:Document.Windows
        if ([int]$windows.Count -lt 1) { return }
        $window = $windows.Item(1)
        $window.Visible = $true
        # xlNormal=-4143, xlMinimized=-4140. Preserve existing maximized windows.
        if ($ApplicationCreated -or [int]$script:Application.WindowState -eq -4140) {
            $script:Application.WindowState = -4143
        }
        if ([int]$window.WindowState -eq -4140) { $window.WindowState = -4143 }
        $window.Activate() | Out-Null
        Repair-WpsWindowBounds -Window $window
        $child = [IntPtr]([int64]$window.Hwnd)
        $root = [WpsSkills.NativeWindowPlacement]::GetAncestor($child, 2)
        if ($root -ne [IntPtr]::Zero) {
            if ([WpsSkills.NativeWindowPlacement]::IsIconic($root)) {
                [void][WpsSkills.NativeWindowPlacement]::ShowWindowAsync($root, 9)
            }
            [void][WpsSkills.NativeWindowPlacement]::SetForegroundWindow($root)
        }
    }
    catch {
        # Display availability is checked by the desktop demo. A headless caller
        # can still use an exact binding; window focus is not document identity.
    }
    finally {
        Release-ExcelReference -Value $window
        Release-ExcelReference -Value $windows
    }
}

function Get-ExcelWorkbookWindowInfo {
    $windows=$null; $window=$null
    try {
        $windows=$script:Document.Windows
        $window=$windows.Item(1)
        $handle=[IntPtr][int64]$window.Hwnd
        $owner=[uint32]0
        [void][WpsSkills.NativeWindowPlacement]::GetWindowThreadProcessId($handle,[ref]$owner)
        if ($handle -eq [IntPtr]::Zero -or $owner -eq 0) { return $null }
        return [ordered]@{hwnd=$handle.ToInt64(); processId=[long]$owner}
    } catch { return $null }
    finally { Release-ExcelReference $window; Release-ExcelReference $windows }
}
