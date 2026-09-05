# Shared sizing of the exact bound WPS document window.

function Repair-WpsWindowBounds {
    param([Parameter(Mandatory = $true)]$Window)

    # WPS's COM Width/Height can diverge from the visible top-level frame.
    # Resolve the bound document's HWND to its root and repair only an
    # implausibly small restored rectangle. 2 is GA_ROOT and
    # MONITOR_DEFAULTTONEAREST; 0x14 is SWP_NOZORDER | SWP_NOACTIVATE.
    $childHwnd = [IntPtr]([int64]$Window.Hwnd)
    if ($childHwnd -eq [IntPtr]::Zero) {
        return
    }
    $rootHwnd = [WpsSkills.NativeWindowPlacement]::GetAncestor($childHwnd, 2)
    if ($rootHwnd -eq [IntPtr]::Zero) {
        return
    }

    $rootRect = New-Object WpsSkills.WindowRect
    $monitor = [WpsSkills.NativeWindowPlacement]::MonitorFromWindow($rootHwnd, 2)
    $monitorInfo = New-Object WpsSkills.MonitorInfo
    $monitorInfo.Size = [Runtime.InteropServices.Marshal]::SizeOf($monitorInfo)
    if (
        $monitor -eq [IntPtr]::Zero -or
        -not [WpsSkills.NativeWindowPlacement]::GetWindowRect(
            $rootHwnd,
            [ref]$rootRect
        ) -or
        -not [WpsSkills.NativeWindowPlacement]::GetMonitorInfo(
            $monitor,
            [ref]$monitorInfo
        )
    ) {
        return
    }

    $workWidth = $monitorInfo.Work.Right - $monitorInfo.Work.Left
    $workHeight = $monitorInfo.Work.Bottom - $monitorInfo.Work.Top
    $currentWidth = $rootRect.Right - $rootRect.Left
    $currentHeight = $rootRect.Bottom - $rootRect.Top
    if (
        $workWidth -le 0 -or
        $workHeight -le 0 -or
        (
            $currentWidth -ge ($workWidth * 0.68) -and
            $currentHeight -ge ($workHeight * 0.68)
        )
    ) {
        return
    }

    $targetWidth = [int][math]::Round($workWidth * 0.8)
    $targetHeight = [int][math]::Round($workHeight * 0.8)
    $targetLeft = $monitorInfo.Work.Left + [int][math]::Round(
        ($workWidth - $targetWidth) / 2
    )
    $targetTop = $monitorInfo.Work.Top + [int][math]::Round(
        ($workHeight - $targetHeight) / 2
    )
    [void][WpsSkills.NativeWindowPlacement]::SetWindowPos(
        $rootHwnd,
        [IntPtr]::Zero,
        $targetLeft,
        $targetTop,
        $targetWidth,
        $targetHeight,
        0x14
    )
}

