# WPS hides DocumentWindow.HWND from IDispatch. The declared native interface
# still implements get_HWND (slot 34; see type_library/wps_ppt_api.py evidence).
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace WpsSkills {
 public static class PptDocumentWindow {
  [UnmanagedFunctionPointer(CallingConvention.StdCall)]
  private delegate int GetHwnd(IntPtr self, out int hwnd);
  public static long Handle(object window) {
   IntPtr unknown=Marshal.GetIUnknownForObject(window), documentWindow=IntPtr.Zero;
   try {
    Guid iid=new Guid("91493457-5A91-11CF-8700-00AA0060263B");
    Marshal.ThrowExceptionForHR(Marshal.QueryInterface(unknown,ref iid,out documentWindow));
    IntPtr table=Marshal.ReadIntPtr(documentWindow);
    var getter=(GetHwnd)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(table,34*IntPtr.Size),typeof(GetHwnd));
    int handle; Marshal.ThrowExceptionForHR(getter(documentWindow,out handle));
    return (long)(uint)handle;
   }finally{if(documentWindow!=IntPtr.Zero)Marshal.Release(documentWindow);Marshal.Release(unknown);}
  }
 }
}
'@

function Get-PptPresentationWindowInfo {
    $windows=$null;$window=$null
    try {
        $windows=$script:Document.Windows
        if([int]$windows.Count -lt 1){return $null}
        $window=$windows.Item(1)
        $handle=[IntPtr][WpsSkills.PptDocumentWindow]::Handle($window)
        $owner=[uint32]0
        [void][WpsSkills.NativeWindowPlacement]::GetWindowThreadProcessId($handle,[ref]$owner)
        if($handle -eq [IntPtr]::Zero -or $owner -eq 0){return $null}
        return [ordered]@{hwnd=$handle.ToInt64();processId=[long]$owner}
    }catch{return $null}
    finally{Release-PptReference $window;Release-PptReference $windows}
}

function Show-PptPresentationWindow {
    param([bool]$ApplicationCreated)
    $windows=$null;$window=$null
    try {
        $script:Application.Visible=-1
        $windows=$script:Document.Windows
        if([int]$windows.Count -lt 1){return}
        $window=$windows.Item(1)
        if([int]$window.WindowState -eq 2){$window.WindowState=1}
        $window.Activate()|Out-Null
        $handle=[WpsSkills.PptDocumentWindow]::Handle($window)
        Repair-WpsWindowBounds -Window ([pscustomobject]@{Hwnd=$handle})
        $root=[WpsSkills.NativeWindowPlacement]::GetAncestor([IntPtr]$handle,2)
        if($root -ne [IntPtr]::Zero){
            if([WpsSkills.NativeWindowPlacement]::IsIconic($root)){[void][WpsSkills.NativeWindowPlacement]::ShowWindowAsync($root,9)}
            [void][WpsSkills.NativeWindowPlacement]::SetForegroundWindow($root)
        }
    }catch{
        # Headless sessions may still bind exactly; the desktop demo verifies visibility.
    }finally{Release-PptReference $window;Release-PptReference $windows}
}
