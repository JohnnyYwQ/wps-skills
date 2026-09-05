"""Read-only verification of the bound WPS document desktop window."""
import ctypes
from ctypes import wintypes
import os

def require_desktop():
    if os.name != 'nt':
        raise RuntimeError('请在 Windows 桌面的 PowerShell 中运行此演示。')
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.ProcessIdToSessionId.argtypes = (wintypes.DWORD, ctypes.POINTER(wintypes.DWORD))
    kernel.ProcessIdToSessionId.restype = wintypes.BOOL
    session = wintypes.DWORD()
    if not kernel.ProcessIdToSessionId(os.getpid(), ctypes.byref(session)):
        raise ctypes.WinError(ctypes.get_last_error())
    if session.value == 0:
        raise RuntimeError('当前是会话 0，窗口无法展示到你的桌面。请在 Windows 桌面打开 PowerShell 后执行同一命令。')
    return session.value


def visible_document_window(window, *, user32=None):
    """Verify the bound document's own HWND, including WPS's shared top-level frame."""
    if not window or not window.get('hwnd') or not window.get('processId'):
        return None
    user = user32 if user32 is not None else ctypes.WinDLL('user32', use_last_error=True)
    user.IsWindow.argtypes = (wintypes.HWND,)
    user.IsWindow.restype = wintypes.BOOL
    user.GetAncestor.argtypes = (wintypes.HWND, wintypes.UINT)
    user.GetAncestor.restype = wintypes.HWND
    user.GetWindowThreadProcessId.argtypes = (wintypes.HWND, ctypes.POINTER(wintypes.DWORD))
    user.GetWindowThreadProcessId.restype = wintypes.DWORD
    user.IsWindowVisible.argtypes = (wintypes.HWND,)
    user.IsWindowVisible.restype = wintypes.BOOL
    user.IsIconic.argtypes = (wintypes.HWND,)
    user.IsIconic.restype = wintypes.BOOL
    user.GetWindowRect.argtypes = (wintypes.HWND, ctypes.POINTER(wintypes.RECT))
    user.GetWindowRect.restype = wintypes.BOOL
    user.GetForegroundWindow.restype = wintypes.HWND
    child = window['hwnd']
    if not user.IsWindow(child):
        return None
    owner = wintypes.DWORD()
    user.GetWindowThreadProcessId(child, ctypes.byref(owner))
    if owner.value != window['processId']:
        return None
    hwnd = user.GetAncestor(child, 2)
    if not hwnd or not user.IsWindowVisible(hwnd) or user.IsIconic(hwnd):
        return None
    rect = wintypes.RECT()
    if not user.GetWindowRect(hwnd, ctypes.byref(rect)):
        return None
    width, height = rect.right - rect.left, rect.bottom - rect.top
    if width < 400 or height < 250:
        return None
    return {'hwnd': int(hwnd), 'documentHwnd': child, 'processId': owner.value,
            'width': width, 'height': height, 'foreground': hwnd == user.GetForegroundWindow()}


