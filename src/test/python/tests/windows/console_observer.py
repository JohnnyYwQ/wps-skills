"""Desktop-only native console-window observation for WPS launch regression."""
import ctypes
from ctypes import wintypes
from pathlib import Path
import subprocess
import threading
import time


def observe_windows():
    user = ctypes.WinDLL('user32', use_last_error=True)
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user.EnumWindows.argtypes = (callback_type, wintypes.LPARAM)
    user.IsWindowVisible.argtypes = (wintypes.HWND,)
    user.GetClassNameW.argtypes = (wintypes.HWND, wintypes.LPWSTR, ctypes.c_int)
    user.GetWindowThreadProcessId.argtypes = (wintypes.HWND, ctypes.POINTER(wintypes.DWORD))
    kernel.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.QueryFullProcessImageNameW.argtypes = (wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD))
    kernel.CloseHandle.argtypes = (wintypes.HANDLE,)
    found = []
    @callback_type
    def visit(hwnd, _):
        if not user.IsWindowVisible(hwnd):
            return True
        cls = ctypes.create_unicode_buffer(256)
        user.GetClassNameW(hwnd, cls, len(cls))
        if cls.value not in {'ConsoleWindowClass', 'CASCADIA_HOSTING_WINDOW_CLASS', 'PseudoConsoleWindow'}:
            return True
        pid = wintypes.DWORD()
        user.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        handle = kernel.OpenProcess(0x1000, False, pid.value)
        name = ''
        if handle:
            try:
                size = wintypes.DWORD(32768)
                path = ctypes.create_unicode_buffer(size.value)
                if kernel.QueryFullProcessImageNameW(handle, 0, path, ctypes.byref(size)):
                    name = Path(path.value).name
            finally:
                kernel.CloseHandle(handle)
        found.append({'hwnd': int(hwnd), 'class': cls.value, 'pid': pid.value, 'executable': name})
        return True
    user.EnumWindows(visit, 0)
    return found


def run(command, output, *, flags):
    baseline = {w['hwnd'] for w in observe_windows()}
    seen = {}
    stopped = threading.Event()
    def monitor():
        while not stopped.is_set():
            for window in observe_windows():
                if window['hwnd'] not in baseline:
                    seen.setdefault(window['hwnd'], window)
            stopped.wait(0.01)
    thread = threading.Thread(target=monitor, daemon=True)
    thread.start()
    try:
        with output.open('wb') as log:
            child = subprocess.Popen(command, stdout=log, stderr=subprocess.STDOUT, creationflags=flags)
            code = child.wait(timeout=120)
        time.sleep(0.5)
    finally:
        stopped.set()
        thread.join(timeout=2)
    return {'returnCode': code, 'newVisibleConsoles': list(seen.values())}
