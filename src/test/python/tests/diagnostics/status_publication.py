"""Windows regression: publish status through a short-lived reader lacking delete sharing."""
import argparse
import ctypes
from ctypes import wintypes
import json
from pathlib import Path
import sys
import tempfile
import threading
import time


def main(module_dir):
    sys.path.insert(0,str(module_dir));from common import write_json,read_json
    api=ctypes.WinDLL('kernel32',use_last_error=True)
    api.CreateFileW.argtypes=[wintypes.LPCWSTR,wintypes.DWORD,wintypes.DWORD,ctypes.c_void_p,wintypes.DWORD,wintypes.DWORD,wintypes.HANDLE]
    api.CreateFileW.restype=wintypes.HANDLE
    api.CloseHandle.argtypes=[wintypes.HANDLE]
    with tempfile.TemporaryDirectory(prefix='wps-status-publication-') as temporary:
        path=Path(temporary)/'status.json';write_json(path,{'state':'old'})
        # Actual Windows reader permits data access but not replacement until it closes.
        handle=api.CreateFileW(str(path),0x80000000,3,None,3,0,None)
        assert handle not in (None,ctypes.c_void_p(-1).value)
        thread=threading.Thread(target=lambda:(time.sleep(.35),api.CloseHandle(handle)))
        thread.start();started=time.perf_counter()
        try:
            write_json(path,{'state':'published'})
            assert read_json(path)=={'state':'published'}
            assert .25<=time.perf_counter()-started<3
            print(json.dumps({'state':'passed','seconds':time.perf_counter()-started}))
        except OSError as error:
            print(json.dumps({'state':'failed','winerror':error.winerror,'error':str(error)}));return 1
        finally:thread.join()
    return 0

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--module-dir',type=Path,required=True)
    raise SystemExit(main(parser.parse_args().module_dir))
