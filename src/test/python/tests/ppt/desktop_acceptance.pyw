"""Observe the complete PPT demo launcher for unwanted visible console windows."""
import argparse
import json
import os
from pathlib import Path
import subprocess
import sys

parser=argparse.ArgumentParser()
parser.add_argument('--repo',type=Path,required=True)
parser.add_argument('--output',type=Path,required=True)
args=parser.parse_args()
args.output.mkdir(parents=True,exist_ok=False)
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from tests.windows.console_observer import run
python=str(Path(sys.executable).with_name('python.exe'))
command=[str(Path(os.environ['WINDIR'])/'System32/WindowsPowerShell/v1.0/powershell.exe'),
         '-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(args.repo/'scripts/demo/ppt.ps1'),
         '-PythonPath',python,'-OutputDirectory',str(args.output/'demo'),'-Delay','0']
report={'status':'running'}
try:
    report['launcher']=run(command,args.output/'demo.log',flags=subprocess.CREATE_NO_WINDOW)
    assert report['launcher']['returnCode']==0,'PPT desktop demo failed'
    assert not report['launcher']['newVisibleConsoles'],'PPT launcher created an auxiliary console'
    demo=json.loads((args.output/'demo/report.json').read_text(encoding='utf-8'))
    assert demo['status']=='passed' and demo['visibleWindow'],'Bound PPT window was not verified'
    report['status']='passed'
except BaseException as exc:
    report['status']='failed';report['error']=str(exc)
    raise
finally:
    (args.output/'report.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
