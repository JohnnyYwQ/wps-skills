"""Verify the Word PowerShell launcher, saved demo and exact native window."""
import argparse
import json
import os
from pathlib import Path
import subprocess
import sys

sys.path[:0] = [str(Path(__file__).resolve().parents[4] / 'main/python'),
                str(Path(__file__).resolve().parents[2])]
from tests.windows.console_observer import run
from tests.word.live_acceptance import verify_demo

parser = argparse.ArgumentParser()
parser.add_argument('--repo', type=Path, required=True)
parser.add_argument('--output', type=Path, required=True)
args = parser.parse_args()
args.output.mkdir(parents=True, exist_ok=False)
python = str(Path(sys.executable).with_name('python.exe'))
command = [str(Path(os.environ['WINDIR']) / 'System32/WindowsPowerShell/v1.0/powershell.exe'),
           '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File',
           str(args.repo / 'scripts/demo/word.ps1'), '-PythonPath', python,
           '-OutputDirectory', str(args.output / 'demo'), '-Delay', '0']
report = {'status': 'running'}
try:
    report['launcher'] = run(command, args.output / 'demo.log', flags=subprocess.CREATE_NO_WINDOW)
    assert report['launcher']['returnCode'] == 0, 'Word demo failed; inspect demo.log'
    assert not report['launcher']['newVisibleConsoles'], 'Word launcher opened a helper console'
    demo = json.loads((args.output / 'demo/report.json').read_text(encoding='utf-8'))
    assert demo['status'] == 'passed'
    report.update(verify_demo(demo), status='passed')
except BaseException as exc:
    report.update(status='failed', error=str(exc))
    raise
finally:
    (args.output / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
