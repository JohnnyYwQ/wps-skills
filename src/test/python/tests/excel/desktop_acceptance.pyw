"""Observe the complete Excel demo and optional controls on the Windows desktop."""
import argparse
import json
import os
from pathlib import Path
import subprocess
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from tests.windows.console_observer import run


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--repo', required=True, type=Path)
    parser.add_argument('--output', required=True, type=Path)
    parser.add_argument('--launcher', action='store_true')
    parser.add_argument('--with-control', action='store_true')
    parser.add_argument('--name-conflict', action='store_true')
    args = parser.parse_args()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    python = str(Path(sys.executable).with_name('python.exe'))
    report = {'status': 'running'}
    try:
        # Same production demo, with no cmd wrapper and no initial console.
        command = [python, '-X', 'utf8', str(args.repo/'scripts/demo/excel.py'),
                   '--output-dir', str(output/'demo'), '--delay', '0']
        if args.launcher:
            command = [str(Path(os.environ['WINDIR'])/'System32/WindowsPowerShell/v1.0/powershell.exe'),
                       '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                       '-File', str(args.repo/'scripts/demo/excel.ps1'), '-PythonPath', python,
                       '-OutputDirectory', str(output/'demo'), '-Delay', '0']
        report['directDemo'] = run(command, output/'demo.log', flags=subprocess.CREATE_NO_WINDOW)
        # Positive control proves that the monitor catches a real flashing console.
        if args.with_control:
            report['cmdControl'] = run(['cmd.exe', '/d', '/c', 'ping -n 3 127.0.0.1 > nul'], output/'control.log', flags=0)
            assert report['cmdControl']['newVisibleConsoles'], 'Monitor failed to detect the positive control console'
        assert report['directDemo']['returnCode'] == 0, 'Desktop demonstration failed'
        assert not report['directDemo']['newVisibleConsoles'], 'Production demo displayed a helper console'
        if args.name_conflict:
            env_before = os.environ.get('PYTHONPATH')
            os.environ['PYTHONPATH'] = str(args.repo/'src/main/python')
            try:
                report['nameConflict'] = run([python, '-X', 'utf8', str(args.repo/'src/test/python/tests/excel/live_name_conflict.py'),
                    '--output-dir', str(output/'name-conflict')], output/'name-conflict.log', flags=subprocess.CREATE_NO_WINDOW)
            finally:
                if env_before is None:
                    os.environ.pop('PYTHONPATH', None)
                else:
                    os.environ['PYTHONPATH'] = env_before
            assert report['nameConflict']['returnCode'] == 0, 'Same-name conflict check failed'
            assert not report['nameConflict']['newVisibleConsoles'], 'Same-name check displayed a console'
        report['status'] = 'passed'
    except BaseException as exc:
        report.update(status='failed', error=repr(exc))
        raise
    finally:
        (output/'report.json').write_text(json.dumps(report, indent=2), encoding='utf-8')


if __name__ == '__main__':
    main()
