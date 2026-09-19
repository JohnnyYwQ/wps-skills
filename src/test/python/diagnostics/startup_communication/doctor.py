"""Read-only environment checks, with explicit opt-in Excel document acceptance."""
import argparse
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import platform
import shutil
import subprocess
import sys
import uuid

APPS = {'word': 'KWPS.Application', 'excel': 'KET.Application', 'ppt': 'KWPP.Application'}
HERE = Path(__file__).resolve().parent


def check(name, passed, detail, remedy=''):
    return dict(name=name, state='passed' if passed else 'failed', detail=detail,
                remedy='' if passed else remedy)


def probe_checks(data, payload):
    return [
        check('powershell.version', data.get('edition') == 'Desktop' and data.get('powershellVersion', '').startswith('5.1'), data,
              'Use Windows PowerShell 5.1 (powershell.exe), not PowerShell Core (pwsh.exe).'),
        check('desktop.session', data.get('sessionId', 0) > 0 and data.get('userInteractive') is True, data.get('sessionId'),
              'Run from the logged-in Windows desktop; for remote execution use an interactive scheduled task. SSH Session 0 cannot validate WPS desktop automation.'),
        check('com.registration', data.get('registered') is True, data.get('progId'),
              'Install or repair the selected WPS application for this Windows user, then rerun. Registration alone does not prove usability.'),
        check('communication.utf8', data.get('echo') == payload, data.get('echo'),
              'Check Python/PowerShell UTF-8 stdin and stdout. Keep Task JSON in UTF-8 files.'),
    ]


def inspect(app, output, timeout=15):
    checks = [check('platform.windows', os.name == 'nt', platform.platform(), 'Run this doctor on the target Windows machine.'),
              check('python.version', sys.version_info >= (3, 10), sys.version, 'Install Python 3.10+ for this diagnostic tool; product requirements are separate.')]
    checks.append(check('output.writable', True, str(output)))
    if os.name != 'nt':
        checks.append(dict(name='windows.probe', state='not_run', detail='Windows required', remedy=''))
        return checks
    executable = Path(os.environ.get('WINDIR', r'C:\Windows')) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
    checks.append(check('powershell.executable', executable.is_file(), str(executable), 'Restore the system Windows PowerShell installation.'))
    if not executable.is_file():
        return checks
    script = HERE / 'environment.ps1'
    if not script.exists():
        script = HERE.parents[4] / 'src/test/resources/diagnostics/startup-communication/environment.ps1'
    payload = '中文通信检查 / 路径 空格 / "引号" / \\ / 换行\n / 🚀'
    try:
        # communicate sends EOF; no shell interpolation, profile, COM activation,
        # or descendants. subprocess.run kills and waits for this owned process
        # on timeout; no remote-SSH timeout is used to own this process.
        cp = subprocess.run([str(executable), '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                             '-File', str(script), '-ProgId', APPS[app]],
                            input=json.dumps({'payload': payload}, ensure_ascii=False)+'\n',
                            encoding='utf-8', capture_output=True, timeout=timeout)
        (output / 'probe.stderr.txt').write_text(cp.stderr, encoding='utf-8')
        (output / 'probe.stdout.txt').write_text(cp.stdout, encoding='utf-8')
        if cp.returncode:
            checks.append(check('windows.probe', False, {'exitCode': cp.returncode, 'stderr': cp.stderr}, 'Inspect probe.stderr.txt; check script execution policy and system restrictions.'))
        else:
            checks.extend(probe_checks(json.loads(cp.stdout.lstrip('\ufeff')), payload))
    except (OSError, ValueError, subprocess.TimeoutExpired) as error:
        checks.append(check('windows.probe', False, str(error), 'Inspect PowerShell availability and endpoint restrictions. The owned probe is stopped on timeout.'))
    return checks


def smoke(output):
    from common import read_json, verify_bundle
    from worker import run
    target = output / 'smoke'
    target.mkdir()
    if (HERE / 'manifest.json').exists():
        verify_bundle(HERE)
        shutil.copytree(HERE, target / 'bundle', ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
    else:
        from bundle import build, RESOURCES
        build(target / 'bundle', RESOURCES / 'com.json')
    # The portable package must carry the explicit COM acceptance profile.
    config = read_json(target / 'bundle/experiment.json')
    if config['cases'] != ['wps', 'wps-open'] or config['profiles']['quick']['maxRounds'] != 1:
        raise ValueError('Smoke test requires a package built with com.json (one create/open round).')
    code = run(target, 'quick')
    status = read_json(target / 'windows/status.json')
    return check('wps.smoke', code == 0 and status['state'] == 'completed' and len(status['cases']) == 2
                 and all(c['state'] == 'passed' for c in status['cases']), status,
                 'Inspect smoke/windows/cases/*/com-summary.json and Task evidence. Do not replay failed Tasks; use a new doctor run after diagnosis.')


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--app', choices=APPS, required=True)
    parser.add_argument('--smoke-test', action='store_true', help='Create/open test documents; currently Excel only')
    parser.add_argument('--output', type=Path, help='New evidence directory; existing directories are refused')
    args = parser.parse_args(argv)
    portable = (HERE / 'manifest.json').exists()
    default_root = HERE.parent if portable else HERE.parents[4] / 'build/runs/doctor'
    output = (args.output or default_root / ('doctor-' + datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S') + '-' + uuid.uuid4().hex[:8])).resolve()
    if portable and (output == HERE or HERE in output.parents):
        parser.error('Place evidence outside the portable bundle so its files and hashes remain immutable.')
    try:
        output.mkdir(parents=True, exist_ok=False)
        # Verify create/read/delete in exactly the evidence directory.
        probe = output / 'write-check.txt'
        probe.write_text('中文', encoding='utf-8')
        if probe.read_text(encoding='utf-8') != '中文':
            raise OSError('Output directory readback mismatch')
        probe.unlink()
    except OSError as error:
        print(json.dumps({'state': 'failed', 'check': 'output.writable', 'error': str(error)}, ensure_ascii=False))
        return 1
    report = dict(schemaVersion=1, app=args.app, utc=datetime.now(timezone.utc).isoformat(),
                  executable=sys.executable, checks=inspect(args.app, output),
                  scope='local Windows environment; SSH and remote scheduling are separate',
                  smokeRequested=args.smoke_test)
    if args.smoke_test:
        if args.app != 'excel':
            report['checks'].append(check('wps.smoke', False, 'Unsupported for '+args.app, 'Real document smoke tests currently support Excel only; basic checks support all three applications.'))
        elif any(c['state'] == 'failed' for c in report['checks']):
            report['checks'].append(dict(name='wps.smoke', state='not_run', detail='Prerequisite checks failed', remedy='Resolve the failed checks first.'))
        else:
            try:
                report['checks'].append(smoke(output))
            except Exception as error:
                report['checks'].append(check('wps.smoke', False, repr(error), 'Inspect preserved smoke evidence; rerun only in a new output directory.'))
    else:
        report['checks'].append(dict(name='wps.smoke', state='not_run', detail='Opt in with --smoke-test to validate real document operations', remedy=''))
    report['state'] = 'failed' if any(c['state'] == 'failed' for c in report['checks']) else 'passed'
    report['documentOperationsVerified'] = any(c['name'] == 'wps.smoke' and c['state'] == 'passed' for c in report['checks'])
    (output / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    for c in report['checks']:
        print(c['state'].upper() + ' ' + c['name'] + (': '+c['remedy'] if c['remedy'] else ''))
    print('Evidence: '+str(output / 'report.json'))
    return 0 if report['state'] == 'passed' else 1


if __name__ == '__main__':
    raise SystemExit(main())
