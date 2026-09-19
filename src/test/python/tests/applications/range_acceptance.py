"""Excel range regressions: direct Task JSON, no Agent; interactive Windows only.

Prepare --root with an isolated wps-excel package and range-input-<root.name>.xlsx
(sheet 工单). Uses fresh requests/output paths and the default 60s Action timeout.
Timing traces and full Task/Action responses are retained for comparison.
"""
import argparse
import ctypes
import json
import os
from pathlib import Path
import subprocess
import sys
import time

from tests.applications.native_acceptance import Plan, ref, excel_plan


def run(root):
    session = ctypes.c_ulong()
    ctypes.windll.kernel32.ProcessIdToSessionId(os.getpid(), ctypes.byref(session))
    if not session.value:
        raise RuntimeError('Run in an interactive Windows session')
    root = root.resolve()
    for name in ('requests', 'outputs'):
        (root / name).mkdir(exist_ok=False)
    env = dict(os.environ, PYTHONUTF8='1', PYTHONIOENCODING='utf-8',
               WPS_SKILLS_TASK_DIR=str(root/'receipts'), WPS_TRACE_DIR=str(root/'traces'))
    cli = root/'wps-excel/scripts/excel.py'
    report = []

    def submit(name, request):
        path = root/'requests'/(name+'.json')
        path.write_text(json.dumps(request, ensure_ascii=False), encoding='utf-8')
        start = time.perf_counter()
        cp = subprocess.run([sys.executable, str(cli), '--app', 'excel', '--task-file', str(path)],
                            env=env, capture_output=True, text=True, encoding='utf-8', timeout=900)
        (root/(name+'.stdout')).write_text(cp.stdout, encoding='utf-8')
        (root/(name+'.stderr')).write_text(cp.stderr, encoding='utf-8')
        records = [json.loads(line) for line in cp.stdout.splitlines() if line.startswith('{')]
        response = next((v for v in reversed(records) if v.get('type') == 'task.response'), None)
        passed = response is not None and response['outcome'] == 'succeeded'
        report.append(dict(name=name, wallSeconds=time.perf_counter()-start, passed=passed, response=response))
        (root/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        assert passed, report[-1]
        assert response['cleanup']['outcome'] == 'succeeded', response
        print(name, round(report[-1]['wallSeconds'], 3), flush=True)

    for address in ('A1:D13', 'A1:H30', 'A1:P60', 'A1:T50'):
        p = Plan('excel', root/('range-input-'+root.name+'.xlsx'))
        p.add('readRange', sheet='工单', address=address)
        submit('read-'+address.replace(':', '-'), p.finish())
    for name, address in (('exact-merge','A1:C3'), ('contains-merge','A1:D4')):
        p = Plan('excel')
        sheet = ref(p.add('listWorksheets', offset=0, limit=100), 'worksheets', 0, 'name')
        def edit(action, area, **params):
            token = p.add('readRange', sheet=sheet, address=area)
            return p.add(action, sheet=sheet, address=area, expectedToken=ref(token, 'token'), **params)
        edit('writeRange', 'A1', values=[['merged heading']])
        edit('mergeRange', 'A1:C3')
        edit('formatRange', address, format={'bold': True, 'italic': True, 'fontSize': 14,
             'fontColor': 255, 'fillColor': 65535, 'wrapText': True,
             'horizontalAlignment': 'center', 'verticalAlignment': 'center', 'numberFormat': '0.00'})
        submit(name, p.finish(root/'outputs'/(name+'-'+root.name+'.xlsx')))
    submit('excel-all-actions', excel_plan(root/'outputs'))
    return 0


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', required=True, type=Path)
    raise SystemExit(run(parser.parse_args().root))
