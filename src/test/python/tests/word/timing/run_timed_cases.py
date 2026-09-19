"""Opt-in: submit deployed requests once, retaining responses and timing separately.

Windows interactive desktop required. Supply fresh deployed requests/output names.
This collector does not implement per-case business assertions: run those against
retained responses separately; successful outcome alone is not a full test pass.
"""
import argparse
import csv
import ctypes
import hashlib
import json
import os
from pathlib import Path
import statistics
import subprocess
import sys
import time


def write_json(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--skill', type=Path, required=True)
    parser.add_argument('--requests', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    if os.name != 'nt':
        parser.error('Requires Windows interactive desktop')
    session = ctypes.c_ulong()
    if not ctypes.windll.kernel32.ProcessIdToSessionId(os.getpid(), ctypes.byref(session)) or session.value == 0:
        parser.error('Requires an interactive session, not SSH session 0')
    sources = sorted(args.requests.resolve().glob('*.json'))
    if not sources:
        parser.error('No JSON requests')
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    env = dict(os.environ, PYTHONIOENCODING='utf-8', WPS_TRACE_DIR=str(output/'traces'),
               WPS_SKILLS_TASK_DIR=str(output/'receipts'))
    env.pop('PYTHONPATH', None)
    python = str(Path(sys.executable).with_name('python.exe'))
    tasks, actions = [], []
    metadata = {'desktopSession': session.value, 'platform': sys.platform,
                'python': sys.version, 'skill': str(args.skill.resolve()),
                'actionBoundary': 'WordTask.execute entry to return/exception',
                'taskBoundary': 'before Popen through communicate/process exit',
                'businessAssertions': 'not evaluated by timing collector'}
    manifest = args.skill/'runtime/files.sha256.json'
    if manifest.is_file():
        metadata['packageManifestSha256'] = hashlib.sha256(manifest.read_bytes()).hexdigest()
    write_json(output/'timing-environment.json', metadata)
    try:
        for source in sources:
            case = source.stem
            request = json.loads(source.read_text(encoding='utf-8-sig'))
            write_json(output/(case+'.request.json'), request)
            measurement = output/(case+'.measurements.json')
            command = [python, str(Path(__file__).with_name('timed_entry.py')), '--skill',
                       str(args.skill.resolve()), '--request', str(source), '--measurements', str(measurement)]
            started = time.perf_counter_ns()
            proc = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    env=env, creationflags=subprocess.CREATE_NO_WINDOW)
            stdout, stderr = proc.communicate()
            ended = time.perf_counter_ns()
            (output/(case+'.stdout.json')).write_bytes(stdout)
            (output/(case+'.stderr.txt')).write_bytes(stderr)
            row = {'case': case, 'returnCode': proc.returncode, 'totalMs': (ended-started)/1e6,
                   'taskOutcome': None, 'actionSumMs': None, 'otherOverheadMs': None}
            tasks.append(row)
            result = json.loads(stdout.decode('utf-8-sig'))
            row.update(taskId=result.get('taskId'), taskOutcome=result.get('outcome'))
            measured = json.loads(measurement.read_text()) if measurement.is_file() else []
            by_trace = {v['traceId']: v for v in measured if v.get('traceId')}
            items = [result['document'], *result['steps'],
                     *[v for v in result['completion'].values() if v]]
            matched = set()
            for item in items:
                response = item.get('response') or {}
                trace_id = response.get('traceId')
                value = by_trace.get(trace_id)
                if value is not None: matched.add(value['sequence'])
                actions.append({'case': case, 'stepId': item['id'], 'action': item['address']['action'],
                                'state': item['state'], 'outcome': response.get('outcome'), 'traceId': trace_id,
                                'durationMs': value['durationNs']/1e6 if value else None,
                                'timingStatus': 'measured' if value else 'not_executed_or_missing'})
            for value in measured:
                if value['sequence'] not in matched:
                    actions.append({'case': case, 'stepId': None, 'action': value['action'],
                                    'state': 'unmatched', 'outcome': value['outcome'], 'traceId': value['traceId'],
                                    'durationMs': value['durationNs']/1e6, 'timingStatus': 'unmatched_measurement'})
            if measurement.is_file():
                row['actionSumMs'] = sum(v['durationNs'] for v in measured)/1e6
                row['otherOverheadMs'] = row['totalMs']-row['actionSumMs']
            write_json(output/'task-timings.json', tasks)
            write_json(output/'action-timings.json', actions)
            if result.get('outcome') == 'unknown':
                break
    finally:
        write_json(output/'task-timings.json', tasks)
        write_json(output/'action-timings.json', actions)
        summaries = []
        for name in sorted({a['action'] for a in actions}):
            selected = [a for a in actions if a['action'] == name]
            values = [a['durationMs'] for a in selected if a['outcome']=='succeeded' and a['durationMs'] is not None]
            summaries.append({'action': name, 'successfulTimedCount': len(values),
                              'failedOrMissingCount': len(selected)-len(values),
                              'meanMs': statistics.mean(values) if values else None,
                              'medianMs': statistics.median(values) if values else None,
                              'minMs': min(values) if values else None, 'maxMs': max(values) if values else None})
        write_json(output/'action-summary.json', summaries)
        for name, rows in [('task-timings',tasks),('action-timings',actions),('action-summary',summaries)]:
            if rows:
                fields = list(dict.fromkeys(k for r in rows for k in r))
                with (output/(name+'.csv')).open('w',encoding='utf-8-sig',newline='') as stream:
                    writer=csv.DictWriter(stream,fieldnames=fields);writer.writeheader();writer.writerows(rows)


if __name__ == '__main__':
    main()
