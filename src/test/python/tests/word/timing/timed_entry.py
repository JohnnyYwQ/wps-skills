"""Test-only timing wrapper around the unmodified packaged Word CLI."""
import argparse
import json
from pathlib import Path
import runpy
import sys
import time


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--skill', type=Path, required=True)
    parser.add_argument('--request', type=Path, required=True)
    parser.add_argument('--measurements', type=Path, required=True)
    args = parser.parse_args()
    sys.path.insert(0, str(args.skill / 'runtime/src/main/python'))
    from wps_skills.word.task.executor import WordTask
    original = WordTask.execute
    measurements = []

    def measured(self, address, params):
        row = {'sequence': len(measurements), 'action': address['action'],
               'taskId': self.task_id, 'outcome': None, 'traceId': None}
        started = time.perf_counter_ns()
        try:
            response = original(self, address, params)
        except BaseException as exc:
            ended = time.perf_counter_ns()
            row.update(outcome='exception', exceptionType=type(exc).__name__)
            raise
        else:
            ended = time.perf_counter_ns()
            row.update(outcome=response.get('outcome'), traceId=response.get('traceId'))
            return response
        finally:
            row['durationNs'] = ended - started
            measurements.append(row)

    WordTask.execute = measured
    sys.argv = [str(args.skill / 'scripts/word.py'), '--app', 'word',
                '--task-file', str(args.request)]
    try:
        runpy.run_path(sys.argv[0], run_name='__main__')
    finally:
        WordTask.execute = original
        args.measurements.write_text(json.dumps(measurements, indent=2), encoding='utf-8')


if __name__ == '__main__':
    main()
