"""Test-only fault switches around the real standalone Task CLI."""
import argparse
import copy
import importlib
import json
import os
from pathlib import Path
import sys
import time


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--skill', required=True, type=Path)
    parser.add_argument('--config', required=True, type=Path)
    parser.add_argument('cli', nargs=argparse.REMAINDER)
    args = parser.parse_args()
    config = json.loads(args.config.read_text(encoding='utf-8-sig'))
    directory = Path(config['caseDir'])
    sys.path.insert(0, str(args.skill / 'runtime/src/main/python'))
    os.environ['WPS_DESIGN_SPEC'] = str(args.config)
    if config.get('noWrite'):
        os.environ['WPS_DESIGN_NO_WRITE'] = '1'
    if config.get('optimisticReadback'):
        os.environ['WPS_DESIGN_OPTIMISTIC'] = '1'

    def event(kind, **fields):
        with (directory / 'child-events.jsonl').open('a', encoding='utf-8') as f:
            f.write(json.dumps(dict(event=kind, pid=os.getpid(), monotonic=time.perf_counter(), **fields), ensure_ascii=False) + '\n')

    def pause(phase, address):
        if config.get('pausePhase') == phase and config.get('action') == address['action']:
            (directory / 'ready.json').write_text(json.dumps({'phase': phase, 'pid': os.getpid(), 'address': address}))
            deadline = time.monotonic() + 45
            while not (directory / 'release').exists():
                if time.monotonic() > deadline:
                    raise TimeoutError('Test controller did not release the stage barrier')
                time.sleep(.02)

    from wps_skills.client import task_client
    if config.get('deferredPreflight'):
        # Controlled ablation: request structure, per-step validation and all native behavior remain.
        task_client._preflight = lambda *unused: None
        event('candidate', change='omit complete-plan static preflight')

    if config.get('finalReceiptOnly'):
        from wps_skills.client import task_store,task_request
        original_publish=task_store.publish
        admitted=set()
        def publish(path,result):
            if result['state'] not in ('running','closing'):
                return original_publish(path,result)
            if str(path) not in admitted:
                # A fair final-only candidate must conservatively admit that every step may run.
                conservative=copy.deepcopy(result)
                for step in task_request.result_steps(conservative):step['state']='unknown'
                original_publish(path,conservative)
                admitted.add(str(path))
                event('candidate',change='admission plus final receipt only; in-progress steps conservatively unknown')
            else:event('receipt_suppressed',state=result['state'])
        task_store.publish=publish

    module = importlib.import_module('wps_skills.' + config['app'] + '.windows.task_factory')
    original = module.build_task

    class Observed:
        def __init__(self, inner):
            self.inner = inner

        @property
        def can_execute(self):
            return self.inner.can_execute

        def execute(self, address, params):
            event('before', address=address)
            pause('before', address)
            response = self.inner.execute(address, params)
            event('after', address=address, response=response)
            pause('after', address)
            return response

        def close(self):
            result = self.inner.close()
            event('actual_cleanup', result=result)
            if config.get('cleanupFailure'):
                # Failure-report aggregation test; actual resources are still cleaned safely.
                return dict(result, outcome='failed', error={'code': 'TEST_CLEANUP_FAILURE', 'message': 'Injected cleanup report failure after real cleanup'})
            return result

    def factory(**kwargs):
        event('executor_created', taskId=kwargs['task_id'])
        return Observed(original(**kwargs))

    module.build_task = factory
    from wps_skills.cli.task import main as task_main
    cli = args.cli[1:] if args.cli[:1] == ['--'] else args.cli
    return task_main(cli, required_application=config['app'])


if __name__ == '__main__':
    raise SystemExit(main())
