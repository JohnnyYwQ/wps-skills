"""Verify same-basename WPS rejection without opening a blocking dialog."""
import argparse
import json
from pathlib import Path
import shutil
import uuid

from wps_skills.client.session_client import ActionFailed
from wps_skills.excel.demo import create_demo_workbook
from wps_skills.excel.skill import open_session


def run(output):
    output.mkdir(parents=True, exist_ok=False)
    first = output / 'first'
    second = output / 'second'
    first.mkdir()
    second.mkdir()
    name = 'name-conflict-' + uuid.uuid4().hex + '.xlsx'
    original = first / name
    conflict = second / name
    distinct = second / ('distinct-' + name)
    create_demo_workbook(original)
    shutil.copyfile(original, conflict)
    shutil.copyfile(original, distinct)
    report = {'status': 'running'}
    try:
        with open_session(timeout=30) as owner:
            owner.call({'app': 'excel', 'action': 'openWorkbook'}, {'path': str(original)})
            with open_session(timeout=30) as caller:
                try:
                    caller.call({'app': 'excel', 'action': 'openWorkbook'}, {'path': str(conflict)})
                except ActionFailed as exc:
                    assert exc.response['outcome'] == 'failed', exc.response
                    assert exc.response['error']['code'] == 'DOCUMENT_OPEN_FAILED', exc.response
                    report['rejection'] = exc.response
                else:
                    raise AssertionError('Expected a definite same-basename rejection')
                # A pre-open rejection must release its guard and leave the Session unbound/usable.
                report['subsequentOpen'] = caller.call({'app': 'excel', 'action': 'openWorkbook'}, {'path': str(distinct)})
                report['info'] = caller.call({'app': 'excel', 'action': 'getWorkbookInfo'}, {})
            assert caller.session_outcome['outcome'] == 'succeeded', caller.session_outcome
        assert owner.session_outcome['outcome'] == 'succeeded', owner.session_outcome
        report['status'] = 'passed'
    except BaseException as exc:
        report.update(status='failed', error=repr(exc))
        raise
    finally:
        (output/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(output/'report.json')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--output-dir', type=Path, required=True)
    args = parser.parse_args()
    run(args.output_dir.resolve())
