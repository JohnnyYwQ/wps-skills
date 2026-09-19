import io
import json
from pathlib import Path
import tempfile
import threading
import unittest
from contextvars import copy_context
from unittest.mock import patch

from wps_skills.core import timing
from wps_skills.core.trace_journal import JsonlTraceJournal
from wps_skills.cli.call import main


class TimingTests(unittest.TestCase):
    def test_nested_spans_use_monotonic_duration_and_keep_thread_identity(self):
        with tempfile.TemporaryDirectory() as directory:
            journal = JsonlTraceJournal(Path(directory), timestamp_factory=lambda: 'fixed-wall-clock')
            ticks = iter([0, 10, 20, 30, 40, 50])
            with timing.submission(journal=journal, clock=lambda: next(ticks)):
                timing.bind_task('task-one')
                with timing.fields(stepId='write', action='writeContent'):
                    with timing.span('action.execute'):
                        context = copy_context()
                        def native():
                            with timing.span('bridge.round_trip'):
                                pass
                        worker = threading.Thread(target=lambda: context.run(native))
                        worker.start(); worker.join()
            rows = [json.loads(x) for x in next((Path(directory)/'requests').glob('*.jsonl')).read_text().splitlines()]
        finished = {r['name']:r for r in rows if r['event']=='span.finished'}
        self.assertEqual(50,finished['request.total']['durationNs'])
        self.assertEqual(10,finished['bridge.round_trip']['durationNs'])
        self.assertEqual(finished['action.execute']['spanId'],finished['bridge.round_trip']['parentSpanId'])
        self.assertEqual('write',finished['bridge.round_trip']['stepId'])
        self.assertEqual('task-one',finished['bridge.round_trip']['taskId'])

    def test_diagnostic_failure_does_not_mask_business_exception(self):
        class Broken:
            def request_event(self, **fields): raise OSError('disk unavailable')
        with self.assertRaisesRegex(ValueError, 'business failure'):
            with timing.submission(journal=Broken()):
                with timing.span('action.execute'):
                    raise ValueError('business failure')
        self.assertIsNone(timing._current.get())

    def test_rejected_json_is_traced_through_response_flush_without_payload(self):
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory); request=root/'bad.json';request.write_text('{secret-invalid')
            output=io.StringIO()
            with patch.dict('os.environ', {'WPS_TRACE_DIR':str(root/'logs')}):
                code=main(['--app','word','--task-file',str(request)],output_stream=output,error_stream=io.StringIO())
            self.assertNotEqual(0,code)
            self.assertEqual('rejected',json.loads(output.getvalue())['state'])
            text=next((root/'logs/requests').glob('*.jsonl')).read_text()
            rows=[json.loads(x) for x in text.splitlines()]
            names={r.get('name') for r in rows}
            self.assertTrue({'request.total','input.read','input.decode','response.encode','response.write_flush'}<=names)
            decode=next(r for r in rows if r.get('name')=='input.decode' and r['event']=='span.finished')
            self.assertEqual('exception',decode['status'])
            self.assertNotIn('secret-invalid',text)

    def test_broken_output_records_terminal_failure(self):
        class BrokenOutput(io.StringIO):
            def write(self, value): raise BrokenPipeError('caller disconnected')
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory); source=root/'bad.json';source.write_text('{')
            with patch.dict('os.environ',{'WPS_TRACE_DIR':str(root/'logs')}):
                with self.assertRaises(BrokenPipeError):
                    main(['--app','word','--task-file',str(source)],output_stream=BrokenOutput(),error_stream=io.StringIO())
            rows=[json.loads(x) for x in next((root/'logs/requests').glob('*.jsonl')).read_text().splitlines()]
            terminal=[r for r in rows if r.get('name') in ('response.write_flush','request.total') and r['event']=='span.finished']
            self.assertEqual(2,len(terminal))
            self.assertTrue(all(r['status']=='exception' for r in terminal))
