import json
from pathlib import Path
import tempfile
import unittest

from wps_skills.core.trace_journal import JsonlTraceJournal


class JsonlTraceJournalTests(unittest.TestCase):
    def test_writes_separate_task_and_action_timing_journals(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            journal = JsonlTraceJournal(
                Path(temporary_directory),
                timestamp_factory=lambda: "2026-09-04T08:00:00.000Z",
                trace_id_factory=lambda: "trace-1",
            )

            task_path = Path(temporary_directory) / "tasks/task-1.jsonl"
            journal.task_event(task_id="task-1", event="task.resources_acquired")
            trace = journal.task_action_trace(task_id="task-1")
            trace.event(
                "action.finished",
                elapsedMs=1250,
                address={"app": "word", "action": "inspectDocument"},
            )

            task_rows = [
                json.loads(line)
                for line in task_path.read_text(encoding="utf-8").splitlines()
            ]
            action_rows = [
                json.loads(line)
                for line in trace.trace_log.read_text(encoding="utf-8").splitlines()
            ]

        self.assertEqual("task.resources_acquired", task_rows[0]["event"])
        self.assertEqual("2026-09-04T08:00:00.000Z", task_rows[0]["ts"])
        self.assertEqual("trace-1", trace.trace_id)
        self.assertEqual("action.finished", action_rows[0]["event"])
        self.assertEqual(1250, action_rows[0]["elapsedMs"])
        self.assertNotIn("params", action_rows[0])

    def test_application_labels_follow_the_task_and_request(self):
        from wps_skills.core import timing
        with tempfile.TemporaryDirectory() as temporary:
            journal = JsonlTraceJournal(Path(temporary))
            for app in ('word', 'excel', 'ppt'):
                journal.task_event(task_id=app, application=app, event='test')
                journal.task_action_trace(task_id=app, application=app).event('test')
                with timing.submission(journal=journal, application=app):
                    timing.bind_task(app)
            for folder in ('tasks', 'actions', 'requests'):
                apps = set()
                for path in (Path(temporary) / folder).glob('*.jsonl'):
                    for line in path.read_text().splitlines():
                        row = json.loads(line)
                        if row.get('taskId'):
                            self.assertEqual(row['taskId'], row['app'])
                            apps.add(row['app'])
                self.assertEqual({'word', 'excel', 'ppt'}, apps)

    def test_unavailable_journal_keeps_stable_null_trace_path(self):
        journal = JsonlTraceJournal(None, trace_id_factory=lambda: "trace-1")

        trace = journal.task_action_trace(task_id="task-1")
        trace.event("action.finished", elapsedMs=1)

        self.assertEqual("trace-1", trace.trace_id)
        self.assertIsNone(trace.trace_log)


if __name__ == "__main__":
    unittest.main()
