"""Word Task integration with the shared exact-document resource mechanisms.

The exhaustive binding matrix remains in Core; these assertions cover Task
ownership, native-result validation, interrupted acquisition and cleanup facts.
"""

import json
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch

from wps_skills.core.action_runtime import (ActionContract, ApplicationContractSet, ControllerResult,
                                           DefiniteEstablishFailure, ProcessCleanup,
                                           UnprovableEstablishFailure)
from wps_skills.core.trace_journal import JsonlTraceJournal
from wps_skills.word.task.executor import WordTask
from tests.core.resource_fixture import FakeApplicationAdapter, FakeDocumentCoordinator


class Adapter(FakeApplicationAdapter):
    application = 'word'


class WordTaskResourceTests(unittest.TestCase):
    def task(self, *, adapter=None, coordinator=None, **kwargs):
        contracts = ApplicationContractSet(application='word', allow_incomplete=True, contracts=[
            ActionContract(name='createDocument', binding_role='establish', risk='write',
                           parameters={'type': 'object'}, result={'type': 'object', 'required': ['created'],
                           'properties': {'created': {'type': 'boolean'}}}),
            ActionContract(name='inspectDocument', binding_role='required', risk='read',
                           parameters={'type': 'object'}, result={'type': 'object', 'required': ['text'],
                           'properties': {'text': {'type': 'string'}}}),
        ])
        task = WordTask(task_id='resource-task', contracts=contracts,
                        adapter=adapter or Adapter(document=object()),
                        coordinator=coordinator or FakeDocumentCoordinator(),
                        journal=kwargs.pop('journal', JsonlTraceJournal(None)), **kwargs)
        self.addCleanup(task.close)
        return task

    def call(self, task, action):
        return task.execute({'app': 'word', 'action': action}, {})

    def test_binding_is_committed_before_success_and_focus_does_not_retarget(self):
        document = object()
        adapter = Adapter(document=document)
        coordinator = FakeDocumentCoordinator()
        task = self.task(adapter=adapter, coordinator=coordinator)
        result = self.call(task, 'createDocument')
        self.assertEqual('succeeded', result['outcome'])
        self.assertEqual('commit', coordinator.calls[-1][0])
        adapter.active_document = object()
        self.assertEqual('succeeded', self.call(task, 'inspectDocument')['outcome'])
        self.assertIs(document, adapter.calls[-1][1])
        self.assertEqual('failed', self.call(task, 'createDocument')['outcome'])
        self.assertEqual(1, sum(call[0] == 'establish' for call in adapter.calls))
        self.assertEqual('succeeded', task.close()['outcome'])
        self.assertIs(document, coordinator.calls[-1][2])

    def test_partial_acquisition_and_lease_conflict_release_through_task(self):
        for error in (DefiniteEstablishFailure(code='DOCUMENT_LEASE_CONFLICT', message='owned'),
                      UnprovableEstablishFailure(code='DOCUMENT_BINDING_UNAVAILABLE', message='lost', outcome='unknown', partial_document=object())):
            with self.subTest(error=type(error).__name__):
                coordinator = FakeDocumentCoordinator()
                task = self.task(adapter=Adapter(establish_script=[error]), coordinator=coordinator)
                result = self.call(task, 'createDocument')
                self.assertNotEqual('succeeded', result['outcome'])
                self.assertEqual('succeeded', task.close()['outcome'])
                self.assertEqual('release', coordinator.calls[-1][0])

    def test_task_does_not_trust_invalid_controller_success(self):
        adapter = Adapter(document=object(), handler_script=[ControllerResult.succeeded(
            data={'text': 123}, controller_state='usable', binding_disposition='unchanged')])
        task = self.task(adapter=adapter)
        self.call(task, 'createDocument')
        result = self.call(task, 'inspectDocument')
        self.assertNotEqual('succeeded', result['outcome'])
        self.assertEqual('INVALID_RESULT', result['error']['code'])

    def test_closing_in_flight_acquisition_is_bounded_and_retains_uncertainty(self):
        entered, finish = threading.Event(), threading.Event()
        coordinator = FakeDocumentCoordinator(begin_entered=entered, begin_wait=finish,
                                               release_state='release_unconfirmed')
        task = self.task(coordinator=coordinator, cleanup_timeout_seconds=0.02)
        responses = []
        worker = threading.Thread(target=lambda: responses.append(self.call(task, 'createDocument')))
        worker.start()
        self.assertTrue(entered.wait(1))
        try:
            cleanup = task.close()
            self.assertEqual('failed', cleanup['outcome'])
            self.assertEqual('release_unconfirmed', cleanup['resources']['document_resources']['state'])
            self.assertEqual([True], coordinator.release_in_flight)
        finally:
            finish.set()
            worker.join(2)
        self.assertFalse(worker.is_alive())
        self.assertNotEqual('succeeded', responses[0]['outcome'])

    def test_action_deadline_stops_waiting_and_quarantines_in_flight_work(self):
        entered, finish = threading.Event(), threading.Event()
        adapter = Adapter(document=object())
        def blocked(document, command):
            entered.set()
            finish.wait(2)
            return ControllerResult.succeeded(data={"text": "late"}, controller_state="usable", binding_disposition="unchanged")
        adapter.handle = blocked
        coordinator = FakeDocumentCoordinator(release_state="release_unconfirmed")
        task = self.task(adapter=adapter, coordinator=coordinator, action_timeout_seconds=0.05)
        self.assertEqual("succeeded", self.call(task, "createDocument")["outcome"])
        try:
            response = self.call(task, "inspectDocument")
            self.assertTrue(entered.is_set())
            self.assertEqual("unknown", response["outcome"])
            self.assertEqual("TASK_ACTION_TIMEOUT", response["error"]["code"])
            self.assertFalse(task.can_execute)
            self.assertEqual("failed", task.close()["outcome"])
            self.assertEqual([True], coordinator.release_in_flight)
        finally:
            finish.set()

    def test_adapter_release_is_bounded_and_failure_is_not_success(self):
        finished = threading.Event()
        adapter = Adapter()
        adapter.close = lambda: finished.wait(1)
        task = self.task(adapter=adapter, cleanup_timeout_seconds=0.01)
        try:
            result = task.close()
            self.assertEqual("failed", result["outcome"])
            self.assertEqual("TASK_CLEANUP_INCOMPLETE", result["error"]["code"])
        finally:
            finished.set()

    def test_startup_failure_retains_cleanup_failure_in_the_task_receipt(self):
        import os
        from wps_skills.client.task_client import execute, exit_code
        for failure in (RuntimeError("assembly failed"), KeyboardInterrupt()):
            with self.subTest(failure=type(failure).__name__), tempfile.TemporaryDirectory() as directory:
                with patch.dict(os.environ, {"WPS_SKILLS_TASK_DIR": directory}), \
                     patch("wps_skills.windows.owned_process.WindowsOwnedProcessLauncher") as launcher, \
                     patch("wps_skills.word.actions.adapter.WordAdapter", side_effect=failure):
                    launcher.return_value.close.side_effect = OSError("release failed")
                    result = execute({"app": "word", "document": {"address": {"app": "word", "action": "createDocument"},
                                      "params": {}}, "steps": [], "completion": []}, "word",
                                     source_path=Path(directory) / "input.json")
                self.assertEqual("startup", result["stop"]["phase"])
                self.assertEqual("unknown", result["cleanup"]["outcome"])
                self.assertEqual("release failed", result["cleanup"]["error"])
                self.assertEqual(4, exit_code(result))
                launcher.return_value.close.assert_called_once()

    def test_launcher_cleanup_and_task_action_diagnostics_have_one_owner(self):
        class Launcher:
            count = 0
            def close(self):
                self.count += 1
                return (ProcessCleanup(pid=123, cleanup_steps=('already_exited',), released=True),)
        launcher = Launcher()
        with tempfile.TemporaryDirectory() as directory:
            task = self.task(launcher=launcher, journal=JsonlTraceJournal(Path(directory)))
            response = self.call(task, 'createDocument')
            task.close()
            task.close()
            self.assertEqual(1, launcher.count)
            rows = [json.loads(line) for p in Path(directory).rglob('*.jsonl') for line in p.read_text().splitlines()]
            self.assertTrue(rows)
            self.assertTrue(all(row['taskId'] == 'resource-task' for row in rows))
            self.assertNotIn('sessionId', json.dumps(rows))
            self.assertEqual('resource-task', response['taskId'])
