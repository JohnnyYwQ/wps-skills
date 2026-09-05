import unittest

from wps_skills.core.action_session import ControllerContext, ControllerCommand, ActionAddress, DefiniteEstablishFailure
from wps_skills.excel.adapter import ExcelAdapter
from wps_skills.excel.contracts import EXCEL_TARGET_CONTRACT_SET
from wps_skills.windows.bridge_types import BackendActionFailure, WindowsDocument
from wps_skills.excel.windows.backend import WindowsExcelBackend
from wps_skills.windows.document_coordinator import WindowsDocumentCoordinator


class Bridge:
    def __init__(self):
        self.calls = []
        self.replies = {
            'prepare_existing_document': {'preparationId': 'p1', 'coordinationIdentity': 'file-1'},
            'acquire_existing_document': {'documentId': 'd1', 'artifact': {'path': r'C:\book.xlsx', 'format': 'xlsx', 'sizeBytes': 2000},
                                          'documentState': {'persistenceState': 'saved', 'readOnly': False}},
            'probe_bound_document': {'live': True},
            'acquire_coordination_guard': {'guardId': 'g1'},
            'commit_document_lease': {'leaseId': 'l1'},
            'release_document_resources': {'state': 'released'},
        }

    def execute(self, operation, arguments, context):
        self.calls.append((operation, arguments))
        result = self.replies[operation]
        if isinstance(result, Exception):
            raise result
        return result

    def close(self):
        return True


class ExcelAdapterTests(unittest.TestCase):
    def setUp(self):
        self.bridge = Bridge()
        self.backend = WindowsExcelBackend(bridge=self.bridge)
        self.adapter = ExcelAdapter(backend=self.backend, contracts=EXCEL_TARGET_CONTRACT_SET)
        self.command = ControllerCommand(ActionAddress('excel', 'openWorkbook'), {'path': r'C:\book.xlsx'}, ControllerContext('r1', 't1', None))

    def bind(self):
        return self.adapter.establish(self.adapter.prepare_establish(self.command), self.command)

    def test_existing_workbook_uses_shared_coordination_and_exact_reference(self):
        acquired = self.bind()
        coordinator = WindowsDocumentCoordinator(bridge=self.bridge)
        guard = coordinator.begin('file-1', self.command.context)
        lease = coordinator.commit(guard, acquired.document, self.command.context)
        self.assertTrue(self.adapter.is_live(acquired.document))
        self.assertEqual('released', coordinator.release(guard, acquired.document, lease).state)
        with self.assertRaises(ValueError):
            self.adapter.is_live(WindowsDocument('d1', r'C:\book.xlsx'))

    def test_another_preparation_or_application_cannot_be_used(self):
        prepared = self.adapter.prepare_establish(self.command)
        wrong = ControllerCommand(ActionAddress('word', 'openWorkbook'), self.command.params, self.command.context)
        with self.assertRaises(ValueError):
            self.adapter.establish(prepared, wrong)
        wrong = ControllerCommand(self.command.address, {'path': r'C:\other.xlsx'}, self.command.context)
        with self.assertRaises(ValueError):
            self.adapter.establish(prepared, wrong)

    def test_missing_file_is_a_definite_establish_failure(self):
        self.bridge.replies['prepare_existing_document'] = BackendActionFailure(outcome='failed', code='DOCUMENT_NOT_FOUND', message='missing', binding_disposition='unchanged')
        with self.assertRaises(DefiniteEstablishFailure):
            self.adapter.prepare_establish(self.command)
        self.assertEqual(1, len(self.bridge.calls))

    def test_unknown_mutation_is_reported_without_retry(self):
        acquired = self.bind()
        self.bridge.replies['write_literal_rectangle'] = BackendActionFailure(outcome='unknown', code='RANGE_WRITE_FAILED', message='partial write', binding_disposition='unchanged')
        command = ControllerCommand(ActionAddress('excel', 'writeRange'), {}, self.command.context)
        result = self.adapter.handle(acquired.document, command)
        self.assertEqual('unknown', result.outcome)
        self.assertEqual('usable', result.controller_state)
        self.assertEqual(1, sum(op == 'write_literal_rectangle' for op, _ in self.bridge.calls))

    def test_missing_handler_fails_assembly(self):
        with self.assertRaises(ValueError):
            ExcelAdapter(backend=self.backend, contracts=EXCEL_TARGET_CONTRACT_SET, handlers={})
