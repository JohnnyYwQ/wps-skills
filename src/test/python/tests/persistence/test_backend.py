import unittest
from wps_skills.core.action_session import ActionAddress, ControllerCommand, ControllerContext
from wps_skills.excel.adapter import ExcelAdapter
from wps_skills.excel.contracts import EXCEL_TARGET_CONTRACT_SET
from wps_skills.excel.windows.backend import WindowsExcelBackend
from wps_skills.ppt.adapter import PptAdapter
from wps_skills.ppt.contracts import PPT_TARGET_CONTRACT_SET
from wps_skills.ppt.windows.backend import WindowsPptBackend
from wps_skills.windows.bridge_types import BackendActionFailure, WindowsDocument
from wps_skills.word.adapter import WordBackendOperation
from wps_skills.word.windows.backend import WindowsWordBackend
from tests.word.windows.test_backend import RecordingBridge


class PersistenceBackendTests(unittest.TestCase):
    def test_create_establishes_once_with_its_own_preparation(self):
        for app,action,adapter_type,backend_type,contracts in (
            ('excel','createWorkbook',ExcelAdapter,WindowsExcelBackend,EXCEL_TARGET_CONTRACT_SET),
            ('ppt','createPresentation',PptAdapter,WindowsPptBackend,PPT_TARGET_CONTRACT_SET)):
            bridge=RecordingBridge({'acquire_new_document':{'documentId':'new-id','documentState':{'persistenceState':'unsaved','readOnly':False}}})
            backend=backend_type(bridge=bridge);adapter=adapter_type(backend=backend,contracts=contracts)
            command=ControllerCommand(ActionAddress(app,action),{},ControllerContext('r','t',None))
            prepared=adapter.prepare_establish(command);acquired=adapter.establish(prepared,command)
            self.assertIsNone(acquired.document.authorized_path)
            self.assertEqual('unsaved',acquired.data['documentState']['persistenceState'])
            with self.assertRaises(ValueError):adapter.establish(prepared,command)
            with self.assertRaises(ValueError):backend.is_live(WindowsDocument('new-id',None))
            self.assertEqual(['prepare_new_document','acquire_new_document'],[c[0] for c in bridge.calls])

    def test_word_save_as_keeps_exact_reference_and_updates_private_save_locator(self):
        bridge=RecordingBridge({'acquire_new_document':{'documentId':'new-id','revision':'r1','persistenceState':'unsaved','readOnly':False},
                                'save_as_artifact':{'artifact':{'path':'C:/new.docx'}},'save_existing_artifact':{}})
        backend=WindowsWordBackend(bridge=bridge)
        document=backend.create_document(backend.prepare_create_document(None),None).document
        backend.invoke(document,WordBackendOperation('save_as_artifact',{'outputPath':'C:/new.docx'}),None)
        backend.invoke(document,WordBackendOperation('save_existing_artifact',{}),None)
        self.assertEqual('C:/new.docx',bridge.calls[-1][1]['authorizedPath'])
        self.assertIsNone(document.authorized_path)
        bridge.replies['save_as_artifact']=BackendActionFailure(outcome='failed',code='OUTPUT_ALREADY_EXISTS',message='exists',binding_disposition='unchanged')
        with self.assertRaises(BackendActionFailure):backend.invoke(document,WordBackendOperation('save_as_artifact',{'outputPath':'C:/exists.docx'}),None)
        backend.invoke(document,WordBackendOperation('save_existing_artifact',{}),None)
        self.assertEqual('C:/new.docx',bridge.calls[-1][1]['authorizedPath'])
        self.assertEqual(2,sum(c[0]=='save_as_artifact' for c in bridge.calls))
