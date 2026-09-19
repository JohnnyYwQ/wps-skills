import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[5]
sys.path.insert(0, str(ROOT / 'src/test/python/diagnostics/startup_communication'))
import doctor


class DoctorTests(unittest.TestCase):
    def test_registration_is_not_document_verification(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(doctor, 'inspect', return_value=[doctor.check('com.registration', True, 'KET.Application')]), patch.object(doctor, 'smoke') as smoke:
            path = Path(tmp)/'run'
            self.assertEqual(0, doctor.main(['--app','excel','--output',str(path)]))
            self.assertFalse(json.loads((path/'report.json').read_text())['documentOperationsVerified'])
            smoke.assert_not_called()

    def test_failed_prerequisites_prevent_document_operations(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(doctor, 'inspect', return_value=[doctor.check('desktop.session', False, 0)]), patch.object(doctor, 'smoke') as smoke:
            self.assertEqual(1, doctor.main(['--app','excel','--smoke-test','--output',str(Path(tmp)/'run')]))
            smoke.assert_not_called()

    def test_session_zero_and_corrupted_unicode_are_separate_failures(self):
        checks=doctor.probe_checks(dict(edition='Desktop',powershellVersion='5.1',sessionId=0,userInteractive=False,registered=True,echo='???'),'中文')
        failed={c['name'] for c in checks if c['state']=='failed'}
        self.assertEqual({'desktop.session','communication.utf8'},failed)

    def test_existing_directory_is_never_reused(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(doctor,'inspect') as inspect:
            self.assertEqual(1,doctor.main(['--app','excel','--output',tmp]))
            inspect.assert_not_called()

    def test_unsupported_smoke_does_not_claim_success(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(doctor,'inspect',return_value=[]), patch.object(doctor,'smoke') as smoke:
            self.assertEqual(1,doctor.main(['--app','word','--smoke-test','--output',str(Path(tmp)/'run')]))
            smoke.assert_not_called()
