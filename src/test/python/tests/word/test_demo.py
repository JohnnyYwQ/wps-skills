from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock
import zipfile
from xml.etree import ElementTree as ET

from wps_skills.word.demo import create_demo_document, run_demo


class WordDemoTests(unittest.TestCase):
    def test_noninteractive_environment_stops_before_creating_a_document(self):
        with tempfile.TemporaryDirectory() as root:
            output = Path(root) / 'demo'
            factory = Mock()
            with self.assertRaisesRegex(RuntimeError, 'desktop'):
                run_demo(output, desktop_check=Mock(side_effect=RuntimeError('desktop unavailable')), task_entry=factory)
            self.assertFalse(output.exists())
            factory.assert_not_called()

    def test_invalid_delay_stops_before_creating_output(self):
        with tempfile.TemporaryDirectory() as root:
            output = Path(root) / 'demo'
            for delay in (-1, 11, float('nan')):
                with self.subTest(delay=delay), self.assertRaises(ValueError):
                    run_demo(output, desktop_check=lambda: 1, delay=delay)
                self.assertFalse(output.exists())

    def test_blank_fixture_contains_no_content_and_refuses_overwrite(self):
        with tempfile.TemporaryDirectory() as root:
            path = Path(root) / 'blank.docx'
            create_demo_document(path)
            original = path.read_bytes()
            with zipfile.ZipFile(path) as archive:
                document = ET.fromstring(archive.read('word/document.xml'))
                self.assertFalse(document.findall('.//{*}t'))
                self.assertFalse(document.findall('.//{*}tbl'))
            with self.assertRaises(FileExistsError):
                create_demo_document(path)
            self.assertEqual(original, path.read_bytes())
