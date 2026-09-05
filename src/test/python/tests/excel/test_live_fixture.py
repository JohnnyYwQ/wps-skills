from pathlib import Path
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET
from tests.excel.live_acceptance import create_fixture, HOST, FENCE_PROBE


class ExcelAcceptanceFixtureTests(unittest.TestCase):
    def test_acceptance_dependencies_and_disposable_fixture_are_available(self):
        self.assertTrue(HOST.is_file())
        self.assertTrue(FENCE_PROBE.is_file())
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'book.xlsx'
            create_fixture(path)
            with zipfile.ZipFile(path) as archive:
                self.assertIsNone(archive.testzip())
                workbook = ET.fromstring(archive.read('xl/workbook.xml'))
                names = [node.attrib['name'] for node in workbook.findall('{*}sheets/{*}sheet')]
                self.assertEqual(['Sheet1', '数据'], names)
            with self.assertRaises(FileExistsError):
                create_fixture(path)
