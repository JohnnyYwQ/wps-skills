from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock
import zipfile
from xml.etree import ElementTree as ET

from wps_skills.excel.demo import create_demo_workbook, run_demo


class ExcelDesktopDemoTests(unittest.TestCase):
    def test_noninteractive_environment_stops_before_creating_a_workbook(self):
        with tempfile.TemporaryDirectory() as root:
            output = Path(root) / 'demo'
            factory = Mock()
            with self.assertRaisesRegex(RuntimeError, 'desktop'):
                run_demo(output, desktop_check=Mock(side_effect=RuntimeError('desktop unavailable')), session_factory=factory)
            self.assertFalse(output.exists())
            factory.assert_not_called()

    def test_no_visible_window_means_no_edit_or_save(self):
        class Clock:
            ticks = 0
            def __call__(self):
                self.ticks += 20
                return self.ticks
        client = Mock()
        client.call.return_value = {'data': {}}
        client.session_outcome = {'outcome': 'succeeded'}
        factory = Mock()
        factory.return_value.__enter__ = Mock(return_value=client)
        factory.return_value.__exit__ = Mock(return_value=False)
        from unittest.mock import patch
        with tempfile.TemporaryDirectory() as root, patch('wps_skills.excel.demo.time.monotonic', new=Clock()):
            with self.assertRaisesRegex(RuntimeError, '可见窗口'):
                run_demo(Path(root) / 'demo', delay=0, desktop_check=lambda: 6,
                         session_factory=factory, window_probe=lambda _: None)
        self.assertEqual(['openWorkbook', 'getWorkbookInfo'], [call.args[0]['action'] for call in client.call.call_args_list])

    def test_display_template_has_one_readable_sheet_and_does_not_overwrite(self):
        with tempfile.TemporaryDirectory() as root:
            path = Path(root) / 'demo.xlsx'
            create_demo_workbook(path)
            with zipfile.ZipFile(path) as archive:
                workbook = ET.fromstring(archive.read('xl/workbook.xml'))
                self.assertEqual(['演示'], [s.attrib['name'] for s in workbook.findall('{*}sheets/{*}sheet')])
                sheet = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
                self.assertEqual('18', sheet.find('{*}cols/{*}col').attrib['width'])
                self.assertEqual('130', sheet.find('{*}sheetViews/{*}sheetView').attrib['zoomScale'])
            with self.assertRaises(FileExistsError):
                create_demo_workbook(path)


    def test_repeated_output_directory_names_get_distinct_workbook_names(self):
        from unittest.mock import patch
        client = Mock()
        client.call.return_value = {'data': {}}
        client.session_outcome = {'outcome': 'succeeded'}
        factory = Mock()
        factory.return_value.__enter__ = Mock(return_value=client)
        factory.return_value.__exit__ = Mock(return_value=False)
        with tempfile.TemporaryDirectory() as root, patch('wps_skills.excel.demo.time.monotonic', side_effect=range(0, 200, 20)):
            for parent in ('first', 'second'):
                with self.assertRaisesRegex(RuntimeError, '可见窗口'):
                    run_demo(Path(root) / parent / 'demo', delay=0, desktop_check=lambda: 6,
                             session_factory=factory, window_probe=lambda _: None)
        opened = [Path(call.args[1]['path']).name for call in client.call.call_args_list
                  if call.args[0]['action'] == 'openWorkbook']
        self.assertEqual(2, len(set(opened)))
