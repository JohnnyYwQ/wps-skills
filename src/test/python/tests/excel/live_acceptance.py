"""Explicit Windows-only acceptance against newly generated disposable workbooks."""
import argparse
import json
import os
from pathlib import Path
import platform
import queue
import subprocess
import threading
import sys
import time
import zipfile
from xml.etree import ElementTree as ET

from wps_skills.client.session_client import ActionFailed, SessionClient

REPO = Path(__file__).resolve().parents[5]
HOST = Path(__file__).with_name('session_host_fixture.py')
FENCE_PROBE = REPO / 'src/test/resources/wps_skills/excel/coordination_fence_probe.ps1'


def create_fixture(path):
    """Small OOXML fixture, independent of COM and third-party Python packages."""
    files = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="数据" sheetId="2" r:id="rId2"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml': '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/><mergeCells count="1"><mergeCell ref="A10:B10"/></mergeCells></worksheet>',
        'xl/worksheets/sheet2.xml': '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
    }
    with zipfile.ZipFile(path, 'x', zipfile.ZIP_DEFLATED) as archive:
        for name, xml in files.items():
            archive.writestr(name, xml.encode('utf-8'))



def verify_replacement_quarantine(output):
    from wps_skills.windows.bridge_runtime import windows_powershell_executable
    fixture = output / 'quarantine-probe.bin'
    fixture.write_text('original', encoding='utf-8')
    command = [windows_powershell_executable(), '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
               '-File', str(FENCE_PROBE),
               '-CommonPath', str(REPO / 'src/main/resources/wps_skills/windows/bridge_common.ps1'),
               '-File', str(fixture)]
    process = subprocess.Popen(command + ['-Mode', 'hold'], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                               stderr=subprocess.PIPE, text=True, creationflags=subprocess.CREATE_NO_WINDOW)
    ready = queue.Queue()
    reader = threading.Thread(target=lambda: ready.put(process.stdout.readline()), daemon=True)
    reader.start()
    try:
        assert ready.get(timeout=15).strip() == 'READY', 'coordination fixture did not establish its fence'
    finally:
        process.kill()
        process.communicate(timeout=10)
        reader.join(timeout=2)
    result = subprocess.run(command + ['-Mode', 'probe'], capture_output=True, text=True, timeout=15,
                            creationflags=subprocess.CREATE_NO_WINDOW)
    assert result.returncode == 0 and result.stdout.strip() == 'QUARANTINED', result.stderr + result.stdout


def run(output, wps_version, *, candidate=False):
    if os.name != 'nt':
        raise RuntimeError('Live acceptance requires Windows and registered WPS Spreadsheets')
    output.mkdir(parents=True, exist_ok=False)
    workbook = output / 'excel-acceptance.xlsx'
    create_fixture(workbook)
    alias = output / 'hardlink-alias.xlsx'
    os.link(workbook, alias)
    report = {'status': 'running', 'platform': platform.platform(), 'python': sys.version, 'wpsVersionReported': wps_version,
              'workbook': str(workbook), 'candidateContracts': candidate, 'steps': []}
    env = dict(os.environ, PYTHONPATH=str(REPO / 'src/main/python'), PYTHONIOENCODING='utf-8')
    def client():
        command = [sys.executable, str(HOST)] if candidate else [sys.executable, '-m', 'wps_skills.cli.call', '--session', '--app', 'excel']
        return SessionClient(command, application='excel', env=env, timeout=90)
    def call(c, name, params):
        response = c.call({'app': 'excel', 'action': name}, params)
        report['steps'].append({'action': name, 'response': response})
        (output / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        return response['data']
    def rejected(c, name, params, code):
        try:
            call(c, name, params)
        except ActionFailed as exc:
            report['steps'].append({'action': name, 'expectedError': code, 'response': exc.response})
            if exc.response['error']['code'] != code:
                raise AssertionError(exc.response)
        else:
            raise AssertionError('expected rejection: ' + code)
    primary = client()
    try:
        with primary as c:
            rejected(c, 'readRange', {'sheet': 'Sheet1', 'address': 'A1'}, 'SESSION_DOCUMENT_NOT_BOUND')
            rejected(c, 'openWorkbook', {'path': str(output / 'missing.xlsx')}, 'DOCUMENT_NOT_FOUND')
            opened = call(c, 'openWorkbook', {'path': str(workbook)})
            assert opened['documentState']['persistenceState'] == 'saved', opened
            info = call(c, 'getWorkbookInfo', {})
            assert info['worksheetCount'] == 2, info
            page = call(c, 'listWorksheets', {'offset': 0, 'limit': 1})
            assert page['nextOffset'] == 1, page
            page2 = call(c, 'listWorksheets', {'offset': 1, 'limit': 1})
            assert page2['worksheets'][0]['name'] == '数据', page2
            with client() as competing:
                rejected(competing, 'openWorkbook', {'path': str(workbook)}, 'DOCUMENT_LEASE_CONFLICT')
                rejected(competing, 'openWorkbook', {'path': str(alias)}, 'DOCUMENT_LEASE_CONFLICT')
            rejected(c, 'openWorkbook', {'path': str(workbook)}, 'SESSION_DOCUMENT_ALREADY_BOUND')
            rejected(c, 'readRange', {'sheet': 'missing', 'address': 'A1'}, 'WORKSHEET_NOT_FOUND')
            merged_region = {'sheet': 'Sheet1', 'address': 'A10:B10'}
            merged = call(c, 'readRange', merged_region)
            rejected(c, 'writeRange', dict(merged_region, expectedToken=merged['token'], values=[[1, 2]]), 'RANGE_UNSUPPORTED')
            region = {'sheet': 'Sheet1', 'address': 'A1:C2'}
            before = call(c, 'readRange', region)
            values = [['中文', 12.5, True], ['=literal', None, -2146826281]]
            after = call(c, 'writeRange', dict(region, expectedToken=before['token'], values=values))
            assert [[x['value'] for x in row] for row in after['cells']] == values, after
            rejected(c, 'writeRange', dict(region, expectedToken=before['token'], values=values), 'STALE_RANGE')
            rejected(c, 'writeRange', dict(region, expectedToken=after['token'], values=[[1]]), 'INVALID_PARAMS')
            formatted = call(c, 'formatRange', dict(region, expectedToken=after['token'], format={'bold': True}))
            assert all(cell['bold'] for row in formatted['cells'] for cell in row), formatted
            numeric_region = {'sheet': 'Sheet1', 'address': 'B1'}
            numeric = call(c, 'readRange', numeric_region)
            call(c, 'formatRange', dict(numeric_region, expectedToken=numeric['token'], format={'numberFormat': '0.00'}))
            formulas_region = {'sheet': 'Sheet1', 'address': 'D1:D2'}
            formula_before = call(c, 'readRange', formulas_region)
            formula_after = call(c, 'setFormulas', dict(formulas_region, expectedToken=formula_before['token'], formulas=[['=SUM(B1,2)'], ['=1/0']]))
            computed = call(c, 'calculateRange', dict(formulas_region, expectedToken=formula_after['token']))
            assert computed['cells'][0][0]['value'] == 14.5, computed
            assert computed['cells'][1][0]['errorCode'] is not None and computed['cells'][1][0]['value'] is None, computed
            extra_region = {'sheet': 'Sheet1', 'address': 'F1:F3'}
            extra = call(c, 'readRange', extra_region)
            extra_written = call(c, 'writeRange', dict(extra_region, expectedToken=extra['token'], values=[[''], [False], [45292]]))
            assert extra_written['cells'][0][0]['value'] == '', extra_written
            date_region = {'sheet': 'Sheet1', 'address': 'F3'}
            date_read = call(c, 'readRange', date_region)
            date_format = call(c, 'formatRange', dict(date_region, expectedToken=date_read['token'], format={'numberFormat': 'yyyy-mm-dd'}))
            assert date_format['cells'][0][0]['value'] == 45292 and date_format['cells'][0][0]['text'] == '2024-01-01', date_format
            unicode_region = {'sheet': '数据', 'address': 'A1'}
            read = call(c, 'readRange', unicode_region)
            call(c, 'writeRange', dict(unicode_region, expectedToken=read['token'], values=[['跨区域']]))
            saved = call(c, 'save', {})
            assert saved['documentState']['persistenceState'] == 'saved', saved
            call(c, 'readRange', region)
            with client() as after_save_competitor:
                rejected(after_save_competitor, 'openWorkbook', {'path': str(workbook)}, 'DOCUMENT_LEASE_CONFLICT')
            # A second replacement must retain protection as well.
            final_unicode = call(c, 'readRange', unicode_region)
            call(c, 'writeRange', dict(unicode_region, expectedToken=final_unicode['token'], values=[['保存验证']]))
            call(c, 'save', {})
        assert primary.session_outcome['outcome'] == 'succeeded', primary.session_outcome
        # Reacquisition proves normal cleanup released the lease and kept the workbook usable.
        with client() as c:
            call(c, 'openWorkbook', {'path': str(workbook)})
            reread = call(c, 'readRange', {'sheet': 'Sheet1', 'address': 'A1:C2'})
            assert [[x['value'] for x in row] for row in reread['cells']] == values, reread
        # Verify the persisted archive independently of the live COM object.
        with zipfile.ZipFile(workbook) as archive:
            assert archive.testzip() is None
            sheet = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
            ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            disk = {node.attrib['r']: node for node in sheet.findall('.//s:c', ns)}
            assert float(disk['B1'].find('s:v', ns).text) == 12.5
            assert disk['D1'].find('s:f', ns).text == 'SUM(B1,2)'
        verify_replacement_quarantine(output)
        report['replacementQuarantineVerified'] = True
        report['status'] = 'passed'
        report['artifactVerified'] = True
        report['workbookLeftOpen'] = True
    except BaseException as exc:
        report['status'] = 'failed'
        report['error'] = repr(exc)
        report['hostStderr'] = primary.stderr
        if isinstance(exc, ActionFailed):
            report['failedResponse'] = exc.response
        raise
    finally:
        report['sessionOutcome'] = primary.session_outcome
        report['completedAt'] = time.strftime('%Y-%m-%dT%H:%M:%S%z')
        (output / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')


def main(argv=None):
    parser = argparse.ArgumentParser(description='Validate the Excel candidate slice on a new disposable .xlsx; existing user files are never edited.')
    parser.add_argument('--output-dir', required=True, type=Path, help='New directory for the disposable workbook and report')
    parser.add_argument('--wps-version', required=True, help='Installed WPS version recorded by the operator')
    parser.add_argument('--candidate', action='store_true', help='Developer acceptance only: use the test Host and target portfolio before production admission')
    args = parser.parse_args(argv)
    run(args.output_dir.resolve(), args.wps_version, candidate=args.candidate)
    print(args.output_dir.resolve() / 'report.json')
    return 0
