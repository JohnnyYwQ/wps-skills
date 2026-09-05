"""Visible desktop demonstration using one production Excel Action Session."""

import argparse
from datetime import datetime
import json
import math
import os
from pathlib import Path
import time
import uuid
import zipfile

from wps_skills.excel.skill import open_session


from wps_skills.windows.desktop import require_desktop, visible_document_window


def create_demo_workbook(path):
    # Prepare a blank display template. All demonstrated edits go through the
    # production Excel Session; this is not a new createWorkbook Action.
    files = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView windowWidth="20000" windowHeight="13000"/></bookViews><sheets><sheet name="演示" sheetId="1" r:id="rId1"/></sheets></workbook>',
        'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        'xl/worksheets/sheet1.xml': '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" zoomScale="130"/></sheetViews><sheetFormatPr defaultRowHeight="28"/><cols><col min="1" max="4" width="18" customWidth="1"/></cols><sheetData/></worksheet>',
    }
    with zipfile.ZipFile(path, 'x', zipfile.ZIP_DEFLATED) as archive:
        for name, xml in files.items():
            archive.writestr(name, xml.encode('utf-8'))


def run_demo(output, *, delay=1.5, desktop_check=require_desktop,
             session_factory=open_session, window_probe=visible_document_window):
    session_id = desktop_check()
    if not math.isfinite(delay) or not 0 <= delay <= 10:
        raise ValueError('delay must be between 0 and 10 seconds')
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    # WPS rejects equal workbook basenames even when their directories differ.
    workbook = output / (output.name + '-' + uuid.uuid4().hex + '.xlsx')
    create_demo_workbook(workbook)
    report = {'status': 'running', 'desktopSessionId': session_id, 'workbook': str(workbook), 'steps': []}

    def stage(message):
        print(message, flush=True)
        if delay:
            time.sleep(delay)

    def call(client, action, params):
        response = client.call({'app': 'excel', 'action': action}, params)
        report['steps'].append(response)
        return response['data']

    def read(client, address):
        return call(client, 'readRange', {'sheet': '演示', 'address': address})

    def mutate(client, action, address, **params):
        before = read(client, address)
        return call(client, action, dict(sheet='演示', address=address, expectedToken=before['token'], **params))

    client = None
    try:
        with session_factory() as client:
            print('正在打开 WPS 表格窗口……', flush=True)
            call(client, 'openWorkbook', {'path': str(workbook)})
            info = call(client, 'getWorkbookInfo', {})
            window_reference = info.get('window')
            deadline = time.monotonic() + 10
            window = window_probe(window_reference)
            while window is None and time.monotonic() < deadline:
                time.sleep(0.2)
                window = window_probe(window_reference)
            if window is None:
                raise RuntimeError('未检测到演示工作簿的可见窗口，已停止演示。请在 Windows 桌面运行，并检查 WPS 窗口。')
            report['windowBefore'] = window
            stage('1/4 窗口已显示，即将填入商品、数量和单价。')
            mutate(client, 'writeRange', 'A1:D5', values=[
                ['商品', '数量', '单价', '金额'], ['键盘', 2, 199, None],
                ['鼠标', 3, 89, None], ['显示器', 1, 1299, None], ['合计', None, None, None],
            ])
            stage('2/4 数据已写入，即将填入公式并计算金额。')
            mutate(client, 'setFormulas', 'D2:D5', formulas=[['=B2*C2'], ['=B3*C3'], ['=B4*C4'], ['=SUM(D2:D4)']])
            computed = mutate(client, 'calculateRange', 'D2:D5')
            if computed['cells'][3][0]['value'] != 1964:
                raise RuntimeError('合计金额未通过验证，已停止保存。')
            stage('3/4 合计为 1964，即将加粗表头和合计行、设置金额格式。')
            mutate(client, 'formatRange', 'A1:D1', format={'bold': True})
            mutate(client, 'formatRange', 'A5:D5', format={'bold': True})
            mutate(client, 'formatRange', 'C2:D5', format={'numberFormat': '0.00'})
            stage('4/4 格式已设置，即将保存并保留工作簿窗口。')
            call(client, 'save', {})
            final = read(client, 'A1:D5')
            if final['cells'][4][3]['value'] != 1964:
                raise RuntimeError('保存后的合计金额未通过验证。')
        report['windowAfter'] = window_probe(window_reference)
        if report['windowAfter'] is None:
            raise RuntimeError('文件已保存，但会话结束后未检测到可见窗口。')
        report['status'] = 'passed'
        print('演示完成，WPS 窗口保持打开。文件：' + str(workbook), flush=True)
    except BaseException as exc:
        report['status'] = 'failed'
        report['error'] = str(exc)
        raise
    finally:
        if client is not None:
            report['sessionOutcome'] = client.session_outcome
        (output / 'demo-report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    return workbook


def main(argv=None):
    parser = argparse.ArgumentParser(description='在 Windows 桌面显示 WPS 窗口，分步演示 Excel 编辑与保存。')
    parser.add_argument('--output-dir', type=Path, default=None, help='尚不存在的演示输出目录')
    parser.add_argument('--delay', type=float, default=1.5, help='各阶段之间的展示停留秒数，0 到 10，默认 1.5')
    args = parser.parse_args(argv)
    output = args.output_dir or Path('build') / ('excel-demo-' + datetime.now().strftime('%Y%m%d-%H%M%S-%f'))
    try:
        run_demo(output, delay=args.delay)
    except Exception as exc:
        parser.exit(1, str(exc) + '\n')
    return 0
