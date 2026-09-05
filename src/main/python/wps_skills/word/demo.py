"""Visible Word demonstration using one production Action Session."""

import argparse
from datetime import datetime
import json
import math
from pathlib import Path
import time
import uuid
import zipfile

from wps_skills.windows.desktop import require_desktop
from wps_skills.word.skill import open_session


INSPECTION = {
    'scope': {'kind': 'document'},
    'limits': {'maxTextCharacters': 4096, 'maxParagraphs': 128, 'maxRuns': 512},
}
TITLE = 'WPS Word · 原生自动化'
TABLE = [['能力', '验证方式'], ['文字与排版', '内容和格式读回'], ['原生表格', '逐格读取与显式保存']]


def create_demo_document(path):
    """Prepare an empty fixture; first-save is not a production Word Action."""
    files = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        'word/document.xml': '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>',
    }
    with zipfile.ZipFile(path, 'x', zipfile.ZIP_DEFLATED) as archive:
        for name, xml in files.items():
            archive.writestr(name, xml.encode('utf-8'))


def run_demo(output, *, delay=1.5, desktop_check=require_desktop, session_factory=open_session):
    desktop = desktop_check()
    if not math.isfinite(delay) or not 0 <= delay <= 10:
        raise ValueError('delay must be between 0 and 10 seconds')
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    document = output / ('wps-word-' + uuid.uuid4().hex + '.docx')
    report = {'status': 'running', 'document': str(document), 'desktopSessionId': desktop, 'steps': []}
    client = None

    def stage(message):
        print(message, flush=True)
        if delay:
            time.sleep(delay)

    def call(action, params):
        response = client.call({'app': 'word', 'action': action}, params)
        report['steps'].append(response)
        return response['data']

    try:
        create_demo_document(document)
        with session_factory() as client:
            stage('1/4 打开独立的 Word 演示文档。')
            call('openDocument', {'path': str(document)})
            stage('2/4 写入标题、正文，并设置字号、颜色与段落间距。')
            call('writeContent', {'anchor': {'kind': 'documentEnd'}, 'blocks': [
                {'kind': 'paragraph', 'runs': [{'text': TITLE, 'format': {'fontSizePt': 24, 'bold': True, 'color': '#245A81'}}], 'format': {'spaceAfterPt': 16}},
                {'kind': 'paragraph', 'runs': [{'text': '同一个精确文档会话：读取、编辑、验证、保存。', 'format': {'fontSizePt': 14, 'bold': False, 'color': '#404040'}}], 'format': {'spaceAfterPt': 12}},
            ]})
            stage('3/4 插入原生表格，核对完整文字和表格结构。')
            call('insertTable', {'anchor': {'kind': 'documentEnd'}, 'data': TABLE, 'headerRow': True})
            observed = call('inspectDocument', INSPECTION)
            if TITLE not in observed['text'] or observed['structure']['tableCount'] != 1:
                raise RuntimeError('演示内容未通过读回验证。')
            for row in TABLE:
                if any(text not in observed['text'] for text in row):
                    raise RuntimeError('演示表格文字未通过读回验证。')
            stage('4/4 显式保存，再次核对文档内容和保存状态。')
            report['saved'] = call('save', {})
            final = call('inspectDocument', INSPECTION)
            if final['text'] != observed['text'] or final['documentState']['persistenceState'] != 'saved':
                raise RuntimeError('保存后的文档未通过验证。')
        report['sessionOutcome'] = client.session_outcome
        if client.session_outcome['outcome'] != 'succeeded':
            raise RuntimeError('文档已保存，但会话清理未成功。')
        report['status'] = 'passed'
        print('演示完成，文档已保存，WPS 窗口保持打开：' + str(document), flush=True)
        return report
    except BaseException as exc:
        report.update(status='failed', error=str(exc))
        raise
    finally:
        if client is not None:
            report['sessionOutcome'] = client.session_outcome
        (output / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')


def main(argv=None):
    parser = argparse.ArgumentParser(description='在 Windows 桌面演示 WPS Word 编辑、表格、验证与保存。')
    parser.add_argument('--output-dir', type=Path, default=Path('build') / ('word-demo-' + datetime.now().strftime('%Y%m%d-%H%M%S-%f')))
    parser.add_argument('--delay', type=float, default=1.5)
    args = parser.parse_args(argv)
    run_demo(args.output_dir, delay=args.delay)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
