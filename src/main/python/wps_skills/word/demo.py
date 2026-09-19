"""Visible Word demonstration using one production Task."""

import argparse
from datetime import datetime
import json
import math
from pathlib import Path
import time
import uuid
import zipfile

from wps_skills.windows.desktop import require_desktop
from wps_skills.word.skill import main as task_main
import io


INSPECTION = {
    'scope': {'kind': 'document'},
    'limits': {'maxTextCharacters': 4096, 'maxParagraphs': 128, 'maxRuns': 512},
}
TITLE = 'WPS Word · 原生自动化'
TABLE = [['能力', '验证方式'], ['文字与排版', '内容和格式读回'], ['原生表格', '逐格读取与显式保存']]


def create_demo_document(path):
    """Prepare an empty fixture; used only for isolated acceptance input."""
    files = {
        '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        'word/document.xml': '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>',
    }
    with zipfile.ZipFile(path, 'x', zipfile.ZIP_DEFLATED) as archive:
        for name, xml in files.items():
            archive.writestr(name, xml.encode('utf-8'))


def run_demo(output, *, delay=1.5, desktop_check=require_desktop, task_entry=task_main):
    desktop = desktop_check()
    if not math.isfinite(delay) or not 0 <= delay <= 10:
        raise ValueError('delay must be between 0 and 10 seconds')
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    document = output / ('wps-word-' + uuid.uuid4().hex + '.docx')
    def action(name, params, identifier=None):
        result = {'address': {'app': 'word', 'action': name}, 'params': params}
        if identifier:
            result['id'] = identifier
        return result
    request = {
        'app': 'word', 'document': action('createDocument', {}),
        'steps': [
            action('writeContent', {'anchor': {'kind': 'documentEnd'}, 'blocks': [
                {'kind': 'paragraph', 'runs': [{'text': TITLE, 'format': {'fontSizePt': 24, 'bold': True, 'color': '#245A81'}}], 'format': {'spaceAfterPt': 16}},
                {'kind': 'paragraph', 'runs': [{'text': '同一个精确文档：读取、编辑、验证、保存。', 'format': {'fontSizePt': 14, 'bold': False, 'color': '#404040'}}], 'format': {'spaceAfterPt': 12}},
            ]}, 'write'),
            action('insertTable', {'anchor': {'kind': 'documentEnd'}, 'data': TABLE, 'headerRow': True}, 'table'),
            action('inspectDocument', INSPECTION, 'inspect'),
        ],
        'completion': [action('saveAs', {'outputPath': str(document), 'overwritePolicy': 'failIfExists'})],
    }
    source = output / 'task.json'
    source.write_text(json.dumps(request, ensure_ascii=False), encoding='utf-8')
    report = {'status': 'running', 'document': str(document), 'desktopSessionId': desktop}
    try:
        print('执行 Word Task：创建、写入、插入表格、读取、保存。', flush=True)
        stream = io.StringIO()
        code = task_entry(['--app', 'word', '--task-file', str(source)], output_stream=stream)
        response = json.loads(stream.getvalue())
        report['task'] = response
        if code != 0:
            raise RuntimeError('Word Task did not complete: ' + json.dumps(response, ensure_ascii=False))
        observed = response['steps'][-1]['response']['data']
        if TITLE not in observed['text'] or observed['structure']['tableCount'] != 1:
            raise RuntimeError('演示内容未通过读回验证。')
        if any(text not in observed['text'] for row in TABLE for text in row):
            raise RuntimeError('演示表格文字未通过读回验证。')
        report.update(status='passed', saved=response['completion']['save']['response']['data'])
        print('演示完成，文档已保存，WPS 窗口保持打开：' + str(document), flush=True)
        return report
    except BaseException as exc:
        report.update(status='failed', error=str(exc))
        raise
    finally:
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
