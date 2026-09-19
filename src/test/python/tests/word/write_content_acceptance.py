"""Opt-in paragraph regression through complete Word Tasks; leaves documents open.

Migrated from the Session scenario: preserves paragraph/format/range, inline,
blank-tail, persisted XML and following-table assertions in the test harness.
"""

import argparse
import io
import json
from pathlib import Path
import sys
import zipfile
from xml.etree import ElementTree as ET


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--skill', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    sys.path.insert(0, str(args.skill.resolve() / 'runtime/src/main/python'))
    from wps_skills.word.skill import main as task_main
    args.output.mkdir(parents=True, exist_ok=False)
    report = {'status': 'running', 'cases': [], 'tasks': []}

    def action(name, params, identifier=None):
        value = {'address': {'app': 'word', 'action': name}, 'params': params}
        return dict(value, id=identifier) if identifier else value

    def ref(identifier, *path):
        return {'$ref': {'step': identifier, 'path': ['data', *path]}}

    def paragraph(text):
        return {'kind': 'paragraph', 'runs': [{'text': text}] if text else []}

    def write(identifier, blocks, anchor=None):
        return action('writeContent', {'anchor': anchor or {'kind': 'documentEnd'}, 'blocks': blocks}, identifier)

    def inspect(identifier):
        return action('inspectDocument', {'scope': {'kind': 'document'},
                      'limits': {'maxTextCharacters': 4096, 'maxParagraphs': 32, 'maxRuns': 64}}, identifier)

    def submit(label, steps, document=None, save=True):
        path = (args.output / (label + '.docx')).resolve()
        request = {'app': 'word', 'document': document or action('createDocument', {}), 'steps': steps,
                   'completion': [action('saveAs', {'outputPath': str(path), 'overwritePolicy': 'failIfExists'})] if save else []}
        source = args.output / (label + '.json')
        source.write_text(json.dumps(request, ensure_ascii=False), encoding='utf-8')
        stream = io.StringIO()
        code = task_main(['--app', 'word', '--task-file', str(source), '--timeout', '90'], output_stream=stream)
        response = json.loads(stream.getvalue())
        report['tasks'].append(response)
        assert code == 0 and response['cleanup']['outcome'] == 'succeeded', response
        return {s['id']: s['response']['data'] for s in response['steps']}, path

    def check(data, expected, label):
        assert not data['truncated']
        assert data['text'] == '\n'.join(expected), (label, data['text'], expected)
        assert data['structure']['paragraphCount'] == len(expected), (label, data)
        report['cases'].append(label)

    def body(path):
        with zipfile.ZipFile(path) as archive:
            return ET.fromstring(archive.read('word/document.xml')).find('{*}body')

    def persisted(path, expected):
        paragraphs = [''.join(t.text or '' for t in p.findall('.//{*}t')) for p in body(path).findall('{*}p')]
        assert paragraphs == expected, (path, paragraphs, expected)
        report['cases'].append(path.name + '-persisted')

    try:
        first = '首段 Chinese 🙂'
        expected = [first, '中间插入', '第二段', '最后标题', '连续追加 inline']
        data, path = submit('paragraphs', [
            write('first', [{'kind': 'paragraph', 'runs': [{'text': first, 'format': {'bold': True}}], 'format': {'alignment': 'center'}}]),
            inspect('single'),
            write('append', [paragraph('第二段'), {'kind': 'heading', 'level': 2, 'runs': [{'text': '最后标题', 'format': {'bold': False}}], 'format': {'alignment': 'right'}}], {'kind': 'after', 'range': ref('first', 'range')}),
            inspect('appended'),
            write('middle', [paragraph('中间插入')], {'kind': 'before', 'range': ref('appended', 'paragraphs', 1, 'range')}),
            inspect('middle_read'),
            write('tail', [paragraph('连续追加')]), inspect('tail_read'),
            write('inline', [{'kind': 'text', 'runs': [{'text': ' inline'}]}]), inspect('final'),
        ])
        check(data['single'], [first], 'single-paragraph')
        assert data['first']['range'] == data['single']['scopeRange']
        assert data['single']['paragraphs'][0]['complete']
        assert data['single']['paragraphs'][0]['format']['alignment'] == 'center'
        assert data['single']['paragraphs'][0]['runs'][0]['format']['bold'] is True
        check(data['appended'], [first, '第二段', '最后标题'], 'append-through-returned-range')
        assert data['appended']['paragraphs'][-1]['kind'] == 'heading'
        assert data['appended']['paragraphs'][-1]['level'] == 2
        assert data['appended']['paragraphs'][-1]['format']['alignment'] == 'right'
        check(data['middle_read'], expected[:4], 'middle-insertion')
        assert data['middle_read']['paragraphs'][0]['format']['alignment'] == 'center'
        assert data['middle_read']['paragraphs'][-1]['kind'] == 'heading'
        check(data['tail_read'], expected[:4] + ['连续追加'], 'repeated-document-end')
        check(data['final'], expected, 'inline-text-unchanged')
        persisted(path, expected)
        _, table_path = submit('following-table', [action('insertTable', {'anchor': {'kind': 'documentEnd'}, 'data': [['cell A', 'cell B']], 'headerRow': False}, 'table')], action('openDocument', {'path': str(path)}))
        xml = body(table_path)
        preceding = [''.join(t.text or '' for t in p.findall('.//{*}t')) for p in xml.findall('{*}p')]
        assert preceding[:len(expected)] == expected
        tables = xml.findall('{*}tbl')
        assert len(tables) == 1 and [t.text for t in tables[0].findall('.//{*}t')] == ['cell A', 'cell B']
        report['cases'].append('table-after-last-paragraph')
        data, path = submit('intentional-blanks', [write('blank', [paragraph(''), paragraph('正文'), paragraph('')], {'kind': 'documentStart'}), inspect('read')])
        check(data['read'], ['', '正文', ''], 'requested-blank-paragraphs')
        persisted(path, ['', '正文', ''])
        data, _ = submit('empty-insertion', [write('empty', [paragraph('')]), inspect('empty_read'), write('fill', [paragraph('填入尾段')]), inspect('filled')], save=False)
        check(data['empty_read'], ['', ''], 'explicit-empty-insertion')
        check(data['filled'], ['', '填入尾段'], 'reuse-existing-empty-tail')
        report['status'] = 'passed'
    except BaseException as exc:
        report.update(status='failed', error=str(exc))
        raise
    finally:
        (args.output / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps({key: report[key] for key in ('status', 'cases')}, ensure_ascii=True))


if __name__ == '__main__':
    main()
