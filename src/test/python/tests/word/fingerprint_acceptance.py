"""Opt-in native fingerprint regression through the standalone Word Task Client.

Fixtures: plain.docx (甲/乙/丙), multi.docx (three sections, six variants,
third linked to second), rich.docx (rich header), populated.docx (primary header).
The companion PowerShell acceptance additionally tests exact production hashes.
All test documents remain open; no Task replay or Saved flag reset.
"""
import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import uuid


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--skill', type=Path, required=True)
    parser.add_argument('--fixtures', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    if os.name != 'nt':
        parser.error('Requires Windows WPS interactive desktop')
    root = args.output.resolve()
    root.mkdir(parents=True, exist_ok=False)
    environment = dict(os.environ, PYTHONIOENCODING='utf-8', WPS_SKILLS_TASK_DIR=str(root / 'receipts'))
    environment.pop('PYTHONPATH', None)
    report = {'status': 'running', 'cases': []}
    prefix = 'fingerprint-' + uuid.uuid4().hex[:8]

    def action(name, params, identifier=None):
        item = {'address': {'app': 'word', 'action': name}, 'params': params}
        return dict(item, id=identifier) if identifier else item

    def ref(identifier, *path):
        return {'$ref': {'step': identifier, 'path': ['data', *path]}}

    def inspect(identifier):
        return action('inspectDocument', {'scope': {'kind': 'document'}, 'limits': {
            'maxTextCharacters': 4096, 'maxParagraphs': 128, 'maxRuns': 512}}, identifier)

    def submit(label, request):
        text = json.dumps(request, ensure_ascii=False, indent=2)
        source = root / (label + '.json')
        source.write_text(text, encoding='utf-8')
        (root / (label + '.original.json')).write_text(text, encoding='utf-8')
        proc = subprocess.run([sys.executable, str(args.skill.resolve() / 'scripts/word.py'),
                               '--app', 'word', '--task-file', str(source)], env=environment,
                              capture_output=True, timeout=150, creationflags=subprocess.CREATE_NO_WINDOW)
        (root / (label + '.stdout.json')).write_bytes(proc.stdout)
        (root / (label + '.stderr.txt')).write_bytes(proc.stderr)
        result = json.loads(proc.stdout)
        assert proc.returncode == 0 and result['outcome'] == 'succeeded', result
        for item in [result['document'], *result['steps'], *[v for v in result['completion'].values() if v]]:
            assert item['state'] == 'succeeded' and item['response']['outcome'] == 'succeeded', item
        report['cases'].append({'case': label, 'taskId': result['taskId'], 'status': 'passed'})
        return result

    try:
        for kind in ('plain', 'multi', 'rich'):
            source = root / (prefix + '-input-' + kind + '.docx')
            shutil.copyfile(args.fixtures / (kind + '.docx'), source)
            result = submit(kind, {'app': 'word', 'document': action('openDocument', {'path': str(source)}),
                                   'steps': [inspect('first'), inspect('second')], 'completion': []})
            assert result['document']['response']['data']['documentState']['persistenceState'] == 'saved', result
            first, second = [s['response']['data'] for s in result['steps']]
            assert first['revision'] == second['revision'] == result['document']['response']['data']['revision']
            assert first['documentState']['persistenceState'] == second['documentState']['persistenceState'] == 'saved'
            assert first['structure']['sections'] == second['structure']['sections']
        source = root / (prefix + '-input-edit.docx')
        shutil.copyfile(args.fixtures / 'plain.docx', source)
        result = submit('edit-save', {'app': 'word', 'document': action('openDocument', {'path': str(source)}),
            'steps': [action('replaceContent', {'target': {'kind': 'query', 'query': {
                'scope': {'kind': 'document'}, 'text': '甲', 'caseSensitive': True, 'wholeWord': False},
                'expectedMatchCount': 1}, 'replacement': {'kind': 'text', 'runs': [{'text': '甲修改'}]}}, 'edit')],
            'completion': [action('saveAs', {'outputPath': str(root / (prefix + '-edited-output.docx')), 'overwritePolicy': 'failIfExists'})]})
        assert result['document']['response']['data']['documentState']['persistenceState'] == 'saved'
        assert result['steps'][0]['response']['data']['matchedCount'] == 1
        assert result['completion']['save']['response']['data']['documentState']['persistenceState'] == 'saved'
        result = submit('new-header', {'app': 'word', 'document': action('createDocument', {}), 'steps': [
            action('writeContent', {'anchor': {'kind': 'documentEnd'}, 'blocks': [
                {'kind': 'paragraph', 'runs': [{'text': '指纹回归'}]}]}, 'write'), inspect('before'),
            action('setHeaderFooter', {'sections': {'kind': 'all', 'revision': ref('before', 'revision')},
                'updates': [{'area': 'header', 'variant': 'primary', 'operation': {'kind': 'replace', 'text': '新页眉'}}]}, 'header'),
            inspect('after')], 'completion': [action('saveAs', {'outputPath': str(root / (prefix + '-new-header.docx')), 'overwritePolicy': 'failIfExists'}),
                action('exportPdf', {'outputPath': str(root / (prefix + '-new-header.pdf')), 'overwritePolicy': 'failIfExists'})]})
        before, after = result['steps'][1]['response']['data'], result['steps'][3]['response']['data']
        assert before['revision'] != after['revision']
        assert next(s for s in after['structure']['sections'][0]['headerFooter']['stories']
                    if s['area'] == 'header' and s['variant'] == 'primary')['text'] == '新页眉'
        report['status'] = 'passed'
    except BaseException as exc:
        report.update(status='failed', error=str(exc))
        raise
    finally:
        (root / 'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')


if __name__ == '__main__':
    main()
