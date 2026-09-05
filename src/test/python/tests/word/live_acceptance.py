"""Native Word demo acceptance: saved content and exact visible document window."""
import argparse
import json
from pathlib import Path
import subprocess
import zipfile
from xml.etree import ElementTree as ET

from wps_skills.windows.bridge_runtime import windows_powershell_executable
from wps_skills.windows.desktop import visible_document_window
from wps_skills.word.demo import run_demo, TABLE, TITLE


def verify_demo(report):
    document = Path(report['document'])
    with zipfile.ZipFile(document) as archive:
        body = ET.fromstring(archive.read('word/document.xml'))
    text = ''.join(node.text or '' for node in body.findall('.//{*}t'))
    assert TITLE in text
    tables = body.findall('.//{*}tbl')
    assert len(tables) == 1
    values = [[''.join(t.text or '' for t in cell.findall('.//{*}t')) for cell in row.findall('{*}tc')]
              for row in tables[0].findall('{*}tr')]
    assert values == TABLE, values
    probe = Path(__file__).resolve().parents[3] / 'resources/wps_skills/word/observe_demo_window.ps1'
    result = subprocess.run([windows_powershell_executable(), '-NoProfile', '-NonInteractive',
                             '-ExecutionPolicy', 'Bypass', '-File', str(probe), '-DocumentPath', str(document)],
                            capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, timeout=30)
    assert result.returncode == 0, result.stderr.decode('utf-8', errors='replace')
    window = visible_document_window(json.loads(result.stdout))
    assert window, 'Exact Word demo window is not visible'
    return {'persistedText': True, 'persistedTable': True, 'visibleWindow': window}


def main(argv=None):
    parser = argparse.ArgumentParser(description='Verify the native Word demo workflow (not all Word Actions).')
    parser.add_argument('--output-dir', type=Path, required=True)
    args = parser.parse_args(argv)
    report = run_demo(args.output_dir, delay=0)
    acceptance = {'status': 'running'}
    try:
        acceptance.update(verify_demo(report), status='passed')
    except BaseException as exc:
        acceptance.update(status='failed', error=str(exc))
        raise
    finally:
        (args.output_dir / 'acceptance.json').write_text(json.dumps(acceptance, ensure_ascii=False, indent=2), encoding='utf-8')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
