"""Build a relocatable Windows Action acceptance kit, separate from product Skills."""
from pathlib import Path
import hashlib
import json
import shutil
import subprocess
import zipfile
from wps_skills.cli.build_skill import build_application_skill

ROOT=Path(__file__).resolve().parents[5]


def build(destination):
    destination=Path(destination).resolve()
    destination.mkdir(parents=True,exist_ok=False)
    for app in ('word','excel','ppt'):
        build_application_skill(app,destination/'skills'/('wps-'+app))
        (destination/('run_'+app+'.py')).write_text('''from pathlib import Path
import sys
ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'lib'))
from tests.applications.action_acceptance import main
if __name__=='__main__':
    raise SystemExit(main('''+repr(app)+''',package_root=ROOT))
''',encoding='utf-8')
    library=destination/'lib/tests/applications';library.mkdir(parents=True)
    (library.parent/'__init__.py').write_text('');(library/'__init__.py').write_text('')
    for name in ('action_acceptance.py','action_suite.py','complex_plan.py','native_acceptance.py','response_expectations.py'):
        shutil.copy2(Path(__file__).parent/name,library/name)
    shutil.copytree(ROOT/'src/test/resources/acceptance',destination/'resources')
    guide=ROOT/'docs/testing/action-execution-plan.md'
    if guide.exists():shutil.copy2(guide,destination/'README.md')
    paths={p.relative_to(destination).as_posix():hashlib.sha256(p.read_bytes()).hexdigest() for p in destination.rglob('*') if p.is_file()}
    (destination/'kit-manifest.json').write_text(json.dumps({'commit':subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT,text=True).strip(),'source':'current workspace','files':paths},indent=2),encoding='utf-8')
    archive=destination.with_suffix('.zip')
    with zipfile.ZipFile(archive,'x',zipfile.ZIP_DEFLATED) as output:
        for path in destination.rglob('*'):
            if path.is_file():output.write(path,path.relative_to(destination).as_posix())
    return archive
