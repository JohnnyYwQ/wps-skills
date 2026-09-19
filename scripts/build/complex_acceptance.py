"""Build current workspace's isolated 60-business-Task Windows test kit."""
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import zipfile
ROOT=Path(__file__).resolve().parents[2]
sys.path[:0]=[str(ROOT/'src/main/python'),str(ROOT/'src/test/python')]
from tests.applications.build_acceptance import build


def main():
    p=argparse.ArgumentParser();p.add_argument('--output',required=True,type=Path);a=p.parse_args()
    dest=a.output.resolve();archive=build(dest)
    for name in ('complex_business_suite.py','complex_business_acceptance.py'):
        shutil.copy2(ROOT/'src/test/python/tests/applications'/name,dest/'lib/tests/applications'/name)
    for app in ('word','excel','ppt'):
        (dest/('run_'+app+'.py')).write_text("from pathlib import Path\nimport sys\nROOT=Path(__file__).resolve().parent\nsys.path.insert(0,str(ROOT/'lib'))\nfrom tests.applications.complex_business_acceptance import main\nif __name__=='__main__': raise SystemExit(main(ROOT, fixed_app="+repr(app)+"))\n",encoding='utf-8')
    shutil.copy2(ROOT/'docs/testing/complex-task-execution-plan.md',dest/'README.md')
    snapshot={}
    for folder in ('src/main','src/test/python/tests/applications','src/test/resources/acceptance','scripts/build'):
        for f in (ROOT/folder).rglob('*'):
            if f.is_file() and '__pycache__' not in f.parts:snapshot[f.relative_to(ROOT).as_posix()]=hashlib.sha256(f.read_bytes()).hexdigest()
    (dest/'source-snapshot.json').write_text(json.dumps({'commit':subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT,text=True).strip(),'workingTreeStatus':subprocess.check_output(['git','status','--porcelain'],cwd=ROOT,text=True),'files':snapshot},ensure_ascii=False,indent=2),encoding='utf-8')
    paths={f.relative_to(dest).as_posix():hashlib.sha256(f.read_bytes()).hexdigest() for f in dest.rglob('*') if f.is_file() and f.name!='kit-manifest.json'}
    (dest/'kit-manifest.json').write_text(json.dumps({'source':'current workspace','files':paths},indent=2),encoding='utf-8')
    with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED) as z:
        for f in dest.rglob('*'):
            if '__pycache__' not in f.parts:z.write(f,f.relative_to(dest).as_posix())
    print(archive)
if __name__=='__main__':main()
