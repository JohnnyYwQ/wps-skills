"""Build isolated, hash-recorded test copies without changing production files."""
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import zipfile

REPO = Path(__file__).resolve().parents[5]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def files(root):
    return {p.relative_to(root).as_posix(): sha(p) for p in root.rglob('*')
            if p.is_file() and '__pycache__' not in p.parts and p.suffix != '.pyc'}


def build(destination):
    from wps_skills.cli.build_skill import build_application_skill
    destination = Path(destination).resolve()
    destination.mkdir(parents=True, exist_ok=False)
    resources = REPO / 'src/test/resources/diagnostics/design-value'
    changes = []
    for app in ('word', 'excel', 'ppt'):
        package = destination / 'skills' / ('wps-' + app)
        build_application_skill(app, package)

        def patch(relative, old, new):
            path = package / relative
            original = path.read_text(encoding='utf-8-sig')
            if original.count(old) != 1:
                raise ValueError('Instrumentation must be reviewed: ' + relative)
            before = sha(path)
            path.write_text(original.replace(old, new), encoding='utf-8-sig')
            changes.append(dict(app=app, file=relative, beforeSha256=before,
                                afterSha256=sha(path), removed=old, inserted=new))

        shared = 'runtime/src/main/resources/wps_skills/windows/'
        target = package / shared / 'design_hook.ps1'
        target.write_text((resources / 'design_hook.ps1').read_text(encoding='utf-8-sig'), encoding='utf-8-sig')
        patch(shared + 'bridge_loop.ps1', '$timingWriter = Get-Command Write-WordBridgeTiming',
              ". (Join-Path $PSScriptRoot 'design_hook.ps1')\n$timingWriter = Get-Command Write-WordBridgeTiming")
        patch(shared + 'bridge_loop.ps1', '            $response = Invoke-BridgeOperation `',
              "            Invoke-DesignHook -Stage 'before' -Operation ([string]$request.operation)\n            $response = Invoke-BridgeOperation `")
        patch(shared + 'bridge_loop.ps1', '        try { Write-BridgeRecord -Record $response }',
              "        Invoke-DesignHook -Stage 'after' -Operation ([string]$request.operation)\n        try { Write-BridgeRecord -Record $response }")
        patch(shared + 'bridge_common.ps1', '        return & $Action',
              "        $designResult = & $Action\n        Invoke-DesignHook -Stage 'inflight_after' -Operation $script:DesignCurrentOperation\n        return $designResult")
        if app == 'excel':
            relative = 'runtime/src/main/resources/wps_skills/excel/windows/excel_actions.ps1'
            line = "                                    [void]$cell.GetType().InvokeMember('Value2', [Reflection.BindingFlags]::SetProperty, $null, $cell, [object[]]@($literal))"
            patch(relative, line, "                                    if (-not $env:WPS_DESIGN_NO_WRITE) {\n" + line + "\n                                    } else { Write-DesignEvent 'write_suppressed' @{address=$Params.address} }")
            line = '        return Get-ExcelSnapshot -Sheet $sheet -Range $range -Address ([string]$Params.address)'
            candidate = '''        if ($env:WPS_DESIGN_OPTIMISTIC -and $Operation -eq 'write_literal_rectangle') {
            for($r=0;$r -lt $before.cells.Count;$r++) {
                for($c=0;$c -lt $before.cells[$r].Count;$c++) {
                    $before.cells[$r][$c].value=$Params.values[$r][$c]
                    $before.cells[$r][$c].text=[string]$Params.values[$r][$c]
                    $before.cells[$r][$c].formula=$null
                    $before.cells[$r][$c].errorCode=$null
                }
            }
            Write-DesignEvent 'optimistic_result' @{address=$Params.address}
            return $before
        }
'''
            patch(relative, line, candidate + line)
        manifest = package / 'runtime/files.sha256.json'
        manifest.unlink()
        manifest.write_text(json.dumps(files(package), indent=2),encoding='utf-8')

    shutil.copytree(Path(__file__).parent, destination / 'design_value',
                    ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
    library = destination / 'lib/tests/applications'
    library.mkdir(parents=True)
    (library.parent / '__init__.py').write_text('')
    (library / '__init__.py').write_text('')
    for name in ('complex_plan.py', 'native_acceptance.py', 'action_suite.py'):
        shutil.copy2(REPO / 'src/test/python/tests/applications' / name, library / name)
    (destination / 'resources').mkdir()
    for source in resources.glob('*.ps1'):
        (destination / 'resources' / source.name).write_text(source.read_text(encoding='utf-8-sig'), encoding='utf-8-sig')
    shutil.copy2(REPO / 'src/test/resources/acceptance/environment.ps1', destination / 'resources/environment.ps1')
    shutil.copy2(REPO / 'docs/testing/host-independent-experiments.md', destination / 'PLAN.md')
    (destination / 'README.md').write_text('''# WPS 设计实验独立包

这是测试包，不是供 Agent 安装的正式 Skill。需要已通过环境检查的 Windows 交互桌面、Python 与 WPS。

在终端运行：

```text
python run.py --help
python run.py --root C:/wps-tests/runs/new-run-01 --groups references preflight replay focus dynamic concurrency stale persistence --trials 3
```

结果目录必须尚不存在。用 `--apps word`、`--apps excel`、`--apps ppt` 可分别运行。故障会停止所在批次；下次换新目录，不重放原 Task。

只关闭已保存并验证的成功测试文档；未保存/失败现场与 Quarantine 保留，不退出 WPS。`PLAN.md` 是事先声明的判断标准。`instrumentation.json` 记录测试补丁；`manifest.json` 记录逐文件 SHA-256。输出中包含原始请求、配置、响应、独立文档观察、回执和阶段日志。

Mac 调度与多轮汇总使用仓库中的 `scripts/debug/design_value.py`，详见仓库 `docs/testing/design-experiment-usage.md`。
''',encoding='utf-8')
    (destination / 'run.py').write_text('''from pathlib import Path
import sys
ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'lib'))
sys.path.insert(0,str(ROOT/'skills/wps-word/runtime/src/main/python'))
from design_value.runner import main
if __name__=='__main__':raise SystemExit(main(ROOT))
''')
    (destination / 'instrumentation.json').write_text(json.dumps({'productionModified': False, 'changes': changes}, ensure_ascii=False, indent=2),encoding='utf-8')
    (destination / 'manifest.json').write_text(json.dumps({'commit': subprocess.check_output(['git','rev-parse','HEAD'],cwd=REPO,text=True).strip(), 'source':'current uncommitted workspace', 'files':files(destination)}, indent=2),encoding='utf-8')
    archive = destination.with_suffix('.zip')
    with zipfile.ZipFile(archive, 'x', zipfile.ZIP_DEFLATED) as output:
        for path in destination.rglob('*'):
            if path.is_file(): output.write(path, path.relative_to(destination).as_posix())
    return archive
