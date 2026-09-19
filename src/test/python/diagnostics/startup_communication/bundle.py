"""Build an isolated, hash-identified workspace snapshot with explicit test instrumentation."""
from pathlib import Path
import shutil
import subprocess
import sys
import zipfile

from common import digest, file_hashes, read_json, utc, validate_config, write_json

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[4]
RESOURCES = REPO / "src/test/resources/diagnostics/startup-communication"
FILES = ("doctor.py", "environment.ps1", "common.py", "worker.py", "child.py", "com_report.py", "com_diagnostics.ps1", "probe.ps1", "close_owned.ps1", "launch.ps1", "snapshot_io.ps1")


def build(destination, config_path):
    config = validate_config(read_json(config_path))
    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=False)
    sys.path.insert(0, str(REPO / "src/main/python"))
    from wps_skills.cli.build_excel_skill import build_excel_skill
    package = build_excel_skill(destination / "wps-excel")
    before = file_hashes(package)
    loop = package / "runtime/src/main/resources/wps_skills/windows/bridge_loop.ps1"
    text = loop.read_text(encoding="utf-8")
    original = """            $response = Invoke-BridgeOperation `
                -RequestId $requestId `
                -Operation ([string]$request.operation) `
                -Arguments $request.arguments"""
    replacement = """            # Test-package-only control-plane readiness probe; never a public Action.
            if ([string]$request.operation -ceq '__diagnostic_ready') {
                $response = New-SuccessRecord -RequestId $requestId -Data $request.arguments
            } else {
""" + original + "\n            }"
    if text.count(original) != 1:
        raise ValueError("Production bridge loop changed; review test instrumentation before building")
    loop.write_text(text.replace(original, replacement), encoding="utf-8")
    changes = [{"file": loop.relative_to(package).as_posix(), "beforeSha256": before[loop.relative_to(package).as_posix()],
                "afterSha256": digest(loop), "purpose": "Test-only readiness response after application scripts initialize; no COM call"}]
    shared = package / "runtime/src/main/resources/wps_skills/windows"
    shutil.copy2(RESOURCES / "com_diagnostics.ps1", shared / "com_diagnostics.ps1")
    changes.append({"file": (shared / "com_diagnostics.ps1").relative_to(package).as_posix(), "beforeSha256": None,
                    "afterSha256": digest(shared / "com_diagnostics.ps1"), "purpose": "Test-only COM observations; preserve original exceptions and fallback behavior"})
    excel = package / "runtime/src/main/resources/wps_skills/excel/windows"
    for name in ("excel_bridge.ps1", "excel_persistence.ps1"):
        target = excel / name
        content = target.read_text(encoding="utf-8")
        substitutions = {
            "[Runtime.InteropServices.Marshal]::GetActiveObject('KET.Application')": "Get-DiagnosticActiveApplication -ProgId 'KET.Application'",
            "New-Object -ComObject 'KET.Application'": "New-DiagnosticApplication -ProgId 'KET.Application'",
            "Get-ExcelComAvailability": "Get-DiagnosticExcelComAvailability",
        }
        if name == "excel_bridge.ps1":
            substitutions[". (Join-Path $PSScriptRoot '../../windows/bridge_common.ps1')"] = (
                ". (Join-Path $PSScriptRoot '../../windows/bridge_common.ps1')\n"
                ". (Join-Path $PSScriptRoot '../../windows/com_diagnostics.ps1')")
            substitutions["function Invoke-AcquireExistingDocument {"] = "function Invoke-OriginalAcquireExistingDocument {"
        else:
            substitutions["function Invoke-AcquireNewDocument {"] = "function Invoke-OriginalAcquireNewDocument {"
        for old, new in substitutions.items():
            if content.count(old) != 1:
                raise ValueError("COM diagnostic instrumentation needs review: " + name + " / " + old)
            content = content.replace(old, new)
        target.write_text(content, encoding="utf-8")
        changes.append({"file": target.relative_to(package).as_posix(), "beforeSha256": before[target.relative_to(package).as_posix()],
                        "afterSha256": digest(target), "purpose": "Observe registration, attach/activation, binding without changing fallback or document selection"})
    # Refresh the package's own manifest after the documented instrumentation.
    package_manifest = package / "runtime/files.sha256.json"
    package_manifest.unlink()
    write_json(package_manifest, file_hashes(package))
    for name in FILES:
        shutil.copy2((HERE if name.endswith(".py") else RESOURCES) / name, destination / name)
    (destination / "host-source").mkdir()
    for name in ("debug.py", "bundle.py"):
        shutil.copy2(HERE / name, destination / "host-source" / name)
    write_json(destination / "experiment.json", config)
    write_json(destination / "instrumentation.json", {
        "productionSourceModified": False,
        "changes": changes,
        "python": "child.py wraps the owned launcher and transport in this test process only",
        "measurementEffect": "Adds a readiness round trip, COM stage logs, registration inventory and an application Version read. Compare only identically instrumented candidates; these are not old baseline timings.",
    })
    commit = subprocess.run(["git", "rev-parse", "HEAD"], cwd=REPO, capture_output=True, text=True, check=True).stdout.strip()
    state = subprocess.run(["git", "status", "--porcelain"], cwd=REPO, capture_output=True, text=True, check=True).stdout
    write_json(destination / "manifest.json", {"createdUtc": utc(), "commit": commit, "dirty": bool(state),
               "workspaceStatus": state.splitlines(), "sourcePackageHashes": before, "files": file_hashes(destination)})
    archive = destination.parent / "bundle.zip"
    with zipfile.ZipFile(archive, "x", compression=zipfile.ZIP_DEFLATED) as output:
        for path in sorted(destination.rglob("*")):
            if path.is_file():
                output.write(path, path.relative_to(destination).as_posix())
    return archive
