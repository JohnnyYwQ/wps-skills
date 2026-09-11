"""Assemble a standalone Application Skill without maintaining a second runtime copy."""

import hashlib
import json
from pathlib import Path
import shutil
import tempfile


MAIN = Path(__file__).resolve().parents[3]


def build_application_skill(application, destination):
    if application not in {"word", "excel", "ppt"}:
        raise ValueError("unsupported Skill application")
    skill_name = "wps-" + application
    destination = Path(destination).resolve()
    if destination.exists():
        raise FileExistsError(f"Destination already exists: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=skill_name + "-build-", dir=destination.parent) as temporary:
        stage = Path(temporary) / skill_name
        ignore = shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store")
        shutil.copytree(MAIN / "resources" / "skills" / skill_name, stage, ignore=ignore)
        runtime = stage / "runtime" / "src" / "main"
        python_source = MAIN / "python" / "wps_skills"
        python_target = runtime / "python" / "wps_skills"
        python_target.mkdir(parents=True)
        shutil.copy2(python_source / "__init__.py", python_target / "__init__.py")
        # Each independently installed Skill carries its own shared runtime copy.
        # Other applications and repository build/demo tools are not dependencies.
        runtime_ignore = shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store", "demo.py", "desktop.py")
        for package in ("core", "client", "host", "windows", application):
            shutil.copytree(python_source / package, python_target / package, ignore=runtime_ignore)
        (python_target / "cli").mkdir()
        for name in ("__init__.py", "call.py"):
            shutil.copy2(python_source / "cli" / name, python_target / "cli" / name)
        resource_source = MAIN / "resources" / "wps_skills"
        resource_target = runtime / "resources" / "wps_skills"
        resource_ignore = shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store", "demo_launcher.ps1", "demo-template.pptx")
        for package in ("windows", application):
            shutil.copytree(resource_source / package, resource_target / package, ignore=resource_ignore)
        files = {
            path.relative_to(stage).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest()
            for path in sorted(stage.rglob("*")) if path.is_file()
        }
        (stage / "runtime" / "files.sha256.json").write_text(
            json.dumps(files, indent=2) + "\n", encoding="utf-8"
        )
        stage.rename(destination)
    return destination

