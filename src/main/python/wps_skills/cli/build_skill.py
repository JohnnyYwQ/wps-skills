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
        shutil.copytree(MAIN / "python" / "wps_skills", runtime / "python" / "wps_skills", ignore=ignore)
        shutil.copytree(MAIN / "resources" / "wps_skills", runtime / "resources" / "wps_skills", ignore=ignore)
        files = {
            path.relative_to(stage).as_posix(): hashlib.sha256(path.read_bytes()).hexdigest()
            for path in sorted(stage.rglob("*")) if path.is_file()
        }
        (stage / "runtime" / "files.sha256.json").write_text(
            json.dumps(files, indent=2) + "\n", encoding="utf-8"
        )
        stage.rename(destination)
    return destination

