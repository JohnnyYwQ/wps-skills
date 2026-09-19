"""Word assembly through the shared Windows Task resource factory."""
from pathlib import Path
from wps_skills.windows.task_factory import build_task as _build_task, TaskStartupError

BRIDGE_SCRIPT = (Path(__file__).resolve().parents[4] / "resources" / "wps_skills" /
                 "word" / "windows" / "word_bridge.ps1")


def build_task(**kwargs):
    return _build_task(application="word", **kwargs)
