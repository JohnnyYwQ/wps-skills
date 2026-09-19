"""Build one ppt Task with the shared resource policy."""
from wps_skills.windows.task_factory import build_task as _build_task


def build_task(**kwargs):
    return _build_task(application="ppt", **kwargs)
