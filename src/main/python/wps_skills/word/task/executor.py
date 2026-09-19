"""Word-owned specialization of the shared Task executor."""

from wps_skills.core.task_executor import ApplicationTask


class WordTask(ApplicationTask):
    def __init__(self, **kwargs):
        super().__init__(application="word", **kwargs)
