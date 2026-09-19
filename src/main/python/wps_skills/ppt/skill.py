"""Ppt Task submission, receipt lookup, and side-effect-free discovery."""

from wps_skills.cli.task import main as _main


def main(argv=None, **kwargs):
    return _main(argv, required_application='ppt', **kwargs)
