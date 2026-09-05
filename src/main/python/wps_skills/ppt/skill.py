"""Ppt Skill entry point and production Session Client assembly."""

import os
from pathlib import Path
import sys

from wps_skills.client.session_client import ActionFailed, SessionClient, SessionClientError
from wps_skills.cli.call import main


def open_session(*, timeout=60):
    env = dict(os.environ)
    env['PYTHONPATH'] = str(Path(__file__).resolve().parents[2])
    env['PYTHONIOENCODING'] = 'utf-8'
    return SessionClient([sys.executable, '-m', 'wps_skills.cli.call', '--session', '--app', 'ppt'],
                         application='ppt', env=env, timeout=timeout)
