"""Candidate-only desktop demo used before production admission."""
import os
from pathlib import Path
import sys
from wps_skills.client.session_client import SessionClient
from wps_skills.ppt.demo import run_demo

ROOT=Path(__file__).resolve().parents[5]

def open_candidate():
    env=dict(os.environ,PYTHONPATH=str(ROOT/'src/main/python'),PYTHONIOENCODING='utf-8')
    return SessionClient([sys.executable,str(Path(__file__).with_name('session_host_fixture.py'))],
                         application='ppt',env=env,timeout=90)

if __name__=='__main__':
    run_demo(Path(sys.argv[1]),delay=0,session_factory=open_candidate)
