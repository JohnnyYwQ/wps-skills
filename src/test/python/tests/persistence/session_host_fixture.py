"""Opt-in three-application candidate host for persistence acceptance."""
import importlib
import sys
import uuid
from wps_skills.host.session_host import SessionHost

app=sys.argv[1]
contracts=getattr(importlib.import_module(f'wps_skills.{app}.contracts'),app.upper()+'_TARGET_CONTRACT_SET')
build_session=importlib.import_module(f'wps_skills.{app}.windows.session').build_session
host=SessionHost(session_factory=lambda *, application, session_id: build_session(session_id=session_id,contracts=contracts),
                 session_id_factory=lambda: 'persistence-'+uuid.uuid4().hex)
raise SystemExit(host.serve(application=app,input_stream=sys.stdin,output_stream=sys.stdout,error_stream=sys.stderr))
