"""Opt-in live acceptance Host, never shipped as a production execution entry."""
import sys
import uuid
from wps_skills.excel.windows.session import build_session
from wps_skills.core.trace_journal import JsonlTraceJournal
from wps_skills.excel.contracts import EXCEL_TARGET_CONTRACT_SET
from wps_skills.host.session_host import SessionHost

journal = JsonlTraceJournal.default()
host = SessionHost(
    session_factory=lambda *, application, session_id: build_session(session_id=session_id, contracts=EXCEL_TARGET_CONTRACT_SET),
    session_id_factory=lambda: 'excel-acceptance-' + uuid.uuid4().hex,
    trace_log_factory=journal.session_log, action_trace_factory=journal.action_trace, session_event_sink=journal.session_event,
)
raise SystemExit(host.serve(application='excel', input_stream=sys.stdin, output_stream=sys.stdout, error_stream=sys.stderr))
