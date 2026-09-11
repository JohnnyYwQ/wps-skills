#!/usr/bin/env python3
"""Run one canonical application-scoped WPS Session Host."""

import argparse
import importlib.util
import json
import math
from pathlib import Path
import sys
import uuid

from wps_skills.core.trace_journal import JsonlTraceJournal
from wps_skills.host.session_host import SessionHost, SessionStartupError


def _parser():
    parser = argparse.ArgumentParser(description="Discover Actions or run a WPS Action Session Host")
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--session", action="store_true")
    mode.add_argument("--index", action="store_true", help="Print the production Action Index without starting WPS")
    mode.add_argument("--resolve", nargs="+", metavar="ACTION", help="Resolve complete production Action Contracts without executing them")
    mode.add_argument("--start", action="store_true", help="Start a task-local managed Session Client")
    mode.add_argument("--call", metavar="HANDLE", help="Submit one Action or retrieve its existing step receipt")
    mode.add_argument("--status", metavar="HANDLE", help="Read Session status or a retained --step receipt")
    mode.add_argument("--close", metavar="HANDLE", help="End the Session without saving or closing the document")
    parser.add_argument("--step", type=int, help="Use nextStep from the previous result; required with --call")
    parser.add_argument("--action", help="Application-local Action name for --call")
    params = parser.add_mutually_exclusive_group()
    params.add_argument("--params-json", help="One JSON parameter object (default: {})")
    params.add_argument("--params-file", help="UTF-8 JSON parameter file; data only, no executable code")
    params.add_argument("--params-stdin", action="store_true", help="Read one JSON parameter object from stdin")
    parser.add_argument("--timeout", type=float, default=60, help="Managed startup/Action timeout or command wait in seconds (default: 60)")
    parser.add_argument(
        "--app",
        choices=("excel", "ppt", "word"),
        required=True,
    )
    parser.add_argument(
        "--debug-close-created-document",
        action="store_true",
        help=(
            "debug only: discard and close a document created by this "
            "Session during cleanup"
        ),
    )
    return parser


def _discover(args, output_stream, error_stream):
    from wps_skills.core.action_session import ActionAddress
    if importlib.util.find_spec("wps_skills." + args.app) is None:
        error_stream.write(f"WPS_DISCOVERY_UNAVAILABLE app={args.app}\n")
        return 4
    if args.app == "word":
        from wps_skills.word.contracts import WORD_PRODUCTION_CONTRACT_SET
        contracts = WORD_PRODUCTION_CONTRACT_SET
    elif args.app == "excel":
        from wps_skills.excel.contracts import EXCEL_PRODUCTION_CONTRACT_SET
        contracts = EXCEL_PRODUCTION_CONTRACT_SET
    elif args.app == "ppt":
        from wps_skills.ppt.contracts import PPT_PRODUCTION_CONTRACT_SET
        contracts = PPT_PRODUCTION_CONTRACT_SET
    else:
        contracts = None
    if contracts is None:
        error_stream.write(f"WPS_DISCOVERY_UNAVAILABLE app={args.app}\n")
        return 4
    if args.index:
        value = {
            "app": args.app,
            "actions": [entry.to_wire() for entry in contracts.action_index()],
        }
    else:
        value = contracts.batch_resolve([
            ActionAddress(app=args.app, action=action) for action in args.resolve
        ])
    output_stream.write(json.dumps(value, ensure_ascii=True, indent=2) + "\n")
    return 0 if args.index or value["status"] == "complete" else 2


def _production_session_factory(
    *,
    application,
    session_id,
    debug_close_created_document=False,
):
    if application == "word":
        from wps_skills.word.windows.session import build_session

        return build_session(
            session_id=session_id,
            debug_close_created_document=debug_close_created_document,
        )
    if application == "excel":
        from wps_skills.excel.windows.session import build_session

        if debug_close_created_document:
            raise SessionStartupError("Excel only opens existing workbooks; debug document cleanup is unavailable")
        return build_session(session_id=session_id)
    if application == "ppt":
        from wps_skills.ppt.windows.session import build_session

        if debug_close_created_document:
            raise SessionStartupError("PPT only opens existing presentations; debug cleanup is unavailable")
        return build_session(session_id=session_id)
    raise SessionStartupError(
        f"no production {application} Application Contract Set is installed"
    )


def main(
    argv=None,
    *,
    input_stream=None,
    output_stream=None,
    error_stream=None,
):
    parser = _parser()
    args = parser.parse_args(argv)
    if args.debug_close_created_document and not args.session:
        parser.error("--debug-close-created-document requires --session")
    managed = args.start or args.call or args.status or args.close
    if not math.isfinite(args.timeout) or args.timeout <= 0:
        parser.error("--timeout must be a positive finite number")
    if args.call and (not args.action or args.step is None):
        parser.error("--call requires --action and --step")
    if args.step is not None and (args.step < 1 or not (args.call or args.status)):
        parser.error("--step must be positive and requires --call or --status")
    if not args.call and (args.action or args.params_json is not None or args.params_file or args.params_stdin):
        parser.error("--action and parameter options require --call")
    input_stream = input_stream or sys.stdin
    output_stream = output_stream or sys.stdout
    error_stream = error_stream or sys.stderr
    if managed:
        from wps_skills.client import managed_session
        try:
            if importlib.util.find_spec("wps_skills." + args.app) is None:
                raise ValueError("Requested application is not installed in this Skill")
            if args.start:
                value = managed_session.start(args.app, timeout=args.timeout)
            elif args.call:
                raw = (Path(args.params_file).read_text(encoding="utf-8-sig") if args.params_file
                       else input_stream.read() if args.params_stdin
                       else args.params_json if args.params_json is not None else "{}")
                value = managed_session.call(args.call, args.app, step=args.step,
                                             action=args.action, params=managed_session.decode(raw), timeout=args.timeout)
            elif args.status:
                value = managed_session.status(args.status, args.app, step=args.step)
            else:
                value = managed_session.close(args.close, args.app, timeout=args.timeout)
        except (OSError, ValueError) as exc:
            value = {"error": {"code": "CLI_REQUEST_FAILED", "message": str(exc)}}
        output_stream.write(json.dumps(value, ensure_ascii=True, allow_nan=False) + "\n")
        return managed_session.exit_code(value, action_result=bool(args.call or (args.status and args.step)))
    if not args.session:
        return _discover(args, output_stream, error_stream)
    trace_journal = JsonlTraceJournal.default()
    host = SessionHost(
        session_factory=lambda **kwargs: _production_session_factory(
            **kwargs,
            debug_close_created_document=(
                args.debug_close_created_document
            ),
        ),
        session_id_factory=lambda: f"session-{uuid.uuid4().hex}",
        trace_log_factory=trace_journal.session_log,
        action_trace_factory=trace_journal.action_trace,
        session_event_sink=trace_journal.session_event,
    )
    return host.serve(
        application=args.app,
        input_stream=input_stream,
        output_stream=output_stream,
        error_stream=error_stream,
    )


if __name__ == "__main__":
    raise SystemExit(main())
