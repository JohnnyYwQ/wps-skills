"""Task-only Skill CLI, with no Session command interface."""

import argparse
import math
import sys

from wps_skills.cli.discovery import discover
from wps_skills.cli.task_io import run
from wps_skills.client import task_client


def main(argv=None, *, input_stream=None, output_stream=None, error_stream=None,
         required_application="word"):
    parser = argparse.ArgumentParser(description="Submit an application Task, query its receipt or discover Actions")
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--index", action="store_true", help="Print the production Action Index")
    mode.add_argument("--resolve", nargs="+", metavar="ACTION", help="Resolve complete Action Contracts")
    mode.add_argument("--task-file", metavar="PATH", help="Submit a UTF-8 JSON Task Request")
    mode.add_argument("--task-status", metavar="TASK_ID", help="Read a retained Task Response")
    mode.add_argument("--task-status-file", metavar="PATH", help="Read a receipt using its original input path")
    parser.add_argument("--app", choices=("word", "excel", "ppt"), required=True)
    parser.add_argument("--timeout", type=float, default=60, help="Action timeout in seconds (default: 60)")
    args = parser.parse_args(argv)
    if required_application is not None and args.app != required_application:
        parser.error("This Skill only accepts --app " + required_application)
    if not math.isfinite(args.timeout) or args.timeout <= 0:
        parser.error("--timeout must be a positive finite number")
    for name in ("task_file", "task_status", "task_status_file"):
        if getattr(args, name) == "":
            parser.error("--" + name.replace("_", "-") + " requires a non-empty value")
    input_stream = input_stream or sys.stdin
    output_stream = output_stream or sys.stdout
    error_stream = error_stream or sys.stderr
    if args.index or args.resolve:
        return discover(args, output_stream, error_stream)
    return run(args, task_client=task_client, consume=True, timing_enabled=True,
               input_stream=input_stream, output_stream=output_stream, error_stream=error_stream)
