"""Shared UTF-8 Task submission and receipt command I/O."""

import importlib.util
import json
from pathlib import Path
from wps_skills.core import timing


def run(args, *, task_client, consume, timing_enabled, input_stream, output_stream, error_stream):
    with timing.submission(timing_enabled):
        from wps_skills.client import task_store
        from wps_skills.client.json_io import decode
        from wps_skills.client.task_request import TaskRequestError
        request = {"taskId": args.task_status}
        source_file = None
        try:
            if importlib.util.find_spec("wps_skills." + args.app) is None:
                raise ValueError("Requested application is not installed in this Skill")
            if args.task_status_file:
                value = task_store.status_file(args.app, args.task_status_file)
            elif args.task_status:
                value = task_store.status(args.app, args.task_status)
            elif args.task_file and not Path(args.task_file).exists():
                # Consumption may precede a lost command response. The original
                # locator is sufficient to retrieve evidence, never to restart.
                value = task_store.status_file(args.app, args.task_file)
            else:
                limit = 8 * 1024 * 1024
                if args.task_file:
                    from wps_skills.client.task_file import ConsumableTaskFile
                    source_file = ConsumableTaskFile(args.task_file)
                    with timing.span("input.read"):
                        raw = source_file.read(limit)
                else:
                    raise ValueError("Task submission requires an input file")
                if len(raw.encode("utf-8")) > limit:
                    raise ValueError("Task Request exceeds 8 MiB")

                def progress(event):
                    error_stream.write(json.dumps(event, ensure_ascii=True) + "\n")
                    error_stream.flush()

                with timing.span("input.decode"):
                    request = decode(raw)
                value = task_client.execute(request, args.app, timeout=args.timeout, progress=progress,
                                            source_path=args.task_file,
                                            on_admitted=source_file.admitted if source_file and consume else None)
        except (OSError, ValueError, RecursionError) as exc:
            value = task_client.rejected(request, args.app, TaskRequestError(str(exc)))
        if source_file is not None and consume:
            value = dict(value, taskFile=dict(source_file.report))
        timing.bind_task(value.get("taskId"))
        with timing.span("response.encode"):
            encoded = json.dumps(value, ensure_ascii=True, allow_nan=False) + "\n"
        with timing.span("response.write_flush", outcome=value.get("outcome")):
            output_stream.write(encoded)
            output_stream.flush()
        return task_client.exit_code(value)
