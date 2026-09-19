"""Application-local Task execution and resource ownership."""

from dataclasses import asdict
import math
import threading
from contextvars import copy_context

from wps_skills.core import timing

from wps_skills.core.action_runtime import ActionAddress, ActionRequest, ExecutionResources, _thaw_json
from wps_skills.core.trace_journal import JsonlTraceJournal


class ApplicationTask:
    def __init__(self, *, application, task_id, contracts, adapter, coordinator, launcher=None,
                 action_timeout_seconds=60, cleanup_timeout_seconds=5, journal=None):
        if not math.isfinite(action_timeout_seconds) or action_timeout_seconds <= 0:
            raise ValueError("Action timeout must be positive and finite")
        self._action_timeout_seconds = action_timeout_seconds
        self.task_id = task_id
        self.application = application
        self.can_execute = True
        self._journal = journal if journal is not None else JsonlTraceJournal.default()
        self._resources = ExecutionResources(
            application=application, error_scope="TASK", contracts=contracts, adapter=adapter, coordinator=coordinator,
            launcher=launcher, action_timeout_seconds=action_timeout_seconds,
            cleanup_timeout_seconds=cleanup_timeout_seconds,
        )
        self._journal.task_event(task_id=task_id, application=application, event="task.resources_acquired")

    def execute(self, address, params):
        trace = self._journal.task_action_trace(task_id=self.task_id, application=self.application)
        request = ActionRequest(ActionAddress(**address), params)
        completed = threading.Event()
        result = {}
        def dispatch():
            try:
                with timing.fields(traceId=trace.trace_id):
                    with timing.span("action.runtime"):
                        result["turn"] = self._resources.execute(request, trace)
            except BaseException as exc:
                result["error"] = exc
            finally:
                completed.set()
        context = copy_context()
        threading.Thread(target=lambda: context.run(dispatch), daemon=True).start()
        if not completed.wait(self._action_timeout_seconds):
            self.can_execute = False
            trace.event("action.timeout", address=address)
            return {"outcome": "unknown", "address": address, "taskId": self.task_id,
                    "traceId": trace.trace_id, "traceLog": str(trace.trace_log) if trace.trace_log else None,
                    "error": {"code": "TASK_ACTION_TIMEOUT", "message": "Action deadline elapsed; effects may exist, do not replay"}}
        if "error" in result:
            raise result["error"]
        turn = result["turn"]
        self.can_execute = turn.continuation == "continue"
        disposition = turn.disposition
        response = {"outcome": disposition.outcome, "address": address, "taskId": self.task_id,
                    "traceId": trace.trace_id, "traceLog": str(trace.trace_log) if trace.trace_log else None}
        if disposition.error is not None:
            response["error"] = asdict(disposition.error)
        else:
            response["data"] = _thaw_json(disposition.data)
        return response

    def close(self):
        self.can_execute = False
        outcome = self._resources.close()
        result = {"outcome": outcome.outcome, "error": asdict(outcome.error) if outcome.error else None,
                  "resources": asdict(outcome.cleanup)}
        self._journal.task_event(task_id=self.task_id, application=self.application, event="task.resources_released", cleanup=result)
        return result
