"""Submit one immutable application plan and retain its execution facts without replay."""

import importlib
import json

from wps_skills.client.task_store import exit_code
from wps_skills.client.json_io import _read, _write, decode
from wps_skills.client import task_request, task_store
from wps_skills.core import timing
from wps_skills.core.action_runtime import DEFERRED_VALUE, _json_equal


from wps_skills.client.applications import contracts_for


def _stop(result, step_id, phase, code, message, *, outcome="failed", state="stopped"):
    result.update(state=state, outcome=outcome,
                  stop={"stepId": step_id, "phase": phase, "error": {"code": code, "message": message}})


def rejected(request, application, error):
    result = {"type": "task.response", "app": application, "taskId": None,
              "document": None, "steps": [], "completion": {"save": None, "pdf": None},
              "cleanup": {"outcome": None, "error": None}, "recordPath": None}
    if isinstance(request, dict) and isinstance(request.get("steps"), list):
        result["steps"] = [
            {"id": step.get("id"), "address": step.get("address"),
             "state": "not_executed", "response": None}
            for step in request["steps"][:task_request.MAX_STEPS] if isinstance(step, dict)
        ]
    _stop(result, error.step_id, "validation", error.code, str(error), state="rejected")
    return result


@timing.timed("task.preflight")
def _preflight(plan, contracts):
    previous = set()
    for step in plan["steps"]:
        name = step["address"]["action"]
        if contracts.resolve(name) is None:
            raise task_request.TaskRequestError("Unknown Action: " + name, step_id=step["id"])
        try:
            if task_request.validate_value(step["params"], previous):
                def deferred(value):
                    if isinstance(value, dict):
                        if "$ref" in value:
                            return DEFERRED_VALUE
                        return {key: deferred(child) for key, child in value.items()}
                    if isinstance(value, list):
                        return [deferred(child) for child in value]
                    return value
                contracts.validate_partial_params(name, deferred(step["params"]))
            else:
                contracts.validate_params(name, step["params"])
        except ValueError as exc:
            raise task_request.TaskRequestError(str(exc), step_id=step["id"])
        previous.add(step["id"])


@timing.timed("task.client")
def execute(request, application, *, timeout=60, executor_factory=None, contracts=None,
            progress=None, on_admitted=None, source_path=None):
    """Admit one application Task; factories are internal dependencies, never JSON."""
    try:
        with timing.span("task.validate_request"):
            request = decode(json.dumps(request, allow_nan=False))
            from wps_skills.client.applications import compile_request
            plan = compile_request(request, application)
        if source_path is None:
            raise task_request.TaskRequestError("Task requires --task-file for submission identity")
        contracts = contracts if contracts is not None else contracts_for(application)
        _preflight(plan, contracts)
    except (ValueError, TypeError, RecursionError) as exc:
        error = exc if isinstance(exc, task_request.TaskRequestError) else task_request.TaskRequestError(str(exc))
        return rejected(request, application, error)
    path = task_store.directory(application, task_store.input_identity(source_path))
    path.mkdir(mode=0o700, parents=True, exist_ok=True)
    with timing.span("task.admission_lock"):
        lock = task_store.acquire(path)
    try:
        request_path = path / "request.json"
        if request_path.exists():
            if not _json_equal(_read(request_path), request):
                return rejected(request, application, task_request.TaskRequestError(
                    "Input path already belongs to a different admitted request; it cannot be replaced",
                    code="TASK_INPUT_CONFLICT"))
            result = task_store.snapshot(path, owner_alive=lock is None)
            if on_admitted is not None and (path / "response.json").is_file():
                with timing.span("task.input_consumption"):
                    on_admitted(path)
            return result
        if lock is None:
            return rejected(request, application, task_request.TaskRequestError(
                "Task admission is in progress; query the same input path", code="TASK_ADMISSION_BUSY"))
        with timing.span("task.admit_request"):
            _write(request_path, request, once=True)
        result = task_store.initial(request, path, plan=plan)
        timing.bind_task(result["taskId"])
        with timing.span("task.receipt_publish"):
            task_store.publish(path, result)
        if on_admitted is not None:
            with timing.span("task.input_consumption"):
                on_admitted(path)
        return _execute(request, plan, result, path, timeout, executor_factory, contracts, progress)
    finally:
        if lock is not None:
            lock.close()


@timing.timed("task.execution")
def _execute(request, plan, result, path, timeout, executor_factory, contracts, progress):
    executor = current = None
    phase = "startup"
    responses = {}

    def publish():
        with timing.span("task.receipt_publish"):
            task_store.publish(path, result)
        if progress is not None:
            try:
                with timing.span("response.progress_write"):
                    progress({"type": "task.progress", "taskId": result["taskId"],
                              "state": result["state"], "stepId": current["id"] if current else None,
                              "stepState": current["state"] if current else None})
            except (OSError, ValueError):
                pass

    try:
        if executor_factory is None:
            factory_module = importlib.import_module("wps_skills." + result["app"] + ".windows.task_factory")
            executor_factory = factory_module.build_task
        with timing.span("task.resources_startup"):
            executor = executor_factory(task_id=result["taskId"], contracts=contracts,
                                        action_timeout_seconds=timeout)
        for step, current in zip(plan["steps"], task_request.result_steps(result)):
            phase = "preparation"
            if not executor.can_execute:
                raise task_request.TaskRequestError("Task resources cannot continue", code="TASK_EXECUTION_UNAVAILABLE")
            with timing.span("action.resolve", stepId=current["id"], action=step["address"]["action"]):
                params = task_request.resolve(step["params"], responses)
            try:
                with timing.span("action.validate", stepId=current["id"], action=step["address"]["action"]):
                    contracts.validate_params(step["address"]["action"], params)
            except ValueError as exc:
                raise task_request.TaskRequestError(str(exc), code="TASK_PARAMS_INVALID")
            current["state"] = "running"
            publish()  # Durable possible-dispatch intent before the native call.
            phase = "execution"
            with timing.fields(stepId=current["id"], action=step["address"]["action"]):
                with timing.span("action.execute") as measured:
                    response = executor.execute(step["address"], params)
                    measured.update(outcome=response.get("outcome"), traceId=response.get("traceId"))
            current.update(state=response["outcome"], response=response)
            publish()  # Confirmed effects survive failures in later work or cleanup.
            if response["outcome"] != "succeeded":
                _stop(result, step["id"], phase, response["error"]["code"], response["error"]["message"],
                      outcome=response["outcome"])
                break
            responses[step["id"]] = response
            if current is result["document"]:
                from wps_skills.client.task_plan import check_document
                phase = "document"
                with timing.span("task.document_check"):
                    check_document(request, response)
        else:
            result.update(state="completed", outcome="succeeded")
    except (Exception, KeyboardInterrupt) as exc:
        if executor is None and getattr(exc, "cleanup", None) is not None:
            result["cleanup"] = exc.cleanup
        unknown = current is not None and current["state"] == "running" and phase == "execution"
        if current is not None and current["state"] == "running":
            current["state"] = "unknown" if unknown else "not_executed"
        code = (exc.code if isinstance(exc, task_request.TaskRequestError)
                else "TASK_INTERRUPTED" if isinstance(exc, KeyboardInterrupt) else "TASK_EXECUTION_FAILED")
        _stop(result, getattr(exc, "step_id", None) or (current["id"] if current else None),
              phase, code, str(exc) or type(exc).__name__, outcome="unknown" if unknown else "failed")
    finally:
        final_state = result["state"]
        if executor is not None:
            result["state"] = "closing"
            try:
                publish()
            except OSError as exc:
                result["recordError"] = str(exc)
            try:
                with timing.span("task.cleanup") as measured:
                    result["cleanup"] = executor.close()
                    measured["outcome"] = result["cleanup"].get("outcome")
            except (Exception, KeyboardInterrupt) as exc:
                result["cleanup"] = {"outcome": "unknown", "error": str(exc) or type(exc).__name__}
        result["state"] = final_state
        try:
            publish()
        except OSError as exc:
            result["recordError"] = str(exc)
    return result
