"""Durable Task receipts and a process-owned lock; never resume an old Task."""

import os
import hashlib
from pathlib import Path

from wps_skills.client.json_io import _read, _write
from wps_skills.client.task_request import _identifier
from wps_skills.client import task_request


def input_identity(source_path):
    """An input locator names one submission, even after consumption.

    Using the complete digest avoids a separate admission index and its crash
    window. A new input path denotes a new Task; reusing one never permits replay.
    """
    locator = os.path.normcase(str(Path(source_path).expanduser().resolve()))
    return "task-" + hashlib.sha256(locator.encode("utf-8")).hexdigest()


def status_file(application, source_path):
    result = status(application, input_identity(source_path))
    return result


def directory(application, task_id):
    _identifier(task_id)
    root = os.environ.get("WPS_SKILLS_TASK_DIR")
    if root:
        root = Path(root).expanduser().resolve()
    elif os.name == "nt":
        root = Path(os.environ["LOCALAPPDATA"]) / "WpsSkills" / "tasks"
    else:
        root = Path.home() / ".local" / "state" / "wps-skills" / "tasks"
    return root / application / task_id


def acquire(path):
    """Lock byte zero on Windows; the OS releases ownership on process death."""
    stream = (path / "owner.lock").open("a+b")
    stream.seek(0)
    try:
        if os.name == "nt":
            import msvcrt
            msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        stream.close()
        return None
    return stream


def initial(request, path, *, plan=None):
    if "version" in request:
        return {"type": "task.response", "app": request.get("app"), "taskId": path.name,
                "state": "interrupted", "outcome": "unknown", "steps": [],
                "stop": {"stepId": None, "phase": "execution", "error": {
                    "code": "HISTORICAL_TASK_RECEIPT_UNAVAILABLE",
                    "message": "Historical request retained without a response; do not replay"}},
                "cleanup": {"outcome": "unknown", "error": "No final receipt"},
                "recordPath": str(path / "response.json")}
    from wps_skills.client.applications import compile_request
    from wps_skills.client.task_plan import SAVE_STEP, PDF_STEP
    if plan is None:
        # Recovery without a receipt has no previously compiled in-memory plan.
        plan = compile_request(request, request["app"])
    records = [
        {"id": step["id"], "address": step["address"],
         "state": "not_executed", "response": None}
        for step in plan["steps"]
    ]
    by_id = {step["id"]: step for step in records}
    return {
        "type": "task.response", "taskId": path.name, "app": request["app"],
        "state": "running", "outcome": None, "stop": None,
        "document": records[0], "steps": records[1:1 + len(request["steps"])],
        "completion": {"save": by_id.get(SAVE_STEP), "pdf": by_id.get(PDF_STEP)},
        "cleanup": {"outcome": None, "error": None},
        "recordPath": str(path / "response.json"),
    }


def snapshot(path, *, owner_alive):
    response = path / "response.json"
    result = _read(response) if response.exists() else initial(_read(path / "request.json"), path)
    if result["state"] in {"running", "closing"} and not owner_alive:
        phase = "cleanup" if result["state"] == "closing" else "execution"
        result.update(state="interrupted", outcome=result["outcome"] or "unknown")
        result["stop"] = result["stop"] or {
            "stepId": None, "phase": phase,
            "error": {"code": "TASK_OWNER_LOST", "message": "Task process ended without a final receipt; do not replay"},
        }
        for step in task_request.result_steps(result):
            if step["state"] == "running":
                step["state"] = "unknown"
                result["stop"]["stepId"] = step["id"]
        if "cleanup" in result:
            result["cleanup"] = {"outcome": "unknown", "error": "Task resource cleanup was not confirmed"}
        else:
            result["cleanupError"] = "Session cleanup was not confirmed"
    return result


def status(application, task_id):
    path = directory(application, task_id)
    if not (path / "request.json").is_file():
        raise ValueError("No Task receipt exists for this taskId")
    lock = acquire(path)
    try:
        return snapshot(path, owner_alive=lock is None)
    finally:
        if lock is not None:
            lock.close()


def publish(path, result):
    _write(path / "response.json", result)


def exit_code(result):
    if (result.get("recordError") or result.get("cleanupError")
            or (result.get("cleanup") or {}).get("outcome") in {"failed", "unknown"}
            or (result.get("sessionOutcome") or {}).get("outcome") == "failed"
            or result["state"] in {"running", "closing", "rejected", "interrupted"}):
        return 4
    return 0 if result["outcome"] == "succeeded" else 2
