"""Task-local Session Client for separate CLI invocations.

The private mailbox contains JSON data only. Atomic, numbered requests and
durable receipts prevent a repeated CLI invocation from replaying an Action.
The worker owns the existing Protocol v1 client, never a second Session.
"""

import json
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import threading
import time
import uuid

from wps_skills.client.session_client import (
    ActionFailed, SessionClient, SessionClientError, _nonfinite, _unique_object,
)


def decode(value):
    result = json.loads(value, object_pairs_hook=_unique_object, parse_constant=_nonfinite)
    json.dumps(result, allow_nan=False)
    return result


def _read(path):
    return decode(path.read_text(encoding="utf-8-sig"))


def _write(path, value, *, once=False):
    """Publish complete JSON atomically; once=True never replaces a request."""
    encoded = json.dumps(value, ensure_ascii=True, allow_nan=False) + "\n"
    descriptor, name = tempfile.mkstemp(prefix=".write-", dir=str(path.parent))
    temporary = Path(name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
        if once:
            os.link(str(temporary), str(path))
        else:
            # A Windows reader can briefly hold a non-delete-sharing handle.
            # Retry only publication of these same bytes, never an Action.
            deadline = time.monotonic() + 1
            while True:
                try:
                    os.replace(str(temporary), str(path))
                    break
                except PermissionError:
                    if os.name != "nt" or time.monotonic() >= deadline:
                        raise
                    time.sleep(0.01)
    finally:
        temporary.unlink(missing_ok=True)


def _root():
    override = os.environ.get("WPS_SKILLS_SESSION_DIR")
    if override:
        root = Path(override).expanduser().resolve()
    elif os.name == "nt":
        root = Path(os.environ["LOCALAPPDATA"]) / "WpsSkills" / "sessions"
    else:
        root = Path.home() / ".local" / "state" / "wps-skills" / "sessions"
    root.mkdir(mode=0o700, parents=True, exist_ok=True)
    return root


def _directory(handle, application):
    if not re.fullmatch(r"[0-9a-f]{32}", handle):
        raise ValueError("Invalid session handle")
    directory = _root() / handle
    config = _read(directory / "config.json")
    if config["app"] != application:
        raise ValueError("Session belongs to a different application")
    return directory


def _error(code, message, **fields):
    return {"error": {"code": code, "message": message}, **fields}


def _state(directory):
    state = _read(directory / "state.json")
    if state["state"] != "closed":
        heartbeat = directory / "heartbeat"
        if time.time() - heartbeat.stat().st_mtime > 10:
            return {**state, "state": "unavailable", "canExecute": False,
                    **_error("SESSION_CLIENT_LOST", "Session Client stopped responding; do not replay or recreate the document"),
                    "mayHaveEffect": state["state"] == "busy"}
    return state


def _wait(directory, timeout, predicate):
    deadline = time.monotonic() + timeout
    while True:
        state = _state(directory)
        result = predicate(state)
        if result is not None:
            return result
        if state["state"] in {"closed", "unavailable"}:
            return {**state, **_error("SESSION_UNAVAILABLE", "No further Action can be submitted; inspect the retained response and cleanup result")}
        if time.monotonic() >= deadline:
            return {**state, **_error("COMMAND_WAIT_TIMEOUT", "Check --status or repeat this same step to retrieve its receipt; do not submit the Action under a new step"),
                    "mayHaveEffect": state["state"] == "busy"}
        time.sleep(0.05)


def start(application, *, timeout=60, host_command=None, idle_timeout=300):
    """Start one package-owned client. Injectable Host is an internal test seam."""
    handle = uuid.uuid4().hex
    directory = _root() / handle
    directory.mkdir(mode=0o700)
    command = host_command or [sys.executable, "-m", "wps_skills.cli.call", "--session", "--app", application]
    _write(directory / "config.json", {
        "app": application, "hostCommand": command, "timeout": timeout,
        "idleTimeoutSeconds": idle_timeout,
    })
    _write(directory / "state.json", {
        "handle": handle, "app": application, "state": "starting",
        "canExecute": False, "nextStep": 1, "response": None,
        "sessionOutcome": None, "cleanupError": None,
    })
    (directory / "heartbeat").touch()
    env = dict(os.environ, PYTHONPATH=str(Path(__file__).resolve().parents[2]), PYTHONIOENCODING="utf-8")
    flags = subprocess.CREATE_NO_WINDOW | subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0
    try:
        with (directory / "worker.stderr.log").open("wb") as errors:
            process = subprocess.Popen(
                [sys.executable, "-m", "wps_skills.client.managed_session", str(directory)],
                stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=errors,
                env=env, creationflags=flags, start_new_session=os.name != "nt",
            )
        # Reap when embedded in a long-lived Python caller; CLI exit does not
        # wait for this daemon thread or terminate its detached worker.
        threading.Thread(target=process.wait, daemon=True).start()
    except OSError as exc:
        state = {**_read(directory / "state.json"), "state": "closed",
                 **_error("SESSION_STARTUP_FAILED", str(exc))}
        _write(directory / "state.json", state)
        return state
    result = _wait(directory, timeout + 10, lambda state: state if state["state"] != "starting" else None)
    if result.get("error"):
        _write(directory / "close.json", {})
    return result


def status(handle, application, *, step=None):
    directory = _directory(handle, application)
    state = _state(directory)
    if step is not None:
        receipt = directory / ("response-%d.json" % step)
        if receipt.exists():
            return _read(receipt)
        return {**state, **_error("STEP_RESULT_UNAVAILABLE", "No receipt is available for this step")}
    return state


def call(handle, application, *, step, action, params, timeout=60):
    directory = _directory(handle, application)
    if not isinstance(params, dict) or not action or step < 1:
        raise ValueError("call requires an Action name, object params, and a positive step")
    request = {"address": {"app": application, "action": action}, "params": params}
    # Validate before any mailbox mutation, including floating point overflow.
    json.dumps(request, allow_nan=False)
    request_path = directory / ("request-%d.json" % step)
    state = _state(directory)
    if not request_path.exists():
        if not state["canExecute"] or state["state"] != "ready" or step != state["nextStep"]:
            return {**state, **_error("STEP_NOT_AVAILABLE", "Read --status and the previous response before selecting the next step")}
        try:
            _write(request_path, request, once=True)
        except FileExistsError:
            pass
    if json.dumps(_read(request_path), sort_keys=True) != json.dumps(request, sort_keys=True):
        return {**state, **_error("STEP_CONFLICT", "This step already names a different Action Request; it cannot be replaced")}
    receipt = directory / ("response-%d.json" % step)
    result = _wait(directory, timeout, lambda state: _read(receipt) if receipt.exists() else None)
    if result.get("error") and not receipt.exists():
        result["mayHaveEffect"] = True
    return result


def close(handle, application, *, timeout=60):
    directory = _directory(handle, application)
    _write(directory / "close.json", {})
    return _wait(directory, timeout, lambda state: state if state["state"] in {"closed", "unavailable"} else None)


def exit_code(result, *, action_result=True):
    if result.get("error") or result.get("cleanupError") or (result.get("sessionOutcome") or {}).get("outcome") == "failed":
        return 4
    response = result.get("response")
    return 2 if action_result and response and response["outcome"] != "succeeded" else 0


def serve(directory):
    """Own one Host until close, terminal failure, or bounded idle expiry."""
    config = _read(directory / "config.json")
    state = _read(directory / "state.json")
    client = SessionClient(config["hostCommand"], application=config["app"], timeout=config["timeout"])
    stop = threading.Event()

    def heartbeat():
        while not stop.wait(1):
            (directory / "heartbeat").touch()

    thread = threading.Thread(target=heartbeat, daemon=True)
    thread.start()

    def publish():
        state.update(sessionOutcome=client.session_outcome,
                     cleanupError=str(client.cleanup_error) if client.cleanup_error else None)
        _write(directory / "state.json", state)

    try:
        client.start()
        state.update(state="ready", canExecute=True, ready=client.ready,
                     clientPid=os.getpid(), idleTimeoutSeconds=config["idleTimeoutSeconds"])
        publish()
        last_action = time.monotonic()
        while client.can_execute:
            if (directory / "close.json").exists():
                state["closeReason"] = "client_close"
                break
            if time.monotonic() - last_action >= config["idleTimeoutSeconds"]:
                state["closeReason"] = "idle_timeout"
                break
            step = state["nextStep"]
            request_path = directory / ("request-%d.json" % step)
            if not request_path.exists():
                time.sleep(0.05)
                continue
            request = _read(request_path)
            state.update(state="busy", canExecute=False, step=step, response=None)
            publish()
            try:
                state["response"] = client.call(request["address"], request["params"])
            except ActionFailed as exc:
                state["response"] = exc.response
            except SessionClientError as exc:
                # last_response may refer to an earlier Action. Only preserve a
                # current authoritative response when this call received one.
                if not exc.may_have_effect and client.last_response is not None:
                    previous = directory / ("response-%d.json" % (step - 1))
                    old = _read(previous).get("response") if previous.exists() else None
                    if client.last_response != old:
                        state["response"] = client.last_response
                state.update(_error("SESSION_CHANNEL_FAILED", str(exc)), mayHaveEffect=exc.may_have_effect)
            state.update(state="ready" if client.can_execute else "closing",
                         canExecute=client.can_execute, nextStep=step + 1)
            state.update(sessionOutcome=client.session_outcome,
                         cleanupError=str(client.cleanup_error) if client.cleanup_error else None)
            # Receipt commits before another step is advertised. Crash in between
            # leaves a retrievable result, never an executable duplicate.
            _write(directory / ("response-%d.json" % step), state)
            publish()
            last_action = time.monotonic()
    except BaseException as exc:
        state.update(_error("SESSION_CLIENT_FAILED", str(exc)), canExecute=False,
                     mayHaveEffect=state["state"] == "busy")
    finally:
        try:
            client.close()
        except SessionClientError:
            pass
        state.update(state="closed", canExecute=False, stderr=client.stderr)
        publish()
        stop.set()
        thread.join(timeout=2)


if __name__ == "__main__":
    serve(Path(sys.argv[1]))
