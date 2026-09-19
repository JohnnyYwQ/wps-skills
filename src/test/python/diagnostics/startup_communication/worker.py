"""Detached Windows supervisor. One run identity is never restarted or replayed."""
import argparse
import ctypes
import json
import os
from pathlib import Path
import platform
import struct
import subprocess
import sys
import time

from common import Journal, file_hashes, read_json, utc, validate_config, verify_bundle, write_json


def supervise(command, directory, timeout, journal):
    with (directory / "stdout.txt").open("w", encoding="utf-8") as stdout, (directory / "stderr.txt").open("w", encoding="utf-8") as stderr:
        with journal.phase("python.process_create", caseDirectory=directory.name):
            process = subprocess.Popen(command, stdout=stdout, stderr=stderr,
                                       env=dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8"))
        journal.event("python.created", childPid=process.pid, caseDirectory=directory.name)
        try:
            return process.wait(timeout=timeout), False
        except subprocess.TimeoutExpired:
            # Only this test child. Its owned bridge Job closes with its owner.
            # WPS itself is never killed, and possible effects remain uncertain.
            journal.event("case.watchdog_expired", childPid=process.pid, caseDirectory=directory.name)
            process.kill()
            process.wait(timeout=10)
            return process.returncode, True


def environment():
    if os.name != "nt":
        raise RuntimeError("Windows interactive desktop required")
    session = ctypes.c_ulong()
    if not ctypes.windll.kernel32.ProcessIdToSessionId(os.getpid(), ctypes.byref(session)) or not session.value:
        raise RuntimeError("Session 0 is not a valid WPS test environment")
    return {"python": sys.version, "executable": sys.executable, "pythonBits": struct.calcsize("P") * 8,
            "platform": platform.platform(), "interactiveSessionId": session.value,
            "candidate": "baseline-anonymous-pipe", "instrumented": True}


def run(root, profile):
    root = root.resolve()
    bundle = root / "bundle"
    evidence = root / "windows"
    evidence.mkdir(exist_ok=True)
    # Exclusive admission persists after exit/crash: never reuse this identity.
    with (root / "run.claim").open("x", encoding="utf-8") as stream:
        stream.write(str(os.getpid()))
    journal = Journal(evidence / "events.jsonl", runId=root.name, clockDomain="windows-supervisor")
    status = {"runId": root.name, "state": "running", "profile": profile, "startedUtc": utc(),
              "cases": [], "activeCase": None, "stopReason": None, "remaining": "not_executed"}
    status.update(observedPassed=0, observedFailed=0, visualReview="not_required", caseHistory="cases.jsonl")

    def publish():
        status["updatedUtc"] = utc()
        write_json(evidence / "status.json", status)

    publish()
    journal.event("supervisor.entry")
    try:
        with journal.phase("environment.validate"):
            observed_environment = environment()
            verify_bundle(bundle)
            config = validate_config(read_json(bundle / "experiment.json"))
            write_json(evidence / "environment.json", observed_environment)
        settings = config["profiles"][profile]
        began = time.monotonic()
        round_number = 0
        while True:
            if (root / "stop.request").exists():
                status.update(state="stopped", stopReason="operator_requested")
                break
            if time.monotonic() - began >= settings["budgetSeconds"]:
                status.update(state="completed", stopReason="time_budget")
                break
            if settings["maxRounds"] is not None and round_number >= settings["maxRounds"]:
                status.update(state="completed", stopReason="round_limit")
                break
            round_number += 1
            for case in config["cases"]:
                if (root / "stop.request").exists():
                    status.update(state="stopped", stopReason="operator_requested")
                    break
                if time.monotonic() - began >= settings["budgetSeconds"]:
                    status.update(state="completed", stopReason="time_budget")
                    break
                name = f"r{round_number:05d}-{case}"
                directory = evidence / "cases" / name
                directory.mkdir(parents=True, exist_ok=False)
                record = {"name": name, "case": case, "round": round_number, "state": "running", "startedUtc": utc()}
                status["cases"].append(record)
                status["cases"] = status["cases"][-50:]
                status["activeCase"] = name
                publish()
                command = [sys.executable, str(bundle / "child.py"), "--bundle", str(bundle), "--directory", str(directory), "--case", case]
                started = time.perf_counter()
                code, timed_out = supervise(command, directory, config["limits"]["caseSeconds"], journal)
                result_path = directory / "result.json"
                result = read_json(result_path) if result_path.exists() else {"state": "unknown", "error": "Child exited without result"}
                if case.startswith("wps"):
                    from com_report import summarize
                    com = summarize(directory)
                    result = dict(result, comConnection=com["connection"], comFailureStage=com["failureStage"], comCoverage=com["coverage"])
                if timed_out:
                    result = {"state": "unknown", "error": "Case watchdog expired; effects may be uncertain; no replay"}
                    write_json(directory / "watchdog.json", result)
                passed = code == 0 and result.get("state") == "passed" and not timed_out
                record.update(state="passed" if passed else "failed", exitCode=code, timedOut=timed_out,
                              durationMs=(time.perf_counter() - started) * 1000, result=result, endedUtc=utc())
                with (evidence / "cases.jsonl").open("a", encoding="utf-8") as history:
                    history.write(json.dumps(record, ensure_ascii=False) + "\n")
                status["observedPassed" if passed else "observedFailed"] += 1
                if result.get("visualReview") == "pending":
                    status["visualReview"] = "pending"
                status["activeCase"] = None
                if not passed:
                    status.update(state="failed", stopReason="unexpected_case_failure")
                publish()
                if not passed:
                    break
            if status["state"] != "running":
                break
        if not status["cases"] and status["state"] == "completed":
            status.update(state="incomplete", stopReason="no_cases_executed")
    except BaseException as error:
        status.update(state="failed", stopReason="supervisor_failure", error=repr(error))
        journal.event("supervisor.failed", error=repr(error))
        import traceback
        (evidence / "supervisor.failure.txt").write_text(traceback.format_exc(), encoding="utf-8")
        traceback.print_exc()
    finally:
        status["endedUtc"] = utc()
        status["conclusion"] = "Observed results only; no comparative benefit or long-term reliability claim."
        journal.event("supervisor.finished", state=status["state"])
        publish()
        # The outer wrapper writes these after this supervisor exits. Their
        # snapshot integrity is covered by the archive hash, not this index.
        hashes = file_hashes(evidence)
        for outer in ("launcher.json", "worker.stdout", "worker.stderr"):
            hashes.pop(outer, None)
        write_json(root / "evidence.sha256.json", hashes)
    return 0 if status["state"] in ("completed", "stopped") else 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", required=True, type=Path)
    parser.add_argument("--profile", required=True, choices=("quick", "long"))
    options = parser.parse_args()
    raise SystemExit(run(options.root, options.profile))
