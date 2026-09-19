"""Combine native COM observations and durable Task evidence without inventing coverage."""
import json
from pathlib import Path
from common import read_json, write_json

NATIVE_STAGES = ("environment", "com.registration", "com.attach", "com.activate", "document.bind_create", "document.bind_open")


def summarize(directory):
    directory = Path(directory)
    stages = {name: {"state": "not_executed"} for name in NATIVE_STAGES}
    observations = []
    malformed = []
    for path in sorted((directory / "com").glob("*.jsonl")):
        for line_number, line in enumerate(path.read_text(encoding="utf-8-sig").splitlines(), 1):
            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                malformed.append({"file": path.name, "line": line_number})
                continue
            observations.append(event)
            name = event.get("stage")
            if name in stages:
                if event.get("event") == "stage.started":
                    stages[name] = {"state": "unknown", "detail": event}
                elif event.get("event") == "stage.finished":
                    stages[name] = {"state": event["outcome"], "detail": event}
    response = read_json(directory / "response.json") if (directory / "response.json").exists() else None
    for name, identifier in (("document.write", "write"), ("document.readback", "after")):
        step = next((s for s in (response or {}).get("steps", []) if s["id"] == identifier), None)
        stages[name] = {"state": step["state"] if step else "not_executed", "detail": step}
    stages["artifact.verify"] = {"state": "not_executed"}
    if (directory / "events.jsonl").exists():
        for line in (directory / "events.jsonl").read_text(encoding="utf-8").splitlines():
            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                continue
            if event.get("name") == "wps.artifact_verify":
                stages["artifact.verify"] = {"state": event.get("outcome", "unknown"), "detail": event}
    if (directory / "artifact-verification.json").exists():
        stages["artifact.verify"] = {"state": "succeeded", "detail": read_json(directory / "artifact-verification.json")}
    attachment = stages["com.attach"]["state"]
    activation = stages["com.activate"]["state"]
    connection = "existing" if attachment == "succeeded" else "activation" if activation == "succeeded" else "unavailable"
    info = next((e.get("data") for e in reversed(observations) if e.get("stage") == "com.application_info"), None)
    input_request = read_json(directory / "input.json") if (directory / "input.json").exists() else {}
    binding = "document.bind_open" if input_request.get("document", {}).get("address", {}).get("action") == "openWorkbook" else "document.bind_create"
    required = ["environment", "com.registration", binding, "document.write", "document.readback", "artifact.verify"]
    # A failed attach followed by successful activation remains visible, but it
    # is not itself a failed Task. An unobserved alternative is never a pass.
    failed = next((name for name in required if stages[name]["state"] in ("failed", "unknown")), None)
    if connection == "unavailable" and (attachment != "not_executed" or activation != "not_executed"):
        failed = "com.activate" if activation != "not_executed" else "com.attach"
    diagnostics_complete = not malformed and connection != "unavailable" and all(stages[n]["state"] == "succeeded" for n in required)
    summary = {"application": "excel", "connection": connection, "applicationInfo": info,
               "stages": stages, "failureStage": failed, "diagnosticsComplete": diagnostics_complete,
               "malformedRecords": malformed, "observations": observations,
               "coverage": {"existing": attachment == "succeeded", "activation": activation == "succeeded",
                            "coldStartProven": False},
               "taskStop": (response or {}).get("stop"),
               "note": "Activation is a COM API path, not proof of a fresh WPS process. No cross-machine reliability claim."}
    write_json(directory / "com-summary.json", summary)
    return summary


def print_summary(summary, label=""):
    print(f"COM {label}: connection={summary['connection']} complete={summary['diagnosticsComplete']} failure={summary['failureStage'] or '-'}")
    environment = summary["stages"]["environment"].get("detail", {}).get("data") or {}
    if environment:
        print(f"  PowerShell {environment.get('powershellVersion')} / {environment.get('processBits')} bit / session {environment.get('sessionId')} / elevated={environment.get('elevated')} / {environment.get('apartmentState')}")
    if summary.get("applicationInfo"):
        print("  WPS version: " + str(summary["applicationInfo"].get("version")))
    for name, stage in summary["stages"].items():
        print(f"  {name:25} {stage['state']}")
        for error in (stage.get("detail") or {}).get("errors", []):
            print("    " + error["hresult"] + " " + error["type"] + ": " + error["message"])
        if name == "com.registration" and stage["state"] == "failed":
            data = (stage.get("detail") or {}).get("data") or {}
            print("    " + str(data.get("reason", "")) + " " + str(data.get("message", "")))
