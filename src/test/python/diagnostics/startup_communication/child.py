"""One supervised experiment case; imports the isolated production package."""
import argparse
import io
import json
import os
from pathlib import Path
import posixpath
import shutil
import subprocess
import sys
import time
import uuid
import xml.etree.ElementTree as ET
import zipfile

from common import Journal, digest, read_json, write_json


def verify_artifact(path, expected):
    """Inspect OOXML independently of the Action response and COM readback."""
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    office_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    with zipfile.ZipFile(path) as archive:
        book = ET.fromstring(archive.read("xl/workbook.xml"))
        first = book.find("m:sheets/m:sheet", ns)
        relationship = first.attrib["{" + office_ns + "}id"]
        links = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        target = next(r.attrib["Target"] for r in links.findall("{" + rel_ns + "}Relationship")
                      if r.attrib["Id"] == relationship)
        sheet_path = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            strings = ["".join(s.itertext()) for s in ET.fromstring(archive.read("xl/sharedStrings.xml")).findall("m:si", ns)]
        sheet = ET.fromstring(archive.read(sheet_path))
        actual = {}
        for cell in sheet.findall(".//m:sheetData/m:row/m:c", ns):
            address = cell.attrib["r"]
            if address not in expected:
                continue
            text = cell.findtext("m:v", default="", namespaces=ns)
            kind = cell.get("t")
            if kind == "s":
                value = strings[int(text)]
            elif kind == "inlineStr":
                value = "".join(cell.find("m:is", ns).itertext())
            elif kind in (None, "n"):
                value = float(text)
            else:
                value = text
            actual[address] = value
        if actual != expected:
            raise AssertionError({"expected": expected, "actual": actual})
        return {"method": "independent OOXML inspection", "sheet": first.attrib["name"], "cells": actual}


def make_request(output, cells, input_path=None):
    def action(name, **params):
        return {"address": {"app": "excel", "action": name}, "params": params}

    def ref(step, *path):
        return {"$ref": {"step": step, "path": ["data", *path]}}

    sheet = ref("worksheets", "worksheets", 0, "name")
    return {
        "app": "excel", "document": action("openWorkbook", path=str(input_path)) if input_path else action("createWorkbook"),
        "steps": [
            dict(action("listWorksheets", offset=0, limit=100), id="worksheets"),
            dict(action("readRange", sheet=sheet, address="A1:B1"), id="before"),
            dict(action("writeRange", sheet=sheet, address="A1:B1", expectedToken=ref("before", "token"),
                        values=[[cells["A1"], cells["B1"]]]), id="write"),
            dict(action("readRange", sheet=sheet, address="A1:B1"), id="after"),
        ],
        "completion": [action("saveAs", outputPath=str(output), overwritePolicy="failIfExists")],
    }


def install_measurements(bundle, directory, config, journal):
    from wps_skills.windows import owned_process, powershell_bridge
    from wps_skills.core import timing
    start = owned_process.WindowsOwnedProcessLauncher.start_bridge

    def measured_start(self, command):
        with journal.phase("bridge.process_create"):
            process = start(self, command)
        journal.event("bridge.created", childPid=process.pid)
        return process

    owned_process.WindowsOwnedProcessLauncher.start_bridge = measured_start
    original = powershell_bridge.JsonLineBridgeTransport

    class MeasuredTransport(original):
        def __init__(self, **kwargs):
            super().__init__(**kwargs)
            nonce = uuid.uuid4().hex
            request = {"requestId": nonce, "traceId": nonce, "operation": "__diagnostic_ready", "arguments": {"nonce": nonce}}
            with journal.phase("bridge.ready_handshake", childPid=kwargs["process"].pid):
                with timing.span("diagnostic.bridge_ready"):
                    response = self.exchange(request, time.monotonic() + config["limits"]["startupSeconds"])
                    if response != {"requestId": nonce, "outcome": "succeeded", "data": {"nonce": nonce}}:
                        raise AssertionError("Invalid readiness response")

        def exchange(self, request, deadline_at):
            with journal.phase("bridge.request", operation=request.get("operation"), bridgeRequestId=request.get("requestId")):
                return super().exchange(request, deadline_at)

    # Production task_factory imports this symbol when building each Task.
    powershell_bridge.JsonLineBridgeTransport = MeasuredTransport
    return MeasuredTransport


def pipe_case(bundle, directory, config, journal, transport_type):
    from wps_skills.windows.owned_process import WindowsOwnedProcessLauncher
    from wps_skills.windows.bridge_runtime import windows_powershell_executable
    from wps_skills.core import timing
    launcher = WindowsOwnedProcessLauncher()
    transport = None
    cleanup = None
    try:
        with timing.submission(application="excel"):
            process = launcher.start_bridge([windows_powershell_executable(), "-NoLogo", "-NoProfile", "-NonInteractive",
                                              "-ExecutionPolicy", "Bypass", "-File", str(bundle / "probe.ps1")])
            transport = transport_type(process=process)
            messages = Journal(directory / "messages.jsonl")
            for index, size in enumerate(config["payloadCharacters"]):
                seed = '中文🙂\\\"\n\r\t pipe | $ ; '
                payload = (seed * (size // len(seed) + 1))[:size]
                identifier = uuid.uuid4().hex
                request = {"requestId": identifier, "traceId": identifier, "operation": "__diagnostic_echo",
                           "arguments": {"payload": payload, "sequence": index}}
                messages.event("request", record=request)
                response = transport.exchange(request, time.monotonic() + config["limits"]["requestSeconds"])
                messages.event("response", record=response)
                if response != {"requestId": identifier, "outcome": "succeeded", "data": request["arguments"]}:
                    raise AssertionError("Echo mismatch at message " + str(index))
            with journal.phase("bridge.graceful_exit"):
                if not transport.close() or process.returncode != 0:
                    raise AssertionError("Probe did not exit cleanly")
    finally:
        # Resource release must run even when an evidence sink itself fails.
        started = time.perf_counter()
        cleanup = launcher.close()
        write_json(directory / "process-cleanup.json", [dict(pid=p.pid, released=p.released, steps=p.cleanup_steps) for p in cleanup])
        tails = getattr(launcher, "_stderr_tails", {})
        write_json(directory / "bridge-stderr.json", {str(pid): list(lines) for pid, lines in tails.items()})
        released = all(p.released for p in cleanup)
        journal.event("phase.finished", name="bridge.owned_cleanup", outcome="succeeded" if released else "failed",
                      durationMs=(time.perf_counter() - started) * 1000)
        if not released:
            raise AssertionError("Owned probe process was not released")
    return {"messagesVerified": len(config["payloadCharacters"]), "cleanup": "succeeded", "visualReview": "not_applicable"}


def wps_case(bundle, directory, config, journal, reopen=False):
    from wps_skills.cli.task import main
    from wps_skills.windows.bridge_runtime import windows_powershell_executable
    output = directory / (directory.parents[2].name + "-" + directory.name + ".xlsx")
    cells = dict(config["assertions"]["artifactCells"])
    input_path = None
    if reopen:
        with journal.phase("fixture.prepare"):
            previous = directory.parent / (directory.name.split("-", 1)[0] + "-wps")
            prior = read_json(previous / "result.json")
            if prior["state"] != "passed" or Path(prior["artifact"]).name != prior["artifact"]:
                raise ValueError("A successful preceding wps fixture is required")
            source = previous / prior["artifact"]
            if digest(source) != prior["artifactSha256"]:
                raise ValueError("The prepared workbook changed after verification")
            input_path = directory / (directory.parents[2].name + "-" + directory.name + "-input.xlsx")
            shutil.copy2(source, input_path)
            write_json(directory / "fixture.json", {"sourceCase": previous.name, "sha256": digest(input_path), "path": str(input_path)})
        cells = {"A1": cells["A1"] + " / reopen", "B1": cells["B1"] + 1}
    request = make_request(output, cells, input_path)
    write_json(directory / "expected-cells.json", cells)
    write_json(directory / "input.json", request)
    submitted = directory / "submission.json"
    write_json(submitted, request)  # CLI consumes this copy; immutable input remains.
    stdout = io.StringIO()
    with journal.phase("wps.task"):
        code = main(["--app", "excel", "--task-file", str(submitted)], output_stream=stdout, required_application="excel")
        raw = stdout.getvalue()
        (directory / "task.stdout").write_text(raw, encoding="utf-8")
        response = json.loads(raw)
        write_json(directory / "response.json", response)
        if code != 0 or response.get("outcome") != config["assertions"]["taskOutcome"]:
            raise AssertionError({"exitCode": code, "stop": response.get("stop"), "outcome": response.get("outcome")})
        if response["cleanup"]["outcome"] != config["assertions"]["cleanupOutcome"]:
            raise AssertionError(response["cleanup"])
    with journal.phase("wps.artifact_verify"):
        artifact = verify_artifact(output, cells)
        write_json(directory / "artifact-verification.json", artifact)
    from com_report import summarize
    com = summarize(directory)
    if not com["diagnosticsComplete"]:
        raise AssertionError("COM diagnostic evidence incomplete; inspect com-summary.json")
    with journal.phase("wps.close_verified_document"):
        cp = subprocess.run([windows_powershell_executable(), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                             "-File", str(bundle / "close_owned.ps1"), "-Path", str(output)],
                            capture_output=True, encoding="utf-8", errors="replace", timeout=20)
        (directory / "document-cleanup.stdout").write_text(cp.stdout, encoding="utf-8")
        (directory / "document-cleanup.stderr").write_text(cp.stderr, encoding="utf-8")
        if cp.returncode:
            raise AssertionError("Verified test document cleanup failed; see document-cleanup.stderr")
    return {"taskId": response["taskId"], "artifact": output.name, "artifactSha256": digest(output), "verification": artifact,
            "comConnection": com["connection"], "comCoverage": com["coverage"],
            "cleanup": "succeeded", "visualReview": "pending" if config["assertions"]["visualReview"] == "required" else "not_required"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--bundle", required=True, type=Path)
    parser.add_argument("--directory", required=True, type=Path)
    parser.add_argument("--case", required=True, choices=("pipe", "wps", "wps-open"))
    args = parser.parse_args()
    bundle, directory = args.bundle.resolve(), args.directory.resolve()
    journal = Journal(directory / "events.jsonl", case=args.case, clockDomain="windows-child")
    journal.event("python.entry")
    sys.path.insert(0, str(bundle / "wps-excel/runtime/src/main/python"))
    os.environ.update(WPS_TRACE_DIR=str(directory / "traces"), WPS_SKILLS_TASK_DIR=str(directory / "receipts"), WPS_COM_DIAGNOSTIC_DIR=str(directory / "com"),
                      PYTHONUTF8="1", PYTHONIOENCODING="utf-8")
    config = read_json(bundle / "experiment.json")
    try:
        with journal.phase("python.import_and_instrument"):
            transport_type = install_measurements(bundle, directory, config, journal)
        if args.case == "pipe":
            result = pipe_case(bundle, directory, config, journal, transport_type)
        else:
            result = wps_case(bundle, directory, config, journal, reopen=args.case == "wps-open")
        write_json(directory / "result.json", dict(state="passed", **result))
        return 0
    except BaseException as error:
        import traceback
        traceback.print_exc()
        failure = {"state": "failed", "error": repr(error), "scene": "preserved where still available"}
        try:
            journal.event("case.failed", error=repr(error))
        except Exception as logging_error:
            failure["loggingError"] = repr(logging_error)
        write_json(directory / "result.json", failure)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
