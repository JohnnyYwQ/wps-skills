"""Windows desktop acceptance using only packaged CLI commands for document work.

Run: python managed_live_acceptance.py --packages DIR --output NEW_DIR
This maintained test harness is not required by Skill consumers.
"""

import argparse
import ctypes
import json
import os
from pathlib import Path
import subprocess
import sys
import traceback
import uuid
import zipfile
import xml.etree.ElementTree as ET


def run(packages, output, application=None):
    output.mkdir(parents=True, exist_ok=False)
    report = {"status": "running", "applications": []}
    try:
        for app, create, extension in (("word", "createDocument", "docx"),
                                       ("excel", "createWorkbook", "xlsx"),
                                       ("ppt", "createPresentation", "pptx")):
            if application is not None and app != application:
                continue
            entry = packages / ("wps-" + app) / "scripts" / (app + ".py")
            item = {"app": app, "commands": [], "status": "running"}
            report["applications"].append(item)
            env = dict(os.environ, PYTHONPATH="", PYTHONIOENCODING="utf-8",
                       WPS_SKILLS_SESSION_DIR=str(output / "sessions"))

            def cli(*args, expected=0):
                process = subprocess.run([sys.executable, str(entry), "--app", app, *args],
                                         env=env, capture_output=True, text=True, encoding="utf-8",
                                         creationflags=subprocess.CREATE_NO_WINDOW, timeout=120)
                result = json.loads(process.stdout)
                item["commands"].append({"args": list(args), "exitCode": process.returncode, "result": result})
                assert process.returncode == expected, (process.returncode, result, process.stderr)
                return result

            ready = cli("--start", "--timeout", "90")
            handle = ready["handle"]
            step = 1

            def action(name, params=None, expected=0):
                nonlocal step
                parameters = output / (app + "-params-%d.json" % step)
                parameters.write_text(json.dumps(params or {}, ensure_ascii=False), encoding="utf-8")
                result = cli("--call", handle, "--step", str(step), "--action", name,
                             "--params-file", str(parameters), "--timeout", "100", expected=expected)
                assert result["response"]["sessionId"] == ready["ready"]["sessionId"]
                step = result["nextStep"]
                return result["response"].get("data") or result["response"].get("error")

            try:
                created = action(create)
                repeated = cli("--call", handle, "--step", "1", "--action", create)
                assert repeated["response"]["data"] == created
                rejected = action(create, expected=2)
                assert rejected["code"] == "SESSION_DOCUMENT_ALREADY_BOUND"
                text = "包内入口验证 中文 😀 " + app
                if app == "word":
                    action("writeContent", {"anchor": {"kind": "documentEnd"},
                           "blocks": [{"kind": "paragraph", "runs": [{"text": text}]}]})
                    observed = action("inspectDocument", {"scope": {"kind": "document"},
                                      "limits": {"maxTextCharacters": 4096, "maxParagraphs": 64, "maxRuns": 256}})
                    assert text in observed["text"] and not observed["truncated"]
                elif app == "excel":
                    sheet = action("listWorksheets", {"offset": 0, "limit": 100})["worksheets"][0]["name"]
                    observed = action("readRange", {"sheet": sheet, "address": "A1"})
                    action("writeRange", {"sheet": sheet, "address": "A1", "expectedToken": observed["token"], "values": [[text]]})
                    observed = action("readRange", {"sheet": sheet, "address": "A1"})
                    assert observed["cells"][0][0]["value"] == text
                else:
                    slides = action("listSlides")
                    slides = action("addSlide", {"position": 1, "expectedToken": slides["token"]})
                    slide_id = slides["slides"][0]["id"]
                    observed = action("getSlideInfo", {"slideId": slide_id})
                    action("addTextBox", {"slideId": slide_id, "expectedToken": observed["token"],
                           "text": text, "left": 40, "top": 40, "width": 600, "height": 100})
                    observed = action("getSlideInfo", {"slideId": slide_id})
                    assert text in json.dumps(observed, ensure_ascii=False)
                path = output / ("managed-" + uuid.uuid4().hex + "." + extension)
                saved = action("saveAs", {"outputPath": str(path), "overwritePolicy": "failIfExists"})
                assert saved["documentState"]["persistenceState"] == "saved"
                assert saved["artifact"]["sizeBytes"] == path.stat().st_size > 0
                with zipfile.ZipFile(path) as archive:
                    content = "\n".join("".join(ET.fromstring(archive.read(name)).itertext())
                                        for name in archive.namelist() if name.endswith(".xml"))
                    assert text in content
                item["artifact"] = str(path)
                window = created.get("window") or saved.get("window")
            finally:
                closed = cli("--close", handle)
                assert closed["sessionOutcome"]["outcome"] == "succeeded"
            assert cli("--close", handle)["sessionOutcome"] == closed["sessionOutcome"]
            # A fresh Session can lease the saved file after normal cleanup;
            # it opens the exact output explicitly and performs no mutation.
            reopened = cli("--start")
            old_handle, handle = handle, reopened["handle"]
            try:
                parameters = output / (app + "-open.json")
                parameters.write_text(json.dumps({"path": str(path)}), encoding="utf-8")
                result = cli("--call", handle, "--step", "1", "--action",
                             {"word": "openDocument", "excel": "openWorkbook", "ppt": "openPresentation"}[app],
                             "--params-file", str(parameters))
                assert result["response"]["outcome"] == "succeeded"
            finally:
                cli("--close", handle)
            if window and window.get("hwnd"):
                assert ctypes.windll.user32.IsWindow(int(window["hwnd"]))
            assert cli("--status", old_handle)["state"] == "closed"
            item["status"] = "passed"
        report["status"] = "passed"
    except BaseException as exc:
        report.update(status="failed", error=repr(exc), traceback=traceback.format_exc())
        raise
    finally:
        (output / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--packages", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--app", choices=("word", "excel", "ppt"))
    args = parser.parse_args()
    run(args.packages.resolve(), args.output.resolve(), args.app)
