"""Opt-in installed-package Word Task acceptance in a new Windows test directory.

The caller runs this in the interactive desktop and independently inspects/closes
only the saved test document afterwards. No installed Skill is modified.
"""

import argparse
import copy
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--skill", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if os.name != "nt":
        parser.error("Native acceptance requires Windows WPS")
    root = args.output.resolve()
    root.mkdir(parents=True, exist_ok=False)
    entry = args.skill.resolve() / "scripts/word.py"
    environment = dict(os.environ, WPS_SKILLS_TASK_DIR=str(root / "receipts"), PYTHONIOENCODING="utf-8")
    report = {"cases": [], "tasks": []}

    def dump(path, value):
        path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")

    def run(label, *arguments, code=0):
        completed = subprocess.run([sys.executable, str(entry), "--app", "word", *arguments],
                                   capture_output=True, env=environment, timeout=150)
        (root / (label + ".stdout.json")).write_bytes(completed.stdout)
        (root / (label + ".stderr.txt")).write_bytes(completed.stderr)
        result = json.loads(completed.stdout)
        assert completed.returncode == code, (label, result, completed.returncode)
        return result

    def action(name, params=None, identifier=None):
        value = {"address": {"app": "word", "action": name}, "params": params or {}}
        if identifier:
            value["id"] = identifier
        return value

    def submit(label, request, code=0):
        source = root / (label + ".json")
        dump(source, request)
        result = run(label, "--task-file", str(source), code=code)
        assert result["cleanup"]["outcome"] == "succeeded", result
        assert result["outcome"] == ("succeeded" if code == 0 else "failed"), result
        assert result["taskFile"]["state"] == "removed" and not source.exists()
        receipt = Path(result["recordPath"])
        assert json.loads(receipt.with_name("request.json").read_text(encoding="utf-8")) == request
        assert "sessionOutcome" not in result and "ready" not in result
        for item in [result["document"], *result["steps"], *(v for v in result["completion"].values() if v is not None)]:
            assert "checks" not in item
            if item["response"] is not None:
                assert "sessionId" not in item["response"]
        expected = dict(result)
        expected.pop("taskFile")
        assert run(label + "-query", "--task-status-file", str(source), code=code) == expected
        assert run(label + "-repeat-missing", "--task-file", str(source), code=code) == expected
        dump(source, request)
        repeated = run(label + "-repeat-restored", "--task-file", str(source), code=code)
        assert repeated == result and not source.exists()
        report["tasks"].append(result["taskId"])
        return source, result

    def paragraphs(path):
        with zipfile.ZipFile(path) as archive:
            body = ET.fromstring(archive.read("word/document.xml")).find("{*}body")
        return body.findall("{*}p")

    def check_docx(path, expected):
        paras = paragraphs(path)
        assert ["".join(t.text or "" for t in p.findall(".//{*}t")) for p in paras] == [expected]
        assert paras[0].find("{*}pPr/{*}jc").attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") == "center"
        for run in paras[0].findall("{*}r"):
            if run.find("{*}t") is not None:
                assert run.find("{*}rPr/{*}b") is not None
                size = run.find("{*}rPr/{*}sz")
                assert size is not None and size.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") == "24"

    output = root / "source.docx"
    original_text = "自动验证：订单 TF-092，数量 17。中文、English 与🙂。"
    request = {
        "app": "word", "document": action("createDocument"),
        "steps": [action("writeContent", {
            "anchor": {"kind": "documentEnd"}, "blocks": [{"kind": "paragraph",
                "format": {"alignment": "center"},
                "runs": [{"text": original_text, "format": {"fontSizePt": 12, "bold": True}}]}],
        }, "write")],
        "completion": [action("saveAs", {"outputPath": str(output), "overwritePolicy": "renameIfExists"})],
    }
    source, created = submit("create", request)
    check_docx(output, original_text)
    report["cases"].append("create-unicode-one-paragraph-format-save-without-agent-checks")
    destination = root / "report.docx"
    sentinel = root / "report (1).docx"
    for path in (destination, sentinel):
        shutil.copyfile(output, path)
    protected = {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in (output, destination, sentinel)}

    modified_text = original_text.replace("数量 17", "数量 23")
    edited = {
        "app": "word", "document": action("openDocument", {"path": str(output)}),
        "steps": [action("replaceContent", {
            "target": {"kind": "query", "query": {"scope": {"kind": "document"}, "text": "数量 17",
                       "caseSensitive": True, "wholeWord": False}, "expectedMatchCount": 1},
            "replacement": {"kind": "text", "runs": [{"text": "数量 23", "format": {"fontSizePt": 12, "bold": True}}]},
        }, "replace")], "completion": [action("saveAs", {"outputPath": str(destination), "overwritePolicy": "renameIfExists"})],
    }
    _, saved = submit("edit", edited)
    save_data = saved["completion"]["save"]["response"]["data"]
    final = root / "report (2).docx"
    assert Path(save_data["artifact"]["path"]) == final
    assert save_data["outputResolution"]["attempts"] == 3
    for filename, digest in protected.items():
        assert hashlib.sha256(Path(filename).read_bytes()).hexdigest() == digest
    check_docx(final, modified_text)
    report["cases"].append("replace-readback-and-save-as-renaming-with-originals-preserved")

    inspection = {"scope": {"kind": "document"},
                  "limits": {"maxTextCharacters": 4096, "maxParagraphs": 32, "maxRuns": 64}}
    unsaved = copy.deepcopy(request)
    unsaved["completion"] = []
    unsaved["steps"].append(action("inspectDocument", inspection, "inspect"))
    _, unsaved_result = submit("new-unsaved", unsaved)
    observed = unsaved_result["steps"][-1]["response"]["data"]
    assert observed["text"] == original_text
    assert observed["documentState"]["persistenceState"] == "unsaved"
    assert unsaved_result["completion"] == {"save": None, "pdf": None}
    report["cases"].append("new-document-empty-completion-remains-unsaved")

    edit_path = root / "edit-without-save.docx"
    shutil.copyfile(output, edit_path)
    original_hash = hashlib.sha256(edit_path.read_bytes()).hexdigest()
    no_save = copy.deepcopy(edited)
    no_save["document"] = action("openDocument", {"path": str(edit_path)})
    no_save["completion"] = []
    no_save["steps"].append(action("inspectDocument", inspection, "inspect"))
    _, memory_edit = submit("existing-unsaved", no_save)
    observed = memory_edit["steps"][-1]["response"]["data"]
    assert observed["text"] == modified_text and observed["documentState"]["persistenceState"] == "modified"
    assert hashlib.sha256(edit_path.read_bytes()).hexdigest() == original_hash
    report["cases"].append("existing-edit-memory-changed-disk-hash-unchanged")

    export_only = {"app": "word", "document": no_save["document"],
                   "steps": [action("inspectDocument", inspection, "inspect")],
                   "completion": [action("exportPdf", {"outputPath": str(root / "only.pdf"), "overwritePolicy": "failIfExists"})]}
    _, exported = submit("pdf-only", export_only)
    assert exported["completion"]["save"] is None
    assert exported["steps"][0]["response"]["data"]["text"] == modified_text
    assert exported["completion"]["pdf"]["response"]["data"]["documentStateBefore"]["persistenceState"] == "modified"
    assert hashlib.sha256(edit_path.read_bytes()).hexdigest() == original_hash
    assert (root / "only.pdf").read_bytes().startswith(b"%PDF-")
    report["cases"].append("pdf-only-preserves-unsaved-document-state")

    stale = {"app": "word", "document": no_save["document"], "completion": [], "steps": [
        action("replaceContent", {"target": {"kind": "range", "range": observed["paragraphs"][0]["range"]},
               "replacement": {"kind": "blocks", "blocks": [{"kind": "paragraph", "runs": [{"text": "must not write"}]}]}}, "stale")
    ]}
    _, stale_result = submit("cross-task-range", stale, code=2)
    assert stale_result["steps"][0]["response"]["error"]["code"] == "STALE_CONTENT_RANGE"
    assert hashlib.sha256(edit_path.read_bytes()).hexdigest() == original_hash
    report["cases"].append("cross-task-old-range-rejected")

    failed_save = copy.deepcopy(export_only)
    failed_save["includeExistingChanges"] = True
    failed_save["completion"].insert(0, action("saveAs", {"outputPath": str(sentinel), "overwritePolicy": "failIfExists"}))
    failed_save["completion"][1]["params"]["outputPath"] = str(root / "must-not-export.pdf")
    _, stopped = submit("save-failure", failed_save, code=2)
    assert stopped["stop"]["stepId"] == "task_save"
    assert stopped["completion"]["save"]["state"] == "failed"
    assert stopped["completion"]["pdf"]["state"] == "not_executed"
    assert stopped["steps"][0]["response"]["data"]["text"] == modified_text
    assert not (root / "must-not-export.pdf").exists()
    assert hashlib.sha256(edit_path.read_bytes()).hexdigest() == original_hash
    report["cases"].append("save-failure-stops-pdf-with-content-and-original-file-retained")

    bad = copy.deepcopy(request)
    bad["completion"][0]["checks"] = [{"actual": {"$ref": {"step": "task_save", "path": ["response", "data"]}},
                                       "op": "equals", "expected": "saved"}]
    bad_input = root / "invalid-checks.json"
    dump(bad_input, bad)
    rejected = run("invalid-checks", "--task-file", str(bad_input), code=4)
    assert rejected["state"] == "rejected" and rejected["recordPath"] is None and bad_input.exists()
    bad.update(version=2, taskId="legacy-bypass")
    dump(bad_input, bad)
    rejected = run("legacy-rejected", "--task-file", str(bad_input), code=4)
    assert rejected["recordPath"] is None and bad_input.exists()
    changed = copy.deepcopy(request)
    changed["completion"][0]["params"]["outputPath"] = str(root / "must-not-exist.docx")
    dump(source, changed)
    rejected = run("input-conflict", "--task-file", str(source), code=4)
    assert rejected["stop"]["error"]["code"] == "TASK_INPUT_CONFLICT"
    assert source.exists() and not (root / "must-not-exist.docx").exists()
    report["cases"].extend(["lost-response-and-repeated-input-do-not-replay", "legacy-and-checks-rejected", "input-conflict-retained"])
    dump(root / "expected.json", {"outputPath": str(final), "lines": [modified_text]})
    dump(root / "report.json", dict(report, passed=True, protected=protected, outputPath=str(final)))
    print(json.dumps(report, ensure_ascii=True))


if __name__ == "__main__":
    main()
