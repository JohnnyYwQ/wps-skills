"""Opt-in native Word Task acceptance; leaves the test document open."""

import argparse
import json
import os
from pathlib import Path
import sys
import zipfile


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--skill", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    skill = args.skill.resolve()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    sys.path.insert(0, str(skill / "runtime/src/main/python"))
    os.environ["WPS_SKILLS_TASK_DIR"] = str(output / "receipts")
    from wps_skills.word.skill import main as task_main
    import io

    sentinel = b"EXISTING FILE MUST REMAIN UNCHANGED"
    for name in ("report.docx", "report (1).docx", "report.pdf", "report (1).pdf"):
        (output / name).write_bytes(sentinel)
    def action(name, params):
        return {"address": {"app": "word", "action": name}, "params": params}
    text = "Action retry keeps this document and writes this text once. 中文验证。"
    request = {
        "app": "word",
        "document": action("createDocument", {}),
        "steps": [
            dict(action("writeContent", {"anchor": {"kind": "documentEnd"},
                "blocks": [{"kind": "paragraph", "runs": [{"text": text}]}]}), id="write"),
            dict(action("inspectDocument", {"scope": {"kind": "document"},
                "limits": {"maxTextCharacters": 4096, "maxParagraphs": 16, "maxRuns": 32}}), id="inspect"),
        ],
        "completion": [action(name, {"outputPath": str(output / filename), "overwritePolicy": "renameIfExists"})
                       for name, filename in (("saveAs", "report.docx"), ("exportPdf", "report.pdf"))],
    }
    source = output / "task.json"
    source.write_text(json.dumps(request, ensure_ascii=False), encoding="utf-8")
    stream = io.StringIO()
    code = task_main(["--app", "word", "--task-file", str(source), "--timeout", "90"], output_stream=stream)
    result = json.loads(stream.getvalue())
    assert code == 0, result
    assert result["steps"][1]["response"]["data"]["text"] == text
    (output / "response.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    assert result["outcome"] == "succeeded", result
    assert result["cleanup"]["outcome"] == "succeeded", result
    for key, extension in (("save", "docx"), ("pdf", "pdf")):
        data = result["completion"][key]["response"]["data"]
        assert Path(data["artifact"]["path"]) == output / ("report (2)." + extension), data
        assert data["outputResolution"]["attempts"] == 3, data
        assert data["outputResolution"]["renamed"] is True, data
        assert not data["replacedExisting"], data
        assert (output / ("report." + extension)).read_bytes() == sentinel
        assert (output / ("report (1)." + extension)).read_bytes() == sentinel
    with zipfile.ZipFile(output / "report (2).docx") as archive:
        xml = archive.read("word/document.xml").decode("utf-8")
        assert xml.count(text) == 1, xml
    assert (output / "report (2).pdf").read_bytes().startswith(b"%PDF-")
    actions = [r["address"]["action"] for r in [result["document"], *result["steps"], *result["completion"].values()]]
    assert actions == ["createDocument", "writeContent", "inspectDocument", "saveAs", "exportPdf"], actions
    print(json.dumps({"status": "passed", "actions": actions, "saveAttempts": 3, "pdfAttempts": 3,
                      "sentinelsUnchanged": True, "textWrittenOnce": True, "output": str(output)}))


if __name__ == "__main__":
    main()
