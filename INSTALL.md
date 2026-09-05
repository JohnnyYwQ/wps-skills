# Windows WPS Session Setup

The repository contains a production Windows Word Session Host and Adapter for a thirteen-Action surface, plus a 31-Action Windows Excel slice, a 33-Action PPT slice and three standalone Application Skills. Document execution runs directly on the Windows machine that has WPS; there is no separate Windows Host service to install.

## Requirements

- Python 3.8 or newer.
- Windows with WPS Writer registered as `KWPS.Application`.
- Native Windows PowerShell under `%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe`. The Runtime does not fall back to the slower WOW64 host.
- No third-party Python package or external service.

The local unit suite itself does not start WPS, COM, or PowerShell.

## Assemble and install the Skill

From the repository root:

```bash
python scripts/build/word.py --output build/skills/wps-word
```

The destination must not already exist. The complete output has this layout:

```text
wps-word/
  SKILL.md
  agents/openai.yaml
  references/
  scripts/word.py
  runtime/
    files.sha256.json
    src/main/python/wps_skills/
    src/main/resources/wps_skills/word/windows/
```

Copy the complete `wps-word` directory to the Skill location supported by the target agent. Do not install only `SKILL.md` or only the source resources directory: the deployed entry point needs the bundled Runtime. To use this Skill on another execution host, place the complete directory on that host too and use that host's paths. No machine names, SSH credentials, or scheduled tasks are embedded.

Verify installation without starting WPS:

```powershell
python "C:\path\to\wps-word\scripts\word.py" --app word --index
python "C:\path\to\wps-word\scripts\word.py" --app word --resolve openDocument inspectDocument writeContent save
```

Successful resolution exits 0; a `partial` or `failed` batch exits 2 while still reporting every requested Action. An unavailable application exits 4 without publishing another application's contracts.

Follow `SKILL.md` and `references/session.md` to execute a task with the Python Session Client. If visible WPS output is needed, execute in the logged-in user's desktop session. SSH execution by itself does not establish desktop visibility; remote desktop launch is environment-specific and is not part of the Skill installer.

## Local verification

From the repository root, run:

```bash
PYTHONPATH=src/main/python python -m unittest discover -s src/test/python -p 'test_*.py'
```

In Windows PowerShell, set the same source root with:

```powershell
$env:PYTHONPATH = "src/main/python"
python -m unittest discover -s src/test/python -p "test_*.py"
```

## Windows production Session

Run this command inside the repository on the Windows WPS machine:

```bash
python scripts/call.py --session --app word
```

It emits `session.ready` on stdout, then accepts strict JSONL Action Requests. A normal existing-file flow is `openDocument`, any supported required Actions, final `inspectDocument`, explicit `save` when Persistence Intent requires it, then `{"control":"close"}`. The production Set also supports `findContent`, `replaceContent`, `insertTable`, `insertImage`, `setHeaderFooter`, `setPageLayout`, `insertBreak`, and `exportPdf`; only `saveAs` remains deferred. The Runtime keeps one exact document and one owned, console-hidden PowerShell bridge for the full Session. Excel uses `--app excel` with the same Session Protocol. PPT uses `--app ppt` with its independent Presentation contracts.

## Excel installation and acceptance

Excel additionally requires WPS Spreadsheets registered as `KET.Application`.
Build and copy the complete independent Skill directory:

```bash
python scripts/build/excel.py --output build/skills/wps-excel
python build/skills/wps-excel/scripts/excel.py --app excel --index
python scripts/call.py --session --app excel
```

The 31 admitted Actions cover existing `.xlsx` workbooks, worksheet management,
region reads/writes/copy/clear/find/replace, formulas, sorting/filtering, row/column
structure edits and sizing, font/color/alignment formats, merge/unmerge and in-place
save. Region operations are bounded to 1000 cells; worksheet structure operations
also require a used range within that limit. See the Excel Skill references for
individual constraints. Workbook creation, Save As, charts, pivot tables and PDF
export are not available. Session cleanup leaves the workbook open.

For a visible desktop demonstration without an extra console, invoke
`scripts/demo/excel.ps1 -PythonPath <absolute-python.exe-path>` from an existing
non-administrator Windows desktop PowerShell, or use the complete inline launcher command in
README.zh-CN.md when the shell does not permit `.ps1` execution. The launcher uses
`ProcessStartInfo.CreateNoWindow` and forwards UTF-8 progress and errors to the
existing terminal. User-only WPS COM registration can be unavailable to elevated
processes; the launcher now detects this before creating a demo workbook. Use a
normal PowerShell window for such installations. `python scripts/demo/excel.py`
remains available. It creates its own workbook, displays WPS, edits and saves it, and
checks that its window remains visible. SSH session 0 is rejected by this demo.

The source type-library snapshot is under
`src/test/resources/wps_skills/excel/type_library/` and is not shipped or imported
by the Runtime. On a Windows WPS machine, opt into live acceptance explicitly:

```powershell
python scripts/validate/excel.py --output-dir build/excel-acceptance --wps-version 12.0.0.28505
python scripts/validate/excel_common.py --output-dir build/excel-common-acceptance
```

The output directory must not exist. The test creates its own workbooks and a
report, checks persisted OOXML independently, verifies file/locator coordination,
and leaves the workbook open. It never edits a caller-supplied workbook. The
coordinator crash probe does not start WPS and leaves a quarantine marker for its
own disposable scratch path; do not use that path for another task.

## PPT installation and acceptance

WPS Presentation must be registered as `KWPP.Application` for the process user.
Use a non-administrator desktop PowerShell when registration is per user.
No pywin32 dependency is required.

```powershell
python scripts/build/ppt.py --output build/skills/wps-ppt
python build/skills/wps-ppt/scripts/ppt.py --app ppt --index
python scripts/call.py --session --app ppt
python scripts/validate/ppt.py --output-dir build/ppt-acceptance
python scripts/validate/ppt_common.py --output-dir build/ppt-common-acceptance
```

Copy the entire assembled `wps-ppt` directory to the agent's Skill location and
the Windows execution host. Discovery and package verification work without WPS;
execution requires Windows. The generated type library remains in test resources
and is excluded from distribution.

For the visible demo run `./scripts/demo/ppt.ps1 -PythonPath <absolute-python.exe-path>`
in a normal Windows desktop terminal, or paste the complete command from
README.zh-CN.md. The launcher uses the shared hidden process chain, and the demo
checks the exact document window rather than a title match. It edits a unique
copy of the bundled blank template, saves and keeps the window open.

PPT supports only existing ordinary `.pptx` files, in-place save and the documented
33 Actions. New presentation creation, first save, Save As and export are not
available. Read the Skill references for shape limits, observation token scopes,
whole-slide deletion, separate Latin/East Asian font properties, shape style tokens,
notes, embedded images and bounded native table text. Common acceptance verifies
persisted notes/table XML and embedded picture bytes independently.

## Admitting more capability

An additional Word, Excel or PPT Action is admitted only after it owns all of the following:

1. A complete Application Contract Set and generated Action Index.
2. An exact-document Application Adapter and controller at the shared seam.
3. Binding, persistence, handler, and real-WPS verification evidence.
4. Its own independently discoverable Application Skill.

Do not add a placeholder Skill, empty production Contract Set, fake production Adapter, global Action Manifest, or compatibility wrapper around the removed execution model.

## Repository command groups

Use `scripts/build/{word,excel,ppt}.py` to assemble standalone Skills,
`scripts/demo/{word,excel,ppt}.ps1` for visible desktop demonstrations, and
`scripts/validate/` for native acceptance. `scripts/call.py` remains the Session
Host and discovery entry. These repository paths are separate from each installed
Skill's `scripts/{word,excel,ppt}.py` client entry.

Word now has a visible demo: `./scripts/demo/word.ps1 -PythonPath <absolute-python.exe-path>`.
It prepares a unique blank DOCX fixture, performs all displayed edits through the
production Session, verifies the text/table, saves, and leaves the window open.
`python scripts/validate/word.py --output-dir build/word-acceptance` additionally
checks persisted DOCX content and the exact visible document window. This is a demo
workflow check, not a claim of full Word Action coverage. See [the script guide](scripts/README.md).
