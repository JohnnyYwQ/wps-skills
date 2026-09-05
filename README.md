# WPS Automation Foundation

[English](README.md) | [简体中文](README.zh-CN.md)

This repository contains the shared WPS Action Session foundation and production Windows Word, Excel and PPT slices, including independently installable `wps-word`, `wps-excel` and `wps-ppt` Application Skills.

## What exists

- An application-scoped `ActionSession` with immutable one-document binding, closed Controller Result handling, exact-document dispatch, and idempotent cleanup.
- A canonical JSONL `SessionHost` with strict Action Request decoding, lifecycle records, per-Action/session timing journals, traced request rejection, and terminal ordering.
- A complete, validated Word target Contract Set and compact Action Index for fourteen designed Actions.
- A production Word Application Contract Set containing fourteen Actions: document create/open; structured write, inspect, find, and replace; table and image insertion; header/footer, page-layout, and break changes; in-place save; and PDF export. Every advertised required Action has a real handler. First save and Save As preserve the exact live document.
- A Session-owned suspended-process launcher with Windows Job Object containment, one lazy native `System32` Windows PowerShell bridge, stable-file-identity acquisition, and a cross-process guard/Lease/quarantine coordinator.
- Real structured Word writing, including separate Western and East Asian run fonts, plus bounded search/replacement, tables, embedded images, headers/footers, layout, page/section breaks, and PDF export, with revision-aware results and operation-specific read-back verification.
- Hidden bridge launch at both process layers: Windows creates the child with `CREATE_NO_WINDOW`, and native PowerShell is also given `-WindowStyle Hidden`, so automation does not open a console window.
- Word establishment makes WPS visible and activates only the exact created or opened document. A newly created WPS application uses normal, not maximized, outer-window state. On either create or attach, a genuinely tiny top-level frame is repaired to a centered 80% of its monitor work area; an already reasonable user window is left alone. Later Actions still dispatch through the retained binding rather than window focus.
- A `scripts/call.py --session --app word` production entry point on Windows. Excel uses `--app excel`; PPT uses `--app ppt` with its own contracts and native Presentation backend.
- Side-effect-free `--app word --index` and `--app word --resolve ACTION...` discovery, generated from the production Contract Set.
- A reusable Python Session Client that preserves terminal Action errors, serializes calls, and separates document results from cleanup outcomes.
- A Word Skill under `src/main/resources/skills/wps-word`, with task guidance, executable client examples, and standalone assembly.
- Local conformance tests for the Session Core, protocol channels, and standalone Skill packaging.

The previous combined Skill, global Manifest and discovery CLI, multi-application Runtime, WPS controllers, Linux/OpenXML backends, and live harnesses have been removed. They are recoverable from Git history but are not compatibility interfaces.

## PPT

The PPT slice supports **37 native Actions** for new and existing `.pptx` presentations.
The 18 newly admitted common Actions add shape fill/border styles, paragraph and
text-box formatting, shape naming/stacking/alignment/distribution, slide background
and visibility settings, speaker notes, literal find/replace, embedded PNG/JPEG
images and native table text read/write. Each edit checks its appropriate
observation token and verifies native readback.

Structure operations support up to 200 slides; slide snapshots support 100
top-level shapes and 10000 UTF-16 units per shape. Tables are bounded to 100 cells;
images to 20 MiB and 40 million pixels. Creation, first save, Save As, PDF and single-slide PNG export are admitted. Charts, animations and shape duplication remain unavailable.

```powershell
python scripts/build/ppt.py --output build/skills/wps-ppt
python build/skills/wps-ppt/scripts/ppt.py --app ppt --index
python scripts/validate/ppt.py --output-dir build/ppt-acceptance
python scripts/validate/ppt_common.py --output-dir build/ppt-common-acceptance
```

Execution requires Windows WPS Presentation registered as `KWPP.Application`.
Run `scripts/demo/ppt.ps1 -PythonPath <absolute-python.exe-path>` from an ordinary
Windows desktop PowerShell for the complete hidden-console launcher. The
[Chinese README](README.zh-CN.md) provides a complete copy-and-paste command.
The demo copies a blank fixture, edits it through the production Session, verifies
the bound document's own visible window and saves while leaving WPS open.

See [the PPT Skill](src/main/resources/skills/wps-ppt/SKILL.md) and
[native capability evidence](src/test/resources/wps_skills/ppt/type_library/EVIDENCE.md).

## Excel

The Excel slice supports 34 Actions for new and existing `.xlsx` workbooks: open/inspect/save,
worksheet discovery/create/rename/copy/move/delete, bounded value and formula edits,
clear/copy/find/replace, sorting/filtering, row/column insertion/deletion, formatting,
merge/unmerge and row/column sizing. Common statistical, lookup, date and text formulas
are supported. Region edits require `readRange` tokens; worksheet and structural edits
require `getWorksheetInfo` tokens. Both observe at most 1000 cells; structural edits
require the entire used range to fit that limit. Every mutation reads back its result.
WPS file replacement during Save is protected by continuous locator/file claims.
Workbook creation, first save, Save As and PDF export are admitted; charts and pivot tables remain unavailable.

Run `python scripts/demo/excel.py` in a Windows desktop terminal to see WPS fill,
calculate, format and save a demo workbook while leaving its window open.
Run `python scripts/validate/excel_common.py --output-dir build/excel-common-acceptance`
for common-action acceptance on a new disposable workbook.

```bash
python scripts/build/excel.py --output build/skills/wps-excel
python build/skills/wps-excel/scripts/excel.py --app excel --index
python scripts/call.py --app excel --resolve openWorkbook readRange writeRange save
```

See [the Excel Skill](src/main/resources/skills/wps-excel/SKILL.md) and
[live capability evidence](src/test/resources/wps_skills/excel/type_library/EVIDENCE.md).
Execution requires Windows WPS Spreadsheets registered as `KET.Application`.
Validated on WPS 12.0.0.28505. Run the opt-in acceptance on Windows with a **new**
output directory; it only edits its own generated workbooks and leaves them open:

```powershell
python scripts/validate/excel.py --output-dir build/excel-acceptance --wps-version 12.0.0.28505
```

## Verify the foundation

Python 3.8 or newer is sufficient:

```bash
PYTHONPATH=src/main/python python -m unittest discover -s src/test/python -p 'test_*.py'
```

The suite exercises local fakes, real subprocess protocol channels, and relocated Skill distributions. No WPS installation or external account is required.

## Build and use the Word Skill

```bash
python scripts/build/word.py --output "<output-dir>/wps-word"
python "<output-dir>/wps-word/scripts/word.py" --app word --index
python "<output-dir>/wps-word/scripts/word.py" --app word --resolve createDocument writeContent inspectDocument
```

Replace `<output-dir>` with your chosen output directory. The build creates a `wps-word/` directory containing `SKILL.md`, references, the thin `scripts/word.py` entry point, and a snapshot of the Python Runtime and PowerShell resources. Copy this complete directory into the target agent's Skill directory. The source tree remains the only maintained implementation; the build includes a SHA-256 file inventory and refuses to overwrite an existing destination. Use `--output <new-directory>/wps-word` for another build.

Read [the Word Skill](src/main/resources/skills/wps-word/SKILL.md) for document intent, discovery, execution, verification, persistence, and failure handling. Its [Session guide](src/main/resources/skills/wps-word/references/session.md) includes a Python task example using `open_session()` and `client.call(address, params)`; the caller decides each next Action after the prior response. Discovery works on macOS/Linux too, while document execution runs on the Windows WPS host.

The source Skill's `scripts/word.py` also works directly from its source location. A deployed Skill uses its bundled Runtime and requires no repository checkout or third-party Python package.

On the Windows WPS host, start a Word Session with:

```bash
python scripts/call.py --session --app word
```

The Host emits `session.ready`, accepts newline-delimited Action Requests, and reuses one exact live Word document and one bridge until `{"control":"close"}`. `saveAs` handles first save and a new output path while retaining all old and new locator/file claims until cleanup.

Protocol v1 remains closed: timing is diagnostic rather than an extra Action Response field. The `traceLog` in each Action Response records that Action's `elapsedMs`; the Session `traceLog` ends with `sessionElapsedMs`, `actionExecutionElapsedMs`, `cleanupElapsedMs`, and `actionCount` so wall time and actual Action execution are not confused.

Normal Session cleanup deliberately leaves the document open. A controlled debug run that creates disposable content may opt into test-only cleanup:

```bash
python scripts/call.py --session --app word --debug-close-created-document
```

That flag discards and closes only a document created by that Session. It never closes a document acquired through `openDocument`, is not a Word Action, and must not be used when the newly created document contains content that should be retained.

## Word desktop demo

Run `./scripts/demo/word.ps1 -PythonPath <absolute-python.exe-path>` in an ordinary Windows desktop PowerShell. It edits a unique blank DOCX fixture through the production Session, writes formatted text and a native table, reads back the result, explicitly saves, and leaves the window open. The native demo workflow is checked by `python scripts/validate/word.py --output-dir build/word-acceptance`.

Repository commands are grouped under `scripts/build/`, `scripts/demo/`, and `scripts/validate/`. See the [script guide](scripts/README.md) for all entry points and their scope.

## Project layout

The repository uses Java-style source sets while retaining Python packages:

```text
src/
  main/
    python/wps_skills/
      cli/          # argument parsing, discovery, and package building
      client/       # caller-side Session Protocol and process lifetime
      core/         # application-independent Action Session Core
      host/         # JSONL Session Host
      word/         # Word contracts and Adapter; windows/ owns backend and assembly
      excel/        # Excel contracts and Adapter; windows/ owns backend and assembly
      ppt/          # PPT contracts and Adapter; windows/ owns backend and assembly
      windows/      # shared process ownership, coordination, transport, and desktop
    resources/
      wps_skills/word/windows/  # PowerShell/WPS bridge and Action resources
      wps_skills/excel/windows/ # Excel bridge, Actions, and demo launcher
      wps_skills/ppt/windows/   # PPT bridge, Actions, and demo launcher
      skills/wps-ppt/          # PPT Skill source, references, and entry point
      wps_skills/windows/      # shared bridge and native coordination
      skills/wps-word/         # Skill source, references, and thin entry point
      skills/wps-excel/        # Excel Skill source, references, and entry point
  test/
    python/tests/   # unit tests, live acceptance, and executable Python fixtures
    resources/      # PowerShell test resources and type-library evidence
scripts/            # thin repository entry points only
build/              # generated packages and acceptance evidence (ignored)
```

Tests use a separate `tests` namespace because a second top-level Python package named `wps_skills` would shadow the production package during discovery.

See [the file-by-file guide (Chinese)](FILE_STRUCTURE.md) for directory rules, each maintained file's purpose, and the execution flow.

## Skill and capability references

- [Word Skill](src/main/resources/skills/wps-word/SKILL.md): task workflow and capability discovery.
- [Session guide](src/main/resources/skills/wps-word/references/session.md): Python client usage and a runnable example.
- [Content guide](src/main/resources/skills/wps-word/references/content.md): ranges, revisions, text formatting, and units.
- [Verification guide](src/main/resources/skills/wps-word/references/verification.md): verification, persistence, and failure handling.
- [Word contracts](src/main/python/wps_skills/word/contracts.py): authoritative Action definitions and the derived production Contract Set.
- [WPS Writer Type Library snapshot](src/test/resources/wps_skills/word/type_library/wps_writer_api.py): capability evidence only; never imported by the Runtime.

## New documents and persistence

The production surface now contains 85 Actions: Word 14, Excel 34 and PPT 37.
Use `createDocument`, `createWorkbook` or `createPresentation` for explicit creation,
then `saveAs` with an absent absolute destination and `overwritePolicy: "failIfExists"`.
Existing outputs are never overwritten by Save As. Later `save` uses the new path.
Excel/PPT export PDF without saving the source; PPT also exports a slide as PNG.
Save As/PDF verification is bounded to 20 Excel worksheets with at most 1000 used
cells each, or 200 PPT slides with at most 100 top-level shapes each. These observed
fields do not prove full fidelity for unsupported document features.

Run `python scripts/validate/persistence.py --output-dir build/persistence-acceptance`
on the Windows desktop for the cross-application native acceptance. Skill packages
are delivered as the complete directories under `build/skills/`.
