# WPS Spreadsheets type-library evidence

`wps_excel_api.py` was supplied in the repository root and is preserved byte-for-byte.
The generated header identifies Upgrade WPS Spreadsheets 3.0 Object Library (Beta),
CLSID `{45541000-5750-5300-4B49-4E47534F4655}`, makepy 0.5.01, Python 3.12.1,
generated 2026-08-28. This Windows `mbcs` source is evidence, never imported by production or tests.

| Candidate Action | Declared COM members | Implementation |
| --- | --- | --- |
| openWorkbook | Workbooks.Open; _Workbook.FullName, FileFormat, Saved, ReadOnly | excel_bridge.ps1 |
| getWorkbookInfo | _Workbook.Name, Saved, ReadOnly, Date1904, Worksheets | excel_actions.ps1 |
| listWorksheets | Worksheets.Count, Item; _Worksheet.Name | excel_actions.ps1 |
| readRange | _Worksheet.Range; Range.Value2, Formula, HasFormula, Text, NumberFormat; Font.Bold | excel_actions.ps1 |
| writeRange | Range.Value2, ClearContents | excel_actions.ps1 |
| setFormulas | Range.Formula | excel_actions.ps1 |
| calculateRange | Range.Calculate | excel_actions.ps1 |
| formatRange | Range.NumberFormat; Font.Bold | excel_actions.ps1 |
| save | _Workbook.Save, Saved, FileFormat, FullName | excel_actions.ps1 |

Live admission requires `scripts/validate/excel.py` on Windows WPS, the negative-case
matrix it reports, and review of real artifacts and traces. Vendor declarations alone
are not runtime compatibility evidence. Generated reports belong under `build/`.

## Live admission, 2026-09-05

Validated on Windows 10.0.26200, Python 3.10.18, WPS Spreadsheets COM Version 12.0,
Build 28505 (observed directly from `KET.Application`). All nine Actions passed
through the actual Session Client, Session Host, owned native PowerShell process,
Excel Adapter and shared document coordinator.

The acceptance covered two worksheets including a Chinese name; page boundaries;
text, numeric, Boolean, blank and empty-string values; literal `=...` text; date
serial/format interpretation; exact formula read-back and calculation; separate
error codes for `#DIV/0!` versus a legitimate numeric value of -2146826281; number
format and bold; repeated in-place save; independent ZIP/XML artifact inspection;
normal cleanup and reacquisition. Negative checks covered missing/unbound targets,
second establishment, missing worksheets, merged-cell writes, matrix mismatch,
stale region tokens, same-file and hard-link lease conflicts, and lease conflicts
after WPS replaced the saved file. An isolated non-WPS subprocess test confirmed
that killing an in-flight owner after replacement quarantines the normalized
locator and blocks acquisition of the replacement identity.

WPS-specific implementation evidence:

- A COM Range returned by a PowerShell helper must be protected from pipeline
  enumeration. Indexed properties use explicit reflection `GetProperty`.
- Keep each exact worksheet COM reference for the Session: refetching released
  wrappers can change IUnknown pointers. Region tokens use a private stable
  Session-local worksheet marker and observed region contents.
- Explicit VARIANT property writes avoid PowerShell's mixed-value setter coercion.
- WPS marshals formula errors as integers. `WorksheetFunction.IsError` on the exact
  cell distinguishes them from ordinary numbers; display text is not the test.
- `Range.Calculate()` return values must be suppressed before emitting JSONL.
- Ordinary Save replaces the file identity. A handle-normalized locator fence and
  retained original/replacement file claims protect the immutable live binding.

This evidence covers the stated WPS build and bounded feature set. It does not
admit Save As, workbook creation, cross-sheet formulas, dynamic arrays, protected
worksheet modification or other file formats. The live script is repository-only;
production discovery is derived from the admitted contracts, never this snapshot.

## Common editing extension, 2026-09-05

The Application Contract Set now contains 31 Actions. On the same WPS build,
`common_live_acceptance.py` exercised all 22 additional Actions through the owned
PowerShell bridge and then the production CLI. These use `Worksheets.Add`,
`Worksheet.Name/Copy/Move/Delete/UsedRange`, `Range.Clear/ClearContents/ClearFormats`,
`Range.Merge/UnMerge/Sort/AutoFilter`, `Worksheet.ShowAllData/AutoFilterMode`,
`Range.EntireRow/EntireColumn.Insert/Delete`, and row/column sizing/AutoFit.
Literal search/replacement and value-only copying are implemented against observed
cells rather than implicit selection or clipboard state.

Verification includes worksheet order and copy content, shifted cell values after
row/column edits, full-row numeric sorting, filtered row visibility, range copy,
clear modes, merge/unmerge with preserved text, font/color/alignment/wrapping
read-back, and 32 additional statistical/lookup/date/text functions with independently
specified expected results. Negative cases include stale worksheet and destination
tokens, duplicate sheet names, deleting the last visible worksheet, merging nonempty
cells, and unmerging only part of a merged area. Persisted worksheet order and formula
XML are checked independently in the saved ZIP archive.

Region snapshots also observe formatting, merging, row/column visibility and sizes.
Worksheet tokens combine the sheet inventory with the complete bounded used-range
snapshot; structure changes invalidate cached worksheet locators and old tokens.
Worksheet structure operations require at most 1000 used cells, and each row/column
edit is limited to 100 rows/columns. Sort accepts constants, copyRange copies values,
and formula input still excludes cross-sheet references and dynamic arrays.

Repository entry: `scripts/validate/excel_common.py --output-dir <new-directory>`.
Generated production evidence is retained under `build/evidence/excel/` locally and
`build/common-production-01/report.json` in the authorized Windows test checkout.


## Desktop presentation verification

A desktop run on session 6 exposed a difference from SSH acceptance: the bound
workbook's child Window.Hwnd belonged to a visible top-level frame titled only
`WPS Office`. Searching top-level titles for the filename incorrectly rejected
that visible workbook. The demo now verifies the HWND reported from the exact
bound Workbook.Windows collection, its process ID, root visibility and size.
A regression test covers generic window titles, destroyed handles and mismatched
process ownership. The originally reported REGDB_E_CLASSNOTREG was later reproduced specifically
in the elevated desktop context; see the elevation comparison below. The window
detection fix is independent of COM registration availability.


## Console flash verification

`src/test/python/tests/windows/console_observer.py` observes new visible
console-class windows while running the actual desktop demo. A positive control
using the previous cmd wrapper created both a Windows Terminal window and a cmd
PseudoConsoleWindow. The same production demo launched with CREATE_NO_WINDOW
created neither. The remote diagnostic wrapper, rather than a missing flag in the
shared Word/Excel bridge, can therefore account for a flashing auxiliary console.

The repository now provides an explicit no-console PowerShell demo launcher. It
starts the Python client with CreateNoWindow at process creation, bypasses cmd,
and forwards stdout/stderr to the existing shell. Remote desktop verification is
started using pythonw.exe; it no longer starts a console PowerShell or cmd task.
Generated results are stored under build/evidence/excel/. The optional positive
control deliberately opens a short-lived console; ordinary probe runs do not.


## Same-basename open guard

The desktop flash regression originally reused `demo.xlsx` in different output
directories. WPS presented a same-filename dialog, which stalled openWorkbook.
Demo workbooks now include a UUID in their basename, independent of directory name.
During exact-file acquisition, the Excel bridge also checks whether another live
workbook already uses that basename. A different file is rejected before Open with
a definite DOCUMENT_OPEN_FAILED; the same exact file can still be reused. The
live_name_conflict.py test verifies the rejection and a subsequent successful open
within the same still-unbound Session.


## Elevated desktop COM registration comparison

The user's actual terminal was elevated. On the same desktop session 6 and same
64-bit PowerShell, a Limited task resolved KET.Application and activated WPS 12.0;
a Highest task returned a null type and reproduced CLSID 00000000-0000-0000-0000-000000000000,
80040154 REGDB_E_CLASSNOTREG. Registration existed under HKCU Software\Classes
but not HKLM Software\Classes. This matches Windows COM's exclusion of per-user
registration for elevated processes:
https://devblogs.microsoft.com/oldnewthing/20190801-00/?p=102745

A shared read-only availability check now runs before the desktop launcher starts
Python and before the bridge prepares a document acquisition. It reports the
need for a non-administrator PowerShell explicitly, without registering classes,
using a hardcoded CLSID, or changing application privileges. Regression evidence
covers both the elevated fail-fast path (no demo directory created) and the
non-elevated complete visible demo with zero new helper consoles.

## Persistence extension

Creation, first save/Save As and export are now admitted. See [three-application native persistence evidence](../../persistence/EVIDENCE.md) for the tested native APIs, limits and validation.
