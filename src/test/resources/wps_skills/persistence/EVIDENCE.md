# Native persistence admission, 2026-09-05

The production surface contains Word 14, Excel 34 and PPT 37 Actions (85 total).
New admissions: Word `saveAs`; Excel `createWorkbook`, `saveAs`, `exportPdf`;
PPT `createPresentation`, `saveAs`, `exportPdf`, `exportSlideImage`.

Tests use ordinary-permission interactive Windows desktop tasks, native WPS COM
and the production Session Client/Host. The Runtime has no pywin32 dependency.
The type-library snapshots remain unchanged and are not shipped as runtime code.

`src/test/python/tests/persistence/live_acceptance.py` verifies native creation,
unsaved `save` rejection, first save, repeated Save As on the same retained object,
editing and ordinary save after migration, old-file content preservation,
existing-target and missing-parent rejection, hard-link aliases, retained source
and destination lease conflicts, cleanup/reacquisition and persisted OOXML text.
Excel/PPT PDF exports are checked in both unsaved and saved states; PPT PNG is
checked for format and exact dimensions. An invalid slide ID is rejected.
The final extension also checks that an unbound Session can explicitly create a
new document after lease conflicts, without changing the first Session's target.

Native findings:

- Word `Documents.Add` and `SaveAs(path,16)` produce ordinary DOCX. WPS reports
  `SaveFormat=12` (XML document) after using default-document format 16.
- Excel `Workbooks.Add(-4167)` creates one worksheet; `SaveAs(path,51)` produces
  ordinary XLSX. `ExportAsFixedFormat(0,path)` produces PDF without saving.
- PPT `Presentations.Add(-1)` creates zero slides; `SaveAs(path,24)` produces PPTX.
  Although declared by the type library, `ExportAsFixedFormat` rejects PowerShell
  binding and returns native `DISP_E_MEMBERNOTFOUND` through IDispatch. The admitted
  PDF implementation uses `SaveCopyAs(path,32)`: PDF signature, original locator,
  observed content and Saved state are verified. It never calls `SaveAs` to PDF.
- `Slide.Export(path,'PNG',width,height)` produces the requested PNG dimensions.
- Word liveness must not final-release its retained Application RCW; doing so
  breaks later persistence that uses that same application reference.

All six native suites passed: new persistence, Word demo acceptance, Excel base
and common Actions, PPT base and common Actions. No new visible auxiliary console
was observed during these suites. Non-WPS regression: 214 tests pass on macOS and
Windows. All three updated Skills pass the Skill validator in UTF-8 mode.

Reports and generated documents are retained under ignored
`build/evidence/persistence-20260905/`; they are not part of the Skill packages.
This is evidence for the tested WPS installation and bounded document features,
not a claim that every WPS version or arbitrary Office document is supported.
