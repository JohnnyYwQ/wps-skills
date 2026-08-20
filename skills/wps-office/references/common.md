# Common File Workflows

Use this reference for cross-application readiness, selection, file lifecycle, and PDF conversion. Read `action-manifest.json` for the exact parameter schema, result schema, prerequisites, and risk before every invocation.

Run one WPS Action per process from the WPS Skill Package directory:

```bash
python3 scripts/wps.py invoke '{"action":"getAppInfo","params":{},"timeout_ms":30000}'
```

## Open, Edit, Save, and Close

1. Choose the application-specific open Action: `openDocument`, `openWorkbook`, or `openPresentation`. Pass a nonblank file path.
2. Invoke application Actions to inspect or edit the active file.
3. Use `save` to persist to the current path, or call `getAppInfo` and use its `appType` with `saveAs` plus an explicit output `path`. The output extension selects an application-specific format.
4. Use `closeDocument`, `closeWorkbook`, or `closePresentation`. Set `saveChanges` to `true` to save before closing or `false` to explicitly discard pending changes.

Each close Action is destructive because it can discard or finalize pending state. Explain the close consequence and obtain explicit confirmation before invoking it with `"confirmed":true`.

## Convert to PDF

Use `convertToPDF` with an explicit `outputPath` ending in `.pdf`. This Action works with the active Word document, Excel workbook, or PowerPoint presentation and returns `appType`, `sourcePath`, and `outputPath`.

`save`, PDF conversion, and `saveAs` are destructive because they can overwrite files. Explain the current or output path and obtain explicit confirmation before including `"confirmed":true`. The Runner normalizes input and output paths before sending the same absolute paths to WPS. Unsupported extensions, missing active files, invalid paths, and WPS failures are returned as structured Runner errors; do not retry non-retryable failures unchanged.

## Other Common Actions

Use `ping` or `wireCheck` for connection diagnostics, `getAppInfo` for the active application, and `getSelectedText` or `setSelectedText` for cross-application selection workflows.

<!-- common-actions:start -->
- Connection and context: `ping`, `wireCheck`, `getAppInfo`
- Selection: `getSelectedText`, `setSelectedText`
- Persistence and conversion: `save`, `saveAs`, `convertToPDF`
<!-- common-actions:end -->
