# Excel Core Workflows

Use this reference for workbook, worksheet, cell, range, formula, and basic data operations. Read `action-manifest.json` for the exact parameter schema, result schema, prerequisites, and risk before every invocation.

Run one WPS Action per process from the WPS Skill Package directory:

```bash
python3 scripts/wps.py invoke '{"action":"getSheetList","params":{},"timeout_ms":30000}'
```

## Complete Read, Write, Formula, and Save Workflow

For a typical workbook update, invoke these Actions sequentially and pass returned values into later calls when needed:

1. Use `getActiveWorkbook` to inspect the active workbook and available sheets.
2. Use `switchSheet` with `{"sheet":"Summary"}` to select the target sheet.
3. Use `setRangeData` with a rectangular two-dimensional `data` array to write values.
4. Use `setFormula` with a `range` and an Excel `formula` beginning with `=`.
5. Use `getRangeData` to verify the stored values and calculated result.
6. Use the common `save` Action to persist the workbook.

Each step is a separate `scripts/wps.py invoke` process. Stop the workflow on the first failure; do not assume that later Actions ran.

## Choose a Core Action

The catalog below is derived from the manifest's `excel_core` reference group. The manifest remains the source of truth for each Action's semantics and contract.

<!-- excel-core-actions:start -->
- Workbook: `getActiveWorkbook`, `openWorkbook`, `getOpenWorkbooks`, `switchWorkbook`, `closeWorkbook`, `createWorkbook`
- Worksheet: `createSheet`, `deleteSheet`, `renameSheet`, `copySheet`, `getSheetList`, `switchSheet`, `moveSheet`
- Cell and range: `getCellValue`, `setCellValue`, `getRangeData`, `setRangeData`, `getCellInfo`, `clearRange`, `getSelection`, `insertRows`, `insertColumns`, `deleteRows`, `deleteColumns`
- Formula: `getFormula`, `setFormula`, `autoSum`, `evaluateFormula`
- Basic data: `cleanData`, `removeDuplicates`, `sortRange`, `findInSheet`, `replaceInSheet`
<!-- excel-core-actions:end -->

For formula-only calculations, `evaluateFormula` returns the calculated value and can optionally write the formula to a target cell. For a simple total written into the workbook, `autoSum` accepts the source range and target cell. Both formula strings are governed by the manifest contract.

## Validation and Results

The Runner validates Action parameters before the Add-in receives an Action. Correct `INVALID_PARAMS` errors instead of retrying unchanged input.

Successful reads always return structured `data`. An empty range is represented as `{"data":[]}` rather than an absent result. A WPS-side exception becomes `WPS_ACTION_FAILED`; preserve its message and do not treat it as an empty success.

For every destructive Action, explain the specific deletion, replacement, clearing, or close consequence and obtain explicit confirmation before setting `"confirmed":true`.
