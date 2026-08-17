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

## Core Action Catalog

The manifest is authoritative when a table summary and its schema differ.

### Workbooks

| WPS Action | Purpose | Risk |
|---|---|---|
| `getActiveWorkbook` | Inspect the active workbook and its worksheets. | read |
| `openWorkbook` | Open a workbook path. | write |
| `getOpenWorkbooks` | List open workbooks. | read |
| `switchWorkbook` | Activate a named workbook. | write |
| `closeWorkbook` | Close a workbook. | destructive |
| `createWorkbook` | Create a workbook. | write |

### Worksheets

| WPS Action | Purpose | Risk |
|---|---|---|
| `createSheet` | Create a worksheet. | write |
| `deleteSheet` | Delete a worksheet. | destructive |
| `renameSheet` | Rename a worksheet. | write |
| `copySheet` | Copy a worksheet. | write |
| `getSheetList` | List worksheets and the active sheet. | read |
| `switchSheet` | Activate a worksheet. | write |
| `moveSheet` | Reorder a worksheet. | write |

### Cells, Ranges, Rows, and Columns

| WPS Action | Purpose | Risk |
|---|---|---|
| `getCellValue` | Read a cell value, display text, and formula. | read |
| `setCellValue` | Write a cell value. | write |
| `getRangeData` | Read a range as a two-dimensional array. | read |
| `setRangeData` | Write a two-dimensional array from a starting range. | write |
| `getCellInfo` | Read cell value, formula, number format, and basic style. | read |
| `clearRange` | Clear range content or formatting. | destructive |
| `getSelection` | Read the selected range address and dimensions. | read |
| `insertRows` | Insert worksheet rows. | write |
| `insertColumns` | Insert worksheet columns. | write |
| `deleteRows` | Delete worksheet rows. | destructive |
| `deleteColumns` | Delete worksheet columns. | destructive |

### Formulas

| WPS Action | Purpose | Risk |
|---|---|---|
| `getFormula` | Read a cell formula and local formula. | read |
| `setFormula` | Write a formula to a cell or range. | write |
| `autoSum` | Write a SUM formula to a target cell and return its numeric result. | write |
| `evaluateFormula` | Evaluate a formula, optionally by writing it to a target cell. | write |

### Basic Data Processing

| WPS Action | Purpose | Risk |
|---|---|---|
| `cleanData` | Apply declared cleanup operations to a range. | write |
| `removeDuplicates` | Remove duplicate rows. | destructive |
| `sortRange` | Sort a range. | write |
| `findInSheet` | Find matching cells. | read |
| `replaceInSheet` | Replace matching cell content. | destructive |

## Validation and Results

The Runner validates Action parameters before the Add-in receives an Action. Correct `INVALID_PARAMS` errors instead of retrying unchanged input.

Successful reads always return structured `data`. An empty range is represented as `{"data":[]}` rather than an absent result. A WPS-side exception becomes `WPS_ACTION_FAILED`; preserve its message and do not treat it as an empty success.

For every destructive Action, explain the specific deletion, replacement, clearing, or close consequence and obtain explicit confirmation before setting `"confirmed":true`.
