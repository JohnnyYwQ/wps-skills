# Excel Core Workflows

Use this reference for workbook, worksheet, cell, range, formula, and basic data operations. Read `action-manifest.json` for the exact parameter schema, result schema, prerequisites, and risk before every invocation.

Run one WPS Action per process from the WPS Skill Package directory:

```bash
python3 scripts/wps.py invoke '{"action":"getSheetList","params":{},"timeout_ms":30000}'
```

## Complete Read, Write, Formula, and Save Workflow

For a typical workbook update, invoke these Actions sequentially and pass returned values into later calls when needed:

1. Use `openWorkbook` for an existing file, or `getActiveWorkbook` to inspect the active workbook and available sheets.
2. Use `switchSheet` with `{"sheet":"Summary"}` to select the target sheet.
3. Use `setRangeData` with a rectangular two-dimensional `data` array to write values.
4. Use `setFormula` with a `range` and an Excel `formula` beginning with `=`.
5. Use `getRangeData` to verify the stored values and calculated result.
6. Use the common `save` Action to persist the workbook.

Each step is a separate `scripts/wps.py invoke` process. Stop the workflow on the first failure; do not assume that later Actions ran.

Open workbooks only through `openWorkbook`. For spreadsheet replacement, use `findInSheet` and `replaceInSheet`; no Word replacement contract is an Excel Action.
To create a PDF, use the common `convertToPDF` Action; for other supported output workflows, use `save` or `saveAs` with the application's documented format handling.

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

## Analysis and Presentation Workflow

Use advanced Actions after the core workflow has established the source data:

1. Use `setCellStyle`, `setNumberFormat`, or `addConditionalFormat` to make the data readable.
2. Use `createChart` for a visualization, or `createPivotTable` for grouped analysis. Retain the returned object name for later updates.
3. Use `updateChart` or `updatePivotTable` with that name to change the object. These Actions require either the object name or the alternate identifier declared by the manifest. When a pivot was created on `destinationSheet`, pass the same worksheet as `sheet` when updating it.
4. Use `getExcelContext`, `getConditionalFormats`, `getDataValidations`, or `getNamedRanges` to read the resulting workbook state.
5. Use the common `save` Action to persist the completed workbook.

For presentation-only changes, omit the pivot step. For analysis-only changes, omit the chart step. Each step remains one separate Runner process.

## Choose an Advanced Action

The catalog below is derived from the manifest's `excel_advanced` reference group.

<!-- excel-advanced-actions:start -->
- Formatting and layout: `setCellFormat`, `setCellStyle`, `mergeCells`, `unmergeCells`, `setColumnWidth`, `setRowHeight`, `autoFitColumn`, `autoFitRow`, `autoFitAll`, `copyFormat`, `clearFormats`, `setBorder`, `setNumberFormat`, `freezePanes`, `unfreezePanes`, `hideRows`, `hideColumns`, `showRows`, `showColumns`, `groupRows`, `groupColumns`, `wrapText`, `setZoom`, `setPrintArea`
- Validation, links, comments, and names: `addConditionalFormat`, `removeConditionalFormat`, `getConditionalFormats`, `addDataValidation`, `removeDataValidation`, `getDataValidations`, `setHyperlink`, `createNamedRange`, `deleteNamedRange`, `getNamedRanges`, `addCellComment`, `deleteCellComment`, `getCellComments`
- Charts and images: `createChart`, `updateChart`, `exportChartAsImage`, `exportRangeAsImage`, `insertExcelImage`
- Analysis and pivots: `autoFilter`, `createPivotTable`, `updatePivotTable`, `diagnoseFormula`, `calculateSheet`, `refreshLinks`, `consolidate`, `setArrayFormula`, `subtotal`
- Advanced data movement: `copyRange`, `pasteRange`, `fillSeries`, `transpose`, `textToColumns`
- Protection and context: `protectSheet`, `unprotectSheet`, `protectWorkbook`, `lockCells`, `getContext`, `getExcelContext`
<!-- excel-advanced-actions:end -->

`getExcelContext` is the canonical context-aware read for callers that need workbook, worksheet, selection, used-range dimensions, and header information together. `getContext` remains available with the same structured result for compatibility with the existing formula-generation workflow.

## Validation and Results

The Runner validates Action parameters before the Add-in receives an Action. Correct `INVALID_PARAMS` errors instead of retrying unchanged input.

Successful reads always return structured `data`. An empty range is represented as `{"data":[]}` rather than an absent result. A WPS-side exception becomes `WPS_ACTION_FAILED`; preserve its message and do not treat it as an empty success.

For every destructive Action, explain the specific deletion, replacement, clearing, or close consequence and obtain explicit confirmation before setting `"confirmed":true`.

Image exports are destructive because their output path can overwrite an existing file. Pasting over cells and replacing an existing cell comment also require confirmation, as do clearing formats, deleting names/comments/rules, and removing validation even when cell values remain intact.
