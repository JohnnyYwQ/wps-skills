---
status: accepted
---

# Allow explicit retirement of legacy capabilities

The WPS Skill Package preserves stable WPS behavior rather than every legacy MCP entry. A legacy capability may be retired only when the migration ledger names it and records that it is MCP-runtime-only, has a broken or ambiguous contract, or is semantically covered by canonical WPS Actions; this supersedes ADR-0003's absolute capability-equivalence requirement while retaining WPS Actions as the compatibility seam.

## Accepted retirement boundary

- Retire generic `openFile`; `openDocument`, `openWorkbook`, and `openPresentation` are its canonical replacements.
- Retire generic `convertFormat`; `convertToPDF` and application save Actions cover the supported conversion workflows.
- Retire legacy `wps_excel_find_replace`; it incorrectly reused the Word `findReplace` contract, and Excel replacement can be expressed as a read/write workflow until a distinct Excel Action is justified.
- Keep the six already-retired MCP cache, unrestricted dynamic-call, and local-proofreading tools retired because they are not stable WPS Actions.
- Preserve Excel cell comments through `addCellComment` and PowerPoint image insertion through `insertPptImage`; these are corrected mappings, not retired capabilities.
- Standardize `addArrow` on the JavaScript Add-in's start/end coordinate contract and retire the conflicting PowerShell bounding-box form.

## Consequences

Retirement is an explicit reviewed decision, not a shortcut for difficult migration work. The checked-in migration ledger, behavioral tests, and Git history retain the evidence needed to explain each exclusion after the Legacy MCP Runtime is removed.
