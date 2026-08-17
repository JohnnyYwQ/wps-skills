# PowerPoint Core Authoring Workflows

Use this reference for presentation lifecycle, slides, text, images, tables, speaker notes, and basic layout. Read `action-manifest.json` for the exact parameter schema, result schema, prerequisites, and risk before every invocation.

Run one WPS Action per process from the WPS Skill Package directory:

```bash
python3 scripts/wps.py invoke '{"action":"getActivePresentation","params":{},"timeout_ms":30000}'
```

## Create, Inspect, and Save a Presentation

For a typical authoring workflow, invoke these Actions sequentially:

1. Use `createPresentation`, or open an existing file.
2. Use `addSlide` with a layout plus optional initial title and content.
3. Use `setSlideContent` when the content placeholder needs a precise update.
4. Use `insertPptImage` to add a local image at an explicit position and size.
5. Use `insertPptTable` to add structured content, then update cells by the returned table name.
6. Use `getSlideInfo` to verify the current layout and observable shapes.
7. Use the common `save` Action to persist the presentation.

Each step is a separate `scripts/wps.py invoke` process. Stop at the first failure and do not assume later changes ran. Slide, shape, row, and column indexes are one-based.

## Choose a Core Action

The catalog below is derived from the manifest's `powerpoint_core` reference group. The manifest remains the source of truth.

<!-- powerpoint-core-actions:start -->
- Presentation lifecycle and context: `createPresentation`, `openPresentation`, `closePresentation`, `getActivePresentation`, `getOpenPresentations`, `switchPresentation`
- Slide lifecycle and context: `addSlide`, `deleteSlide`, `duplicateSlide`, `moveSlide`, `getSlideCount`, `getSlideInfo`, `switchSlide`
- Layout, backgrounds, and notes: `setSlideLayout`, `setSlideSize`, `setSlideBackground`, `setBackgroundColor`, `setBackgroundImage`, `getSlideNotes`, `setSlideNotes`
- Text and placeholders: `addTextBox`, `deleteTextBox`, `getTextBoxes`, `setTextBoxText`, `setTextBoxStyle`, `setSlideTitle`, `getSlideTitle`, `setSlideSubtitle`, `setSlideContent`
- Images: `insertPptImage`, `deletePptImage`, `setImageStyle`, `exportSlideAsImage`
- Tables: `insertPptTable`, `setPptTableCell`, `getPptTableCell`, `setPptTableStyle`, `setPptTableCellStyle`, `setPptTableRowStyle`
<!-- powerpoint-core-actions:end -->

Use names returned by image, text-box, and table creation Actions for later updates when possible. Table indexes count only table shapes on the slide; they are not raw shape indexes. Use the context reads before an update when the current presentation, slide, or object identity is uncertain.

## Validation, Safety, and Results

The Runner rejects zero or negative indexes, missing paths, incomplete dimensions, and missing object identifiers before the Add-in receives an Action. Correct `INVALID_PARAMS` errors instead of retrying unchanged input. A WPS-side exception becomes `WPS_ACTION_FAILED`; preserve its message and do not treat it as an empty presentation or slide.

The following core Actions are destructive and require explicit confirmation plus `"confirmed":true`:

- `deleteSlide`, `deleteTextBox`, and `deletePptImage` remove presentation content.
- `closePresentation` can discard pending presentation state.
- `exportSlideAsImage` can overwrite an existing output file.

Explain the exact deletion, close consequence, or output path before asking for confirmation. Ordinary authoring writes—including changing title/content text, notes, a table cell, layout, size, or background—follow the user's current editing request without a confirmation marker.

Advanced shapes, charts, diagrams, animation, transitions, themes, masters, external slide insertion, beautification, and in-place image replacement are outside this core slice.
