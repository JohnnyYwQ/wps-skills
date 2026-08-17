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

## Advanced Design Workflows

Use this section for shape composition, data visualization, diagrams, visual styling, animation, transitions, themes, masters, and cross-presentation editing. Build and inspect in small steps: create the object, apply its style or behavior, then read the resulting slide context before saving.

<!-- powerpoint-advanced-actions:start -->
- Shapes and composition: `addShape`, `deleteShape`, `getShapes`, `setShapeStyle`, `setShapeFill`, `setShapeText`, `setShapePosition`, `setShapeBorder`, `setShapeGradient`, `setShapeShadow`, `setShapeTransparency`, `setShapeRoundness`, `setShapeFullStyle`, `alignShapes`, `distributeShapes`, `smartDistribute`, `groupShapes`, `duplicateShape`, `setShapeZOrder`, `addConnector`, `addArrow`, `setFontColor`
- Charts and visual components: `insertPptChart`, `setPptChartData`, `setPptChartStyle`, `createProgressBar`, `createGauge`, `createMiniCharts`, `createDonutChart`, `createGrid`, `createStyledTable`, `createKpiCards`
- Diagrams and polish: `createFlowChart`, `createOrgChart`, `createTimeline`, `applyColorScheme`, `autoBeautifySlide`, `beautifySlide`, `beautifyAllSlides`, `autoLayout`, `addTitleDecoration`, `addPageIndicator`, `unifyFont`
- Animation and transitions: `addAnimation`, `addAnimationPreset`, `addEmphasisAnimation`, `getAnimations`, `setAnimationOrder`, `removeAnimation`, `setSlideTransition`, `applyTransitionToAll`, `removeSlideTransition`
- Themes, masters, and 3D: `getSlideMaster`, `setMasterBackground`, `addMasterElement`, `setSlideTheme`, `set3DRotation`, `set3DDepth`, `set3DMaterial`, `create3DText`, `setBackgroundGradient`
- External content and presentation controls: `insertSlidesFromFile`, `replacePptImage`, `findPptText`, `replacePptText`, `addPptHyperlink`, `removePptHyperlink`, `setSlideNumber`, `setPptFooter`, `setPptDateTime`, `startSlideShow`, `endSlideShow`
<!-- powerpoint-advanced-actions:end -->

For `insertSlidesFromFile`, provide the source presentation in `filePath`; optionally set `afterIndex`, `slideStart`, and `slideEnd`. For `replacePptImage`, provide `filePath` and the one-based `slideIndex`, plus the target image's `name` or `shapeIndex`; the replacement preserves the target geometry and rotation. Both operations return diagnostic errors from WPS when the source or target cannot be used.

## Validation, Safety, and Results

The Runner rejects zero or negative indexes, missing paths, incomplete dimensions, and missing object identifiers before the Add-in receives an Action. Correct `INVALID_PARAMS` errors instead of retrying unchanged input. A WPS-side exception becomes `WPS_ACTION_FAILED`; preserve its message and do not treat it as an empty presentation or slide.

The following core Actions are destructive and require explicit confirmation plus `"confirmed":true`:

- `deleteSlide`, `deleteTextBox`, and `deletePptImage` remove presentation content.
- `closePresentation` can discard pending presentation state.
- `exportSlideAsImage` can overwrite an existing output file.

Explain the exact deletion, close consequence, or output path before asking for confirmation. Ordinary authoring writes—including changing title/content text, notes, a table cell, layout, size, or background—follow the user's current editing request without a confirmation marker.

The advanced Actions that remove, replace, or reorder presentation content are destructive where declared by the manifest and require explicit confirmation plus `"confirmed":true`; ordinary design edits remain normal writes. Read the manifest before invoking an Action with complex arrays, file paths, or object identifiers.
