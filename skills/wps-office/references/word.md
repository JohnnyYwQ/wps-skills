# Word Document Workflows

Use this reference for document lifecycle, content, formatting, navigation, templates, comments, bookmarks, and tracked revisions. Read `action-manifest.json` for the exact parameter schema, result schema, prerequisites, and risk before every invocation.

Run one WPS Action per process from the WPS Skill Package directory:

```bash
python3 scripts/wps.py invoke '{"action":"getActiveDocument","params":{},"timeout_ms":30000}'
```

## Create, Inspect, Revise, and Save a Document

For a typical editing and proofreading workflow, invoke these Actions sequentially:

1. Use `createDocument`, or use `openDocument` with an existing path.
2. Use `insertText` to add initial content.
3. Use `getDocumentParagraphs` to obtain one-based paragraph indexes and exact character ranges.
4. Use `findInDocument` to locate text without changing it.
5. After explicit confirmation, use `replaceRange` for a precise replacement. Use `enableTrackChanges` first when the edit must remain reviewable.
6. Use `getTrackChangesStatus` to verify revision tracking and the current revision count.
7. Use the common `save` Action to persist the document.

Each step is a separate `scripts/wps.py invoke` process. Stop at the first failure and do not assume later changes ran. Character ranges use a zero-based start and exclusive end; paragraph indexes are one-based.

## Choose a Word Action

The catalog below is derived from the manifest's `word_workflows` reference group. The manifest remains the source of truth.

<!-- word-actions:start -->
- Document lifecycle and context: `createDocument`, `openDocument`, `closeDocument`, `getActiveDocument`, `getOpenDocuments`, `switchDocument`, `getDocumentText`, `getDocumentStats`, `getDocumentParagraphs`
- Content and structure: `insertText`, `insertTable`, `insertImage`, `insertHyperlink`, `insertPageBreak`, `insertSectionBreak`, `insertHeader`, `insertFooter`, `generateTOC`
- Find, template, and range operations: `findInDocument`, `findReplace`, `smartFillField`, `replaceRange`
- Formatting: `setFont`, `setTextColor`, `applyStyle`, `setParagraph`, `setLineSpacing`, `setPageSetup`
- Bookmarks, comments, and revisions: `insertBookmark`, `getBookmarks`, `replaceBookmarkContent`, `addComment`, `getComments`, `enableTrackChanges`, `getTrackChangesStatus`
<!-- word-actions:end -->

Use `findInDocument` before a range edit so the returned `start` and `end` values identify the intended occurrence. `replaceRange` reports both the original and resulting range. Do not reuse old positions after a replacement changes document length; locate the next target again.

Use `smartFillField` for templates whose value belongs after a label, colon, underline, or placeholder. Use `replaceBookmarkContent` when the template exposes a stable bookmark. Both operations preserve the label or bookmark identity, unlike an unrestricted find-and-replace.

## Validation, Safety, and Results

The Runner rejects invalid parameters before the Add-in receives an Action. Correct `INVALID_PARAMS` errors instead of retrying unchanged input. A WPS-side exception becomes `WPS_ACTION_FAILED`; preserve its message and do not report an empty document or empty search result unless the Action succeeded with structured data.

The following Word Actions are destructive and require explicit confirmation plus `"confirmed":true`:

- `findReplace` can replace or delete one or many matches.
- `replaceRange` overwrites an exact character range.
- `replaceBookmarkContent` overwrites the current bookmark contents.
- `smartFillField` can replace an existing placeholder, underline field, or value after a colon.
- `closeDocument` can save or discard pending document state as requested.

Explain the exact replacement or close consequence before asking for confirmation. Ordinary writes such as inserting text, adding comments, changing formatting, or enabling revision tracking follow the user's current editing request without a confirmation marker.
