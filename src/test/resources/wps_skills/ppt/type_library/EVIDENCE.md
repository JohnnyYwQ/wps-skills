# WPS Presentation capability evidence

## Type-library snapshot

`wps_ppt_api.py` is the original, byte-preserved generated COM snapshot, moved from
repository root. SHA-256:
`e2b18e4db247cded255b6d56beb95be0ee095cd146011bf9b34dad8f08c488ea`.
It declares Upgrade WPS Presentation 3.0 Object Library (Beta), library GUID
`44720440-94BF-4940-926D-4F38FECF2A48`, version 3.0, generated 2026-08-28.
The Runtime and assembled Skill do not import or distribute it.

## Native findings (2026-09-05)

Tests run through `ssh win` with ordinary-permission interactive desktop tasks,
`pythonw.exe` and hidden child processes. No credentials, user account names or
absolute personal paths belong in this evidence. Generated reports live under
ignored `build/evidence/ppt/`.

- The current user's COM registry contains `KWPP.Application`. It activates WPS
  Presentation; do not infer an activatable class from an interface GUID.
- `Application.Version` reports `12.0`; the blank native fixture's build metadata
  reports `2052-12.1.0.28505`. These are different vendor version fields.
- Native Add, AddTextbox, TextRange.Text, SaveAs and Save were probed on disposable
  documents. Probe-only creation and first-save calls are not admitted Actions.
- `DocumentWindow.HWND` is absent from the installed IDispatch surface. Access by
  name and DISPID 2020 both fail with `DISP_E_MEMBERNOTFOUND`; IOleWindow is not
  supported. The declared DocumentWindow interface
  `91493457-5A91-11CF-8700-00AA0060263B` implements native get_HWND successfully.
  `ppt_window.ps1` queries that exact IID and calls slot 34 (snapshot vtable byte
  offset 272 on 64-bit, pointer-sized slot arithmetic). It returns the bound
  document's own HWND; no title matching or application-global HWND fallback.
- Native `Slides.FindBySlideID(256)` returned null while `Slides.Item(1).SlideID`
  returned 256 in that same collection. Production lookup enumerates only the
  bound Presentation's Slides and compares SlideID. The identity semantics are
  consistent with [Microsoft's SlideID documentation](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slides.findbyslideid);
  actual member availability is established by WPS tests, not Microsoft examples.
- Nested PPT collection reads can share one RCW. Releasing a nested helper's
  collection immediately detached the outer caller's reference and broke Add.
  Temporary PPT references are deduplicated and released once at the Action
  boundary; the Document/Application/Documents binding references are excluded.
- Reading a ZIP package with default sharing failed while WPS held the file open.
  PPT package validation now uses read access with ReadWrite/Delete sharing; the
  exclusive automation Lease still belongs to the shared coordinator.
- `demo-template.pptx` is a native blank, zero-slide, 960×540 point fixture with
  author and custom metadata cleared. The demo copies it to a unique path and
  performs every visible edit through the Action Session. This fixture copy is
  not a createPresentation or Save As capability.

- Mixed Latin/East Asian text returns an empty aggregate `Font.Name`. Native
  `NameAscii` and `NameFarEast` retain separate values. `formatText` therefore
  exposes `latinName` and `eastAsianName`; Arial and 宋体 were independently set
  and read back. Assigning a Latin-only font to East Asian text can be ignored by
  WPS, so mismatches remain an uncertain result instead of claiming success.
- Ordinary native `Save` replaces the backing file identity in this WPS build.
  PPT reuses the shared locator fence plus retained old/new identity claims;
  actual competing Sessions were rejected before and after Save, including a
  hard-link alias of the old file. Cleanup then allowed reacquisition of the
  saved file, with persisted content independently checked in the PPTX XML.

## Initial admission (15 Actions)

All 15 Actions were admitted on 2026-09-05 after `candidate-06` completed 33
successful calls and `candidate-demo-01` completed the visible two-slide demo.
The final production rerun is `production-01`; the complete console-observed
PowerShell launcher rerun is `desktop-01`. Both passed. Local and Windows
regression each passed 196 unittests. Standalone Skill build and relocated
package discovery passed. A separate desktop process imported the built Skill
from its bundled Runtime, opened the saved two-slide demo and verified its
window and saved state (`desktop-01/package-report.json`); the build includes no
generated COM type library.

`desktop-01/report.json` reports `newVisibleConsoles: []`; its demo report verifies
the bound Presentation's native window handle and a visible 1536×826 frame. The
normal cleanup path leaves the presentation open, with its changes explicitly
saved by the Task. A prior native positive-control test of the shared console
observer is documented in the Excel evidence; no positive-control console is
introduced into the PPT user demo.

| Admitted Actions | Native members and validation |
| --- | --- |
| openPresentation / getPresentationInfo | Presentations.Open, exact stable file identity/reuse, PageSetup, Saved, ReadOnly, native DocumentWindow HWND |
| listSlides / getSlideInfo | Slides.Item, SlideID, SlideIndex, Shapes.Item, Shape.Id, TextFrame.TextRange and font/geometry properties; bounded complete snapshots |
| addSlide / duplicateSlide / moveSlide / deleteSlide | Slides.Add(blank), Slide.Duplicate/MoveTo/Delete; count/ID/order readback and source-content comparison |
| addTextBox / addShape | Shapes.AddTextbox/AddShape; unique new shape, literal text and geometry readback |
| setShapeText / formatText | TextRange.Text and Font.NameAscii/NameFarEast/Size/Bold/Italic/Color.RGB; exact text and requested font readback |
| setShapeGeometry / deleteShape | Shape.Left/Top/Width/Height/Delete; requested dimensions or absent ID verification |
| save | Presentation.Save, Saved, nonempty ordinary PPTX, continuous locator and file-identity fences; independent persisted XML verification |

Candidate-only host and demo entries remain under `src/test/python/tests/ppt/`.
At this initial milestone, production discovery was exclusively
`PPT_PRODUCTION_CONTRACT_SET`; first save, creation, Save As, images, tables/charts,
notes/animations and export were not advertised. The extension below supersedes
that initial capability boundary. Tokens cover documented returned fields, not the complete native
slide model. File-name collisions with another open presentation are rejected
conservatively by this slice; no claim is made that native WPS has Excel's exact
same-name limitation.


## Common Contract Set extension (33 Actions, 2026-09-05)

18 additional Actions passed candidate native acceptance (`common-candidate-04`)
and the final production acceptance (`common-production-02`). Production now
explicitly admits 33 Actions. The original 15-Action native suite passed again
in `common-production-base-01`, including document binding and save fences.
Local and Windows regression each passed 204 unittests.

Reports and saved fixtures are retained under ignored
`build/evidence/ppt/common/`; the remote runner report is
`common-final-runner-report.json`, and individual suites are under its `build/`
subdirectory. `common-desktop-01/report.json` reports
`launcher.newVisibleConsoles: []`. Its enhanced two-slide demo adds native table
text and speaker notes, explicitly saves, and leaves its bound window open.
The independent bundled-Skill process passed
`common-desktop-01/package-report.json`: it imported the assembled Runtime,
opened that saved demo, and read back its notes and table. Package discovery
also reported the full production index.

| Additional admitted Actions | Native members and validation |
| --- | --- |
| getShapeStyle / formatShape | Shape.Fill, Shape.Line, TextFrame, individual ParagraphFormat; requested style readback, stale-token and contradictory-patch rejection |
| formatParagraph / setTextBoxLayout | ParagraphFormat.Alignment, SpaceBefore/After, LineRuleBefore/After, Bullet.Visible; TextFrame margins, VerticalAnchor, WordWrap; every paragraph and requested textbox field checked |
| renameShape / setShapeOrder | Shape.Name and ZOrder; stable ID/name and full slide stacking readback |
| alignShapes / distributeShapes | Explicit Left/Top on stable IDs; six collective-bounds alignments and both equal-edge-gap distributions verified without selection routing |
| getSlideSettings / setSlideSettings | Slide.Name, SlideShowTransition.Hidden, FollowMasterBackground, Background.Fill; own observation token and setting readback |
| getSlideNotes / setSlideNotes | Unique NotesPage body placeholder (type 2), TextFrame.TextRange.Text; separate notes token, exact text readback and saved notes XML |
| findText / replaceText | Literal top-level shape text search; native Characters ranges edited in reverse; UTF-16 offsets including emoji, exact replacement count and final text |
| addImage | Shapes.AddPicture with LinkToFile false / SaveWithDocument true; actual PNG/JPEG validation, unique picture ID/geometry, embedded PNG bytes checked in saved PPTX |
| addTable / readTable / writeTable | Shapes.AddTable, Table.Cell.TextFrame.TextRange.Text; complete bounded text matrix, dimension/stale-token rejection, reacquired Session readback and saved table XML |

Native findings determine several deliberate boundaries:

- WPS ignores a border color assignment while the border is hidden, and hidden
  color cannot be read back reliably. `lineColor` therefore enables the border;
  requesting a color and `lineVisible: false` together is rejected before any
  write. Fill color/transparency similarly enables fill, and a conflicting
  hidden-fill patch is rejected. Hidden border and hidden/non-solid fill colors
  read as null; setting fill color explicitly chooses a solid fill.
- Paragraph spacing is explicitly set to points by clearing its native line-rule
  flags. Tokens include individual paragraph fields, avoiding aggregate mixed
  values falsely validating the result.
- Native Shape.Duplicate returned without creating another shape in the probe.
  Shape duplication is therefore not admitted; the separately tested
  Slide.Duplicate remains supported.
- The schema includes zOrder and autoShapeType before immutable ActionContracts
  are constructed. A regression checks returned snapshots through those actual
  constructed contracts. Table result comparisons account for the shared
  ControllerResult's frozen arrays.
- These additions retain explicit scope limits: top-level ungrouped shapes,
  unrotated multi-shape alignment/distribution, literal text tables with fixed
  dimensions, and unique notes-body placeholders. Table tokens cover IDs,
  dimensions and text, not merged-cell geometry or formatting. Full notes/cell
  text replacement can reset rich formatting. Read-after-write mismatches remain
  uncertain outcomes and must not be automatically retried.

Creation, first save, Save As, shape duplication, charts, animations and export
remain outside the production Contract Set. Full limits and token selection are
specified in the Action Contracts and the Skill's `references/common.md`.
