"""Word Action examples and published constraints."""


_WORD_EXAMPLE_PARAMS = {
    "createDocument": {},
    "openDocument": {"path": "C:/docs/report.docx"},
    "writeContent": {
        "anchor": {"kind": "documentEnd"},
        "blocks": ({
            "kind": "paragraph",
            "runs": ({
                "text": "你好, world",
                "format": {
                    "westernFontFamily": "Times New Roman",
                    "eastAsiaFontFamily": "Microsoft YaHei",
                },
            },),
        },),
    },
    "inspectDocument": {
        "scope": {"kind": "document"},
        "limits": {
            "maxTextCharacters": 4096,
            "maxParagraphs": 128,
            "maxRuns": 512,
        },
    },
    "findContent": {
        "query": {
            "scope": {"kind": "document"},
            "text": "Hello",
            "caseSensitive": False,
            "wholeWord": True,
        },
        "limit": 50,
    },
    "replaceContent": {
        "target": {
            "kind": "query",
            "query": {
                "scope": {"kind": "document"},
                "text": "old",
                "caseSensitive": True,
                "wholeWord": True,
            },
            "expectedMatchCount": 1,
        },
        "replacement": {
            "kind": "text",
            "runs": ({"text": "new"},),
        },
    },
    "insertTable": {
        "anchor": {"kind": "documentEnd"},
        "data": (("Name", "Value"), ("A", "1")),
        "headerRow": True,
    },
    "insertImage": {
        "anchor": {"kind": "documentEnd"},
        "source": {"kind": "file", "path": "C:/images/logo.png"},
        "placement": {"kind": "inline"},
        "size": {"kind": "intrinsic"},
        "alternativeText": {"kind": "decorative"},
    },
    "setHeaderFooter": {
        "sections": {"kind": "all", "revision": "revision-1"},
        "updates": ({
            "area": "header",
            "variant": "primary",
            "operation": {"kind": "replace", "text": "Report"},
        },),
    },
    "setPageLayout": {
        "sections": {"kind": "all", "revision": "revision-1"},
        "layout": {"orientation": "landscape"},
    },
    "insertBreak": {
        "anchor": {"kind": "documentEnd"},
        "type": "page",
    },
    "save": {},
    "saveAs": {
        "outputPath": "C:/docs/output.docx",
        "overwritePolicy": "renameIfExists",
    },
    "exportPdf": {
        "outputPath": "C:/docs/output.pdf",
        "overwritePolicy": "renameIfExists",
    },
}

_WORD_CONSTRAINTS = {
    "createDocument": (
        "Creates one blank unsaved document and never accepts a path.",
        "Binding and Lease commit before a successful response is observable.",
    ),
    "openDocument": (
        "The path is an absolute existing .docx resolved by stable identity.",
        "Active UI state, display name, and open order never select the document.",
        "Document observation preserves its saved/modified state; no hidden save or reset.",
    ),
    "writeContent": (
        "Blocks are either one inline text block or a paragraph/heading sequence.",
        "fontFamily is a unified setting and cannot be mixed in one run with westernFontFamily or eastAsiaFontFamily.",
        "All limits, revision, bounds, alignment, and protection checks precede writing.",
    ),
    "inspectDocument": (
        "Every returned fact belongs to one coherent Content Revision.",
        "Body output may be bounded, but section/layout/header-footer facts never truncate.",
        "Header/footer observation reads live in-memory XML without materializing COM stories.",
    ),
    "findContent": (
        "Search is literal, bounded, non-wrapping, and non-overlapping.",
        "All matches and remaining range belong to one Content Revision.",
    ),
    "replaceContent": (
        "All targets and expected match counts are preflighted before the first write.",
        "Multi-match replacement executes from document end toward start and is never replayed.",
    ),
    "insertTable": (
        "Data is a non-empty rectangular plain-text matrix within hard ceilings.",
        "Success uses read-back cell data rather than request echo.",
    ),
    "insertImage": (
        "Only staged embedded PNG/JPEG files are accepted; URLs and links are forbidden.",
        "Success reads back media hash, dimensions, placement, anchor, and alternative text.",
    ),
    "setHeaderFooter": (
        "Section selection is revision-bound and update area/variant pairs are unique.",
        "linkToPrevious never targets section 0; selecting all sections therefore forbids it.",
        "All selected stories are preflighted before any write and read back after mutation.",
    ),
    "setPageLayout": (
        "At least one layout field is explicit and a margin object always contains all four sides.",
        "All selected sections are preflighted and observed in normalized points.",
    ),
    "insertBreak": (
        "The handler duplicates and collapses the Body Anchor before insertion.",
        "Only page and explicit section break types are supported and read back.",
    ),
    "save": (
        "Uses only the Binding's existing .docx locator and accepts no path.",
        "An already-saved document may succeed only after the same artifact verification.",
    ),
    "saveAs": (
        "Output is fixed .docx and never overwrites. renameIfExists tries the requested name then numbered siblings within the same Action (at most 20 attempts).",
        "Only definite pre-write name conflicts retry; unknown outcomes, permission errors and exhausted names stop. Read the actual artifact.path and outputResolution.",
        "The same live document remains bound while destination Lease migration has no gap.",
    ),
    "exportPdf": (
        "Exports the complete document to PDF without saving or retargeting the Word document.",
        "renameIfExists retries only definite name conflicts against numbered siblings (at most 20 attempts); it never overwrites or retries unknown effects. Read artifact.path and outputResolution.",
        "Success proves one readable artifact while revision and document state remain unchanged.",
    ),
}
