"""Shared Word parameter and result schemas."""


def _object(properties, required=(), **keywords):
    schema = {
        "type": "object",
        "properties": properties,
        "required": tuple(required),
        "additionalProperties": False,
    }
    schema.update(keywords)
    return schema


def _array(items, *, minimum=0, maximum=None, **keywords):
    schema = {
        "type": "array",
        "items": items,
        "minItems": minimum,
    }
    if maximum is not None:
        schema["maxItems"] = maximum
    schema.update(keywords)
    return schema


def _string(*, minimum=0, maximum=None, enum=None, **keywords):
    schema = {"type": "string", "minLength": minimum}
    if maximum is not None:
        schema["maxLength"] = maximum
    if enum is not None:
        schema["enum"] = tuple(enum)
    schema.update(keywords)
    return schema


def _integer(*, minimum=None, maximum=None, **keywords):
    schema = {"type": "integer"}
    if minimum is not None:
        schema["minimum"] = minimum
    if maximum is not None:
        schema["maximum"] = maximum
    schema.update(keywords)
    return schema


def _number(*, minimum=None, maximum=None, **keywords):
    schema = {"type": "number"}
    if minimum is not None:
        schema["minimum"] = minimum
    if maximum is not None:
        schema["maximum"] = maximum
    schema.update(keywords)
    return schema


def _const(value):
    return {"const": value}


def _nullable(schema):
    return {"oneOf": (schema, {"type": "null"})}


def _length_between_points(minimum_pt, maximum_pt):
    units_per_point = {
        "pt": 1,
        "in": 1 / 72,
        "cm": 2.54 / 72,
        "mm": 25.4 / 72,
    }
    return {
        "oneOf": tuple(
            _object(
                {
                    "value": _number(
                        minimum=minimum_pt * factor,
                        maximum=maximum_pt * factor,
                    ),
                    "unit": _const(unit),
                },
                ("value", "unit"),
            )
            for unit, factor in units_per_point.items()
        )
    }


CONTENT_REVISION = _string(minimum=1, maximum=128)

CONTENT_RANGE = _object(
    {
        "start": _integer(minimum=0),
        "end": _integer(minimum=0),
        "revision": CONTENT_REVISION,
    },
    ("start", "end", "revision"),
    **{"x-ordered": ("start", "end")},
)

NONEMPTY_CONTENT_RANGE = _object(
    {
        "start": _integer(minimum=0),
        "end": _integer(minimum=0),
        "revision": CONTENT_REVISION,
    },
    ("start", "end", "revision"),
    **{"x-strictOrdered": ("start", "end")},
)

BODY_ANCHOR = {
    "oneOf": (
        _object({"kind": _const("documentStart")}, ("kind",)),
        _object({"kind": _const("documentEnd")}, ("kind",)),
        _object(
            {"kind": _const("before"), "range": CONTENT_RANGE},
            ("kind", "range"),
        ),
        _object(
            {"kind": _const("after"), "range": CONTENT_RANGE},
            ("kind", "range"),
        ),
    ),
}

BODY_SCOPE = {
    "oneOf": (
        _object({"kind": _const("document")}, ("kind",)),
        _object(
            {"kind": _const("range"), "range": CONTENT_RANGE},
            ("kind", "range"),
        ),
    ),
}

POSITIVE_LENGTH = _length_between_points(0.01, 1584)
OFFSET_LENGTH = _length_between_points(-1584, 1584)
MARGIN_LENGTH = _length_between_points(0, 720)

POINT_LENGTH = _object(
    {
        "value": _number(minimum=0, maximum=1584),
        "unit": _const("pt"),
    },
    ("value", "unit"),
)

POSITIVE_POINT_LENGTH = _object(
    {
        "value": _number(minimum=0.01, maximum=1584),
        "unit": _const("pt"),
    },
    ("value", "unit"),
)

POINT_OFFSET = _object(
    {
        "value": _number(minimum=-1584, maximum=1584),
        "unit": _const("pt"),
    },
    ("value", "unit"),
)

DOCUMENT_STATE = _object(
    {
        "persistenceState": _string(
            enum=("unsaved", "saved", "modified")
        ),
        "readOnly": {"type": "boolean"},
    },
    ("persistenceState", "readOnly"),
)

CREATED_DOCUMENT_STATE = _object(
    {
        "persistenceState": _const("unsaved"),
        "readOnly": _const(False),
    },
    ("persistenceState", "readOnly"),
)

SAVED_DOCUMENT_STATE = _object(
    {
        "persistenceState": _const("saved"),
        "readOnly": {"type": "boolean"},
    },
    ("persistenceState", "readOnly"),
)

OPENED_DOCUMENT_STATE = _object(
    {
        "persistenceState": _string(enum=("saved", "modified")),
        "readOnly": {"type": "boolean"},
    },
    ("persistenceState", "readOnly"),
)

DOCX_PATH = _string(
    minimum=1,
    maximum=4096,
    format="absoluteDocxPath",
)

PDF_PATH = _string(
    minimum=1,
    maximum=4096,
    format="absolutePdfPath",
)

IMAGE_PATH = _string(
    minimum=1,
    maximum=4096,
    format="absoluteImagePath",
)

DOCX_ARTIFACT = _object(
    {
        "path": DOCX_PATH,
        "format": _const("docx"),
        "sizeBytes": _integer(minimum=1),
    },
    ("path", "format", "sizeBytes"),
)

PDF_ARTIFACT = _object(
    {
        "path": PDF_PATH,
        "format": _const("pdf"),
        "sizeBytes": _integer(minimum=1),
    },
    ("path", "format", "sizeBytes"),
)

TEXT_FORMAT_PATCH = _object(
    {
        "fontFamily": _string(minimum=1, maximum=128),
        "westernFontFamily": _string(minimum=1, maximum=128),
        "eastAsiaFontFamily": _string(minimum=1, maximum=128),
        "fontSizePt": _number(minimum=1, maximum=300),
        "bold": {"type": "boolean"},
        "italic": {"type": "boolean"},
        "underline": _string(enum=("none", "single")),
        "color": _string(pattern=r"^#[0-9A-Fa-f]{6}$"),
    },
    (),
    **{
        "x-atLeastOne": (
            "fontFamily",
            "westernFontFamily",
            "eastAsiaFontFamily",
            "fontSizePt",
            "bold",
            "italic",
            "underline",
            "color",
        )
    },
)

LINE_SPACING = {
    "oneOf": (
        _object({"kind": _const("single")}, ("kind",)),
        _object({"kind": _const("oneAndHalf")}, ("kind",)),
        _object({"kind": _const("double")}, ("kind",)),
        _object(
            {
                "kind": _const("exact"),
                "points": _number(minimum=0.01, maximum=1584),
            },
            ("kind", "points"),
        ),
        _object(
            {
                "kind": _const("atLeast"),
                "points": _number(minimum=0.01, maximum=1584),
            },
            ("kind", "points"),
        ),
        _object(
            {
                "kind": _const("multiple"),
                "value": _number(minimum=0.01, maximum=100),
            },
            ("kind", "value"),
        ),
    ),
}

PARAGRAPH_FORMAT_PATCH = _object(
    {
        "alignment": _string(
            enum=("left", "center", "right", "justify")
        ),
        "lineSpacing": LINE_SPACING,
        "spaceBeforePt": _number(minimum=0, maximum=1584),
        "spaceAfterPt": _number(minimum=0, maximum=1584),
        "leftIndentPt": _number(minimum=0, maximum=1584),
        "rightIndentPt": _number(minimum=0, maximum=1584),
        "firstLineIndentPt": _number(minimum=-1584, maximum=1584),
    },
    (),
    **{
        "x-atLeastOne": (
            "alignment",
            "lineSpacing",
            "spaceBeforePt",
            "spaceAfterPt",
            "leftIndentPt",
            "rightIndentPt",
            "firstLineIndentPt",
        )
    },
)

TEXT_RUN = _object(
    {
        "text": _string(
            minimum=1,
            maximum=32768,
            format="wordRunText",
            **{"x-maxUtf16Length": 32768},
        ),
        "format": TEXT_FORMAT_PATCH,
    },
    ("text",),
)

TEXT_RUNS = _array(
    TEXT_RUN,
    minimum=1,
    maximum=128,
    **{"x-maxUtf16Text": 262144},
)

INLINE_TEXT_BLOCK = _object(
    {
        "kind": _const("text"),
        "runs": TEXT_RUNS,
    },
    ("kind", "runs"),
)

PARAGRAPH_BLOCK = _object(
    {
        "kind": _const("paragraph"),
        "runs": _array(
            TEXT_RUN,
            minimum=0,
            maximum=128,
            **{"x-maxUtf16Text": 262144},
        ),
        "format": PARAGRAPH_FORMAT_PATCH,
    },
    ("kind", "runs"),
)

HEADING_BLOCK = _object(
    {
        "kind": _const("heading"),
        "level": _integer(minimum=1, maximum=9),
        "runs": TEXT_RUNS,
        "format": PARAGRAPH_FORMAT_PATCH,
    },
    ("kind", "level", "runs"),
)

STRUCTURED_BLOCKS = {
    "oneOf": (
        _array(
            INLINE_TEXT_BLOCK,
            minimum=1,
            maximum=1,
            **{"x-maxUtf16Text": 262144},
        ),
        _array(
            {"oneOf": (PARAGRAPH_BLOCK, HEADING_BLOCK)},
            minimum=1,
            maximum=128,
            **{"x-maxUtf16Text": 262144},
        ),
    ),
}

TEXT_QUERY = _object(
    {
        "scope": BODY_SCOPE,
        "text": _string(
            minimum=1,
            maximum=4096,
            format="wordQueryText",
        ),
        "caseSensitive": {"type": "boolean"},
        "wholeWord": {"type": "boolean"},
    },
    ("scope", "text", "caseSensitive", "wholeWord"),
)

EFFECTIVE_TEXT_FORMAT = _object(
    {
        "fontFamily": _nullable(_string()),
        "westernFontFamily": _nullable(_string()),
        "eastAsiaFontFamily": _nullable(_string()),
        "fontSizePt": _nullable(_number(minimum=0)),
        "bold": _nullable({"type": "boolean"}),
        "italic": _nullable({"type": "boolean"}),
        "underline": _nullable(
            _string(enum=("none", "single", "other"))
        ),
        "color": _nullable(
            {
                "oneOf": (
                    _string(pattern=r"^#[0-9A-F]{6}$"),
                    _const("automatic"),
                )
            }
        ),
    },
    (
        "fontFamily",
        "westernFontFamily",
        "eastAsiaFontFamily",
        "fontSizePt",
        "bold",
        "italic",
        "underline",
        "color",
    ),
)

OBSERVED_RUN = _object(
    {
        "range": CONTENT_RANGE,
        "text": _string(maximum=32768, format="wordNormalizedText"),
        "format": EFFECTIVE_TEXT_FORMAT,
    },
    ("range", "text", "format"),
)

OBSERVED_LINE_SPACING = _nullable({
    "oneOf": LINE_SPACING["oneOf"] + (
        _object({"kind": _const("other")}, ("kind",)),
    )
})

EFFECTIVE_PARAGRAPH_FORMAT = _object(
    {
        "alignment": _nullable(
            _string(enum=("left", "center", "right", "justify", "other"))
        ),
        "lineSpacing": OBSERVED_LINE_SPACING,
        "spaceBeforePt": _nullable(_number()),
        "spaceAfterPt": _nullable(_number()),
        "leftIndentPt": _nullable(_number()),
        "rightIndentPt": _nullable(_number()),
        "firstLineIndentPt": _nullable(_number()),
    },
    (
        "alignment",
        "lineSpacing",
        "spaceBeforePt",
        "spaceAfterPt",
        "leftIndentPt",
        "rightIndentPt",
        "firstLineIndentPt",
    ),
)

PARAGRAPH_SNAPSHOT_BASE = {
    "range": CONTENT_RANGE,
    "complete": {"type": "boolean"},
    "text": _string(maximum=65536, format="wordNormalizedText"),
    "runs": _array(OBSERVED_RUN, maximum=2048),
    "format": EFFECTIVE_PARAGRAPH_FORMAT,
}

PARAGRAPH_SNAPSHOT = {
    "oneOf": (
        _object(
            {
                **PARAGRAPH_SNAPSHOT_BASE,
                "kind": _const("paragraph"),
            },
            (
                "kind",
                "range",
                "complete",
                "text",
                "runs",
                "format",
            ),
        ),
        _object(
            {
                **PARAGRAPH_SNAPSHOT_BASE,
                "kind": _const("heading"),
                "level": _integer(minimum=1, maximum=9),
            },
            (
                "kind",
                "level",
                "range",
                "complete",
                "text",
                "runs",
                "format",
            ),
        ),
    )
}

SECTION_SELECTOR = {
    "oneOf": (
        _object(
            {
                "kind": _const("all"),
                "revision": CONTENT_REVISION,
            },
            ("kind", "revision"),
        ),
        _object(
            {
                "kind": _const("indexes"),
                "indexes": _array(
                    _integer(minimum=0),
                    minimum=1,
                    maximum=64,
                    uniqueItems=True,
                    **{"x-strictlyIncreasing": True},
                ),
                "revision": CONTENT_REVISION,
            },
            ("kind", "indexes", "revision"),
        ),
    )
}

MARGINS_INPUT = _object(
    {
        "top": MARGIN_LENGTH,
        "right": MARGIN_LENGTH,
        "bottom": MARGIN_LENGTH,
        "left": MARGIN_LENGTH,
    },
    ("top", "right", "bottom", "left"),
)

MARGINS_SNAPSHOT = _object(
    {
        "top": POINT_LENGTH,
        "right": POINT_LENGTH,
        "bottom": POINT_LENGTH,
        "left": POINT_LENGTH,
    },
    ("top", "right", "bottom", "left"),
)

LAYOUT_SNAPSHOT = _object(
    {
        "orientation": _string(enum=("portrait", "landscape")),
        "margins": MARGINS_SNAPSHOT,
    },
    ("orientation", "margins"),
)

HEADER_FOOTER_STORY = _object(
    {
        "area": _string(enum=("header", "footer")),
        "variant": _string(
            enum=("primary", "firstPage", "evenPages")
        ),
        "exists": {"type": "boolean"},
        "linkToPrevious": {"type": "boolean"},
        "text": _string(
            maximum=32768,
            format="wordStoryText",
            **{"x-maxUtf16Length": 32768},
        ),
    },
    ("area", "variant", "exists", "linkToPrevious", "text"),
)

HEADER_FOOTER_SNAPSHOT = _object(
    {
        "firstPageEnabled": {"type": "boolean"},
        "evenPagesEnabled": {"type": "boolean"},
        "stories": _array(
            HEADER_FOOTER_STORY,
            minimum=6,
            maximum=6,
            **{"x-maxUtf16Text": 196608},
        ),
    },
    ("firstPageEnabled", "evenPagesEnabled", "stories"),
)

SECTION_SNAPSHOT = _object(
    {
        "index": _integer(minimum=0),
        "layout": LAYOUT_SNAPSHOT,
        "headerFooter": HEADER_FOOTER_SNAPSHOT,
    },
    ("index", "layout", "headerFooter"),
)

STRUCTURE_SNAPSHOT = _object(
    {
        "paragraphCount": _integer(minimum=0),
        "headingCount": _integer(minimum=0),
        "tableCount": _integer(minimum=0),
        "inlineImageCount": _integer(minimum=0),
        "floatingImageCount": _integer(minimum=0),
        "sectionCount": _integer(minimum=1, maximum=64),
        "pageBreakCount": _integer(minimum=0),
        "sectionBreakCount": _integer(minimum=0),
        "sections": _array(SECTION_SNAPSHOT, minimum=1, maximum=64),
    },
    (
        "paragraphCount",
        "headingCount",
        "tableCount",
        "inlineImageCount",
        "floatingImageCount",
        "sectionCount",
        "pageBreakCount",
        "sectionBreakCount",
        "sections",
    ),
)

TABLE_DATA = _array(
    _array(
        _string(
            maximum=32768,
            format="wordStoryText",
            **{"x-maxUtf16Length": 32768},
        ),
        minimum=1,
        maximum=50,
    ),
    minimum=1,
    maximum=100,
    **{
        "x-rectangular": True,
        "x-maxCells": 2000,
        "x-maxUtf16Length": 1048576,
    },
)

ALTERNATIVE_TEXT = {
    "oneOf": (
        _object({"kind": _const("decorative")}, ("kind",)),
        _object(
            {
                "kind": _const("description"),
                "text": _string(
                    minimum=1,
                    maximum=2048,
                    format="wordStoryText",
                    **{"x-maxUtf16Length": 2048},
                ),
            },
            ("kind", "text"),
        ),
    )
}

IMAGE_PLACEMENT = {
    "oneOf": (
        _object({"kind": _const("inline")}, ("kind",)),
        _object(
            {
                "kind": _const("floating"),
                "wrap": _string(
                    enum=(
                        "square",
                        "topBottom",
                        "behindText",
                        "inFrontOfText",
                    )
                ),
                "horizontal": _object(
                    {
                        "relativeTo": _string(
                            enum=("page", "margin", "column")
                        ),
                        "offset": OFFSET_LENGTH,
                    },
                    ("relativeTo", "offset"),
                ),
                "vertical": _object(
                    {
                        "relativeTo": _string(
                            enum=("page", "margin", "paragraph")
                        ),
                        "offset": OFFSET_LENGTH,
                    },
                    ("relativeTo", "offset"),
                ),
            },
            ("kind", "wrap", "horizontal", "vertical"),
        ),
    )
}

IMAGE_SIZE = {
    "oneOf": (
        _object({"kind": _const("intrinsic")}, ("kind",)),
        _object(
            {"kind": _const("width"), "width": POSITIVE_LENGTH},
            ("kind", "width"),
        ),
        _object(
            {"kind": _const("height"), "height": POSITIVE_LENGTH},
            ("kind", "height"),
        ),
        _object(
            {
                "kind": _const("box"),
                "width": POSITIVE_LENGTH,
                "height": POSITIVE_LENGTH,
                "fit": _string(enum=("contain", "stretch")),
            },
            ("kind", "width", "height", "fit"),
        ),
    )
}

OBSERVED_IMAGE_SOURCE = _object(
    {
        "mediaType": _string(enum=("image/png", "image/jpeg")),
        "byteLength": _integer(minimum=1, maximum=26214400),
        "sha256": _string(pattern=r"^[0-9a-f]{64}$"),
    },
    ("mediaType", "byteLength", "sha256"),
)

OBSERVED_IMAGE_SIZE = _object(
    {
        "width": POSITIVE_POINT_LENGTH,
        "height": POSITIVE_POINT_LENGTH,
    },
    ("width", "height"),
)

OBSERVED_IMAGE = {
    "oneOf": (
        _object(
            {
                "kind": _const("inline"),
                "range": NONEMPTY_CONTENT_RANGE,
                "embedded": _const(True),
                "source": OBSERVED_IMAGE_SOURCE,
                "size": OBSERVED_IMAGE_SIZE,
                "alternativeText": ALTERNATIVE_TEXT,
            },
            (
                "kind",
                "range",
                "embedded",
                "source",
                "size",
                "alternativeText",
            ),
        ),
        _object(
            {
                "kind": _const("floating"),
                "anchorRange": CONTENT_RANGE,
                "embedded": _const(True),
                "source": OBSERVED_IMAGE_SOURCE,
                "size": OBSERVED_IMAGE_SIZE,
                "wrap": _string(
                    enum=(
                        "square",
                        "topBottom",
                        "behindText",
                        "inFrontOfText",
                    )
                ),
                "horizontal": _object(
                    {
                        "relativeTo": _string(
                            enum=("page", "margin", "column")
                        ),
                        "offset": POINT_OFFSET,
                    },
                    ("relativeTo", "offset"),
                ),
                "vertical": _object(
                    {
                        "relativeTo": _string(
                            enum=("page", "margin", "paragraph")
                        ),
                        "offset": POINT_OFFSET,
                    },
                    ("relativeTo", "offset"),
                ),
                "alternativeText": ALTERNATIVE_TEXT,
            },
            (
                "kind",
                "anchorRange",
                "embedded",
                "source",
                "size",
                "wrap",
                "horizontal",
                "vertical",
                "alternativeText",
            ),
        ),
    )
}

HEADER_FOOTER_OPERATION = {
    "oneOf": (
        _object(
            {
                "kind": _const("replace"),
                "text": _string(
                    minimum=1,
                    maximum=32768,
                    format="wordStoryText",
                    **{"x-maxUtf16Length": 32768},
                ),
            },
            ("kind", "text"),
        ),
        _object({"kind": _const("clear")}, ("kind",)),
        _object({"kind": _const("linkToPrevious")}, ("kind",)),
    )
}

HEADER_FOOTER_UPDATE = _object(
    {
        "area": _string(enum=("header", "footer")),
        "variant": _string(
            enum=("primary", "firstPage", "evenPages")
        ),
        "operation": HEADER_FOOTER_OPERATION,
    },
    ("area", "variant", "operation"),
)

OBSERVED_HEADER_FOOTER_STORY = _object(
    {
        "sectionIndex": _integer(minimum=0),
        "area": _string(enum=("header", "footer")),
        "variant": _string(
            enum=("primary", "firstPage", "evenPages")
        ),
        "variantEnabled": {"type": "boolean"},
        "exists": {"type": "boolean"},
        "linkToPrevious": {"type": "boolean"},
        "text": _string(
            maximum=32768,
            format="wordStoryText",
            **{"x-maxUtf16Length": 32768},
        ),
    },
    (
        "sectionIndex",
        "area",
        "variant",
        "variantEnabled",
        "exists",
        "linkToPrevious",
        "text",
    ),
)

OVERWRITE_POLICY = _string(
    enum=("failIfExists", "renameIfExists", "replaceExisting")
)

BINDING_ERRORS = (
    "DOCUMENT_CLOSED",
    "DOCUMENT_BINDING_UNAVAILABLE",
)

MUTATION_ERRORS = BINDING_ERRORS + (
    "DOCUMENT_READ_ONLY",
)

CONTENT_READ_ERRORS = BINDING_ERRORS + (
    "STALE_CONTENT_RANGE",
    "CONTENT_RANGE_OUT_OF_BOUNDS",
    "CONTENT_LIMIT_EXCEEDED",
    "CONTENT_CHANGED_DURING_ACTION",
)

CONTENT_MUTATION_ERRORS = CONTENT_READ_ERRORS + (
    "DOCUMENT_READ_ONLY",
    "CONTENT_PROTECTED",
)

COMMON_CONTRACT_ERRORS = (
    "INVALID_PARAMS",
    "INVALID_RESULT",
    "RESPONSE_LOST",
    "WORD_CAPABILITY_UNAVAILABLE",
)
