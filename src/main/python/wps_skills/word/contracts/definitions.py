"""Authoritative Word Action contract definitions."""

from wps_skills.core.action_runtime import ActionContract, ApplicationContractSet
from wps_skills.word.persistence.output_names import MAX_OUTPUT_ATTEMPTS
from wps_skills.word.contracts.formats import (
    WORD_FORMAT_VALIDATORS,
)
from wps_skills.word.contracts.schemas import (
    ALTERNATIVE_TEXT,
    BINDING_ERRORS,
    BODY_ANCHOR,
    BODY_SCOPE,
    COMMON_CONTRACT_ERRORS,
    CONTENT_MUTATION_ERRORS,
    CONTENT_RANGE,
    CONTENT_READ_ERRORS,
    CONTENT_REVISION,
    CREATED_DOCUMENT_STATE,
    DOCUMENT_STATE,
    DOCX_ARTIFACT,
    DOCX_PATH,
    HEADER_FOOTER_UPDATE,
    IMAGE_PATH,
    IMAGE_PLACEMENT,
    IMAGE_SIZE,
    LAYOUT_SNAPSHOT,
    MARGINS_INPUT,
    MUTATION_ERRORS,
    NONEMPTY_CONTENT_RANGE,
    OBSERVED_HEADER_FOOTER_STORY,
    OBSERVED_IMAGE,
    OPENED_DOCUMENT_STATE,
    OVERWRITE_POLICY,
    PARAGRAPH_SNAPSHOT,
    PDF_ARTIFACT,
    PDF_PATH,
    SAVED_DOCUMENT_STATE,
    SECTION_SELECTOR,
    STRUCTURED_BLOCKS,
    STRUCTURE_SNAPSHOT,
    TABLE_DATA,
    TEXT_QUERY,
    TEXT_RUNS,
    _array,
    _const,
    _integer,
    _nullable,
    _object,
    _string,
)
from wps_skills.word.contracts.examples import (
    _WORD_CONSTRAINTS,
    _WORD_EXAMPLE_PARAMS,
)
from wps_skills.word.contracts.validation import (
    _semantic_params_error,
    _semantic_result_error,
)


def _contract(
    *,
    name,
    category,
    purpose,
    binding_role,
    risk,
    parameters,
    result,
    stable_errors,
    verification,
):
    role_errors = (
        ("TASK_DOCUMENT_ALREADY_BOUND",)
        if binding_role == "establish"
        else ("TASK_DOCUMENT_NOT_BOUND",)
        if binding_role == "required"
        else ()
    )
    errors = tuple(dict.fromkeys(
        COMMON_CONTRACT_ERRORS + role_errors + tuple(stable_errors)
    ))
    prerequisites = (
        "Document Intent is resolved for this Word Action.",
        "The Word Task is UNBOUND.",
    ) if binding_role == "establish" else (
        "The Word Task has one exact live Word Document Binding.",
        "The bound document and controller remain provably usable.",
    )
    return ActionContract(
        name=name,
        category=category,
        purpose=purpose,
        binding_role=binding_role,
        risk=risk,
        parameters=parameters,
        result=result,
        stable_errors=errors,
        verification=verification,
        prerequisites=prerequisites,
        constraints=_WORD_CONSTRAINTS[name],
        examples=({"params": _WORD_EXAMPLE_PARAMS[name]},),
        parameter_validator=(
            lambda params, action=name: _semantic_params_error(action, params)
        ),
        semantic_validator=(
            lambda params, result, action=name: _semantic_result_error(
                action,
                params,
                result,
            )
        ),
    )


WORD_TARGET_CONTRACT_SET = ApplicationContractSet(
    application="word",
    format_validators=WORD_FORMAT_VALIDATORS,
    contracts=(
        _contract(
            name="createDocument",
            category="documentEstablishment",
            purpose=(
                "Create one blank live Word document and establish the "
                "Document Binding."
            ),
            binding_role="establish",
            risk="write",
            parameters=_object({}, ()),
            result=_object(
                {"revision": CONTENT_REVISION, "documentState": CREATED_DOCUMENT_STATE},
                ("revision", "documentState"),
            ),
            stable_errors=(
                "DOCUMENT_CREATE_FAILED",
                "DOCUMENT_BINDING_UNAVAILABLE",
                "RESPONSE_LOST",
            ),
            verification=(
                "Prove the exact object returned by create is live, unsaved, "
                "Lease-bound, and committed before success becomes observable."
            ),
        ),
        _contract(
            name="openDocument",
            category="documentEstablishment",
            purpose=(
                "Open or reuse one exact user-identified existing .docx and "
                "establish the Document Binding."
            ),
            binding_role="establish",
            risk="write",
            parameters=_object({"path": DOCX_PATH}, ("path",)),
            result=_object(
                {
                    "revision": CONTENT_REVISION,
                    "artifact": DOCX_ARTIFACT,
                    "documentState": OPENED_DOCUMENT_STATE,
                },
                ("revision", "artifact", "documentState"),
            ),
            stable_errors=(
                "DOCUMENT_NOT_FOUND",
                "DOCUMENT_ACCESS_DENIED",
                "DOCUMENT_PASSWORD_REQUIRED",
                "DOCUMENT_FORMAT_UNSUPPORTED",
                "DOCUMENT_OPEN_FAILED",
                "DOCUMENT_LEASE_CONFLICT",
                "DOCUMENT_QUARANTINED",
                "DOCUMENT_BINDING_UNAVAILABLE",
                "RESPONSE_LOST",
            ),
            verification=(
                "Prove stable file identity, exact open-or-reused object, "
                "observed .docx artifact/state, and atomic Binding plus Lease. "
                "Fingerprint observation must not change the document persistence state."
            ),
        ),
        _contract(
            name="writeContent",
            category="bodyContent",
            purpose="Insert structured body content at an explicit body anchor.",
            binding_role="required",
            risk="write",
            parameters=_object(
                {"anchor": BODY_ANCHOR, "blocks": STRUCTURED_BLOCKS},
                ("anchor", "blocks"),
            ),
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "range": NONEMPTY_CONTENT_RANGE,
                },
                ("revisionBefore", "revisionAfter", "range"),
            ),
            stable_errors=CONTENT_MUTATION_ERRORS + (
                "CONTENT_ANCHOR_NOT_PARAGRAPH_BOUNDARY",
                "CONTENT_FORMAT_UNSUPPORTED",
                "CONTENT_WRITE_FAILED",
                "CONTENT_VERIFICATION_FAILED",
            ),
            verification=(
                "Read back inserted text, block structure, semantic headings, "
                "and every explicitly supplied format field at the fresh range."
            ),
        ),
        _contract(
            name="inspectDocument",
            category="bodyContent",
            purpose=(
                "Return one revision-coherent normalized snapshot of bound "
                "Word body content, structure, layout, and document state."
            ),
            binding_role="required",
            risk="read",
            parameters=_object(
                {
                    "scope": BODY_SCOPE,
                    "limits": _object(
                        {
                            "maxTextCharacters": _integer(
                                minimum=1, maximum=65536
                            ),
                            "maxParagraphs": _integer(
                                minimum=1, maximum=512
                            ),
                            "maxRuns": _integer(
                                minimum=1, maximum=2048
                            ),
                        },
                        (
                            "maxTextCharacters",
                            "maxParagraphs",
                            "maxRuns",
                        ),
                    ),
                },
                ("scope", "limits"),
            ),
            result=_object(
                {
                    "revision": CONTENT_REVISION,
                    "scopeRange": CONTENT_RANGE,
                    "returnedRange": CONTENT_RANGE,
                    "text": _string(
                        maximum=65536,
                        format="wordNormalizedText",
                    ),
                    "paragraphs": _array(
                        PARAGRAPH_SNAPSHOT,
                        maximum=512,
                    ),
                    "structure": STRUCTURE_SNAPSHOT,
                    "documentState": DOCUMENT_STATE,
                    "truncated": {"type": "boolean"},
                    "remainingRange": _nullable(CONTENT_RANGE),
                },
                (
                    "revision",
                    "scopeRange",
                    "returnedRange",
                    "text",
                    "paragraphs",
                    "structure",
                    "documentState",
                    "truncated",
                    "remainingRange",
                ),
            ),
            stable_errors=CONTENT_READ_ERRORS + (
                "CONTENT_READ_FAILED",
                "STALE_DOCUMENT_REVISION",
            ),
            verification=(
                "Take all body, section, layout, header/footer, and persistence "
                "facts from one coherent revision without changing persistence state; "
                "never return mixed snapshots. Live header/footer text comes from "
                "the in-memory document XML, including unsaved changes."
            ),
        ),
        _contract(
            name="findContent",
            category="bodyContent",
            purpose=(
                "Find bounded, literal, non-overlapping body-text matches "
                "against one coherent content revision."
            ),
            binding_role="required",
            risk="read",
            parameters=_object(
                {
                    "query": TEXT_QUERY,
                    "limit": _integer(minimum=1, maximum=200),
                },
                ("query", "limit"),
            ),
            result=_object(
                {
                    "revision": CONTENT_REVISION,
                    "scopeRange": CONTENT_RANGE,
                    "matches": _array(
                        _object(
                            {
                                "range": NONEMPTY_CONTENT_RANGE,
                                "text": _string(
                                    minimum=1,
                                    maximum=4096,
                                    format="wordQueryText",
                                ),
                            },
                            ("range", "text"),
                        ),
                        maximum=200,
                    ),
                    "truncated": {"type": "boolean"},
                    "remainingRange": _nullable(CONTENT_RANGE),
                },
                (
                    "revision",
                    "scopeRange",
                    "matches",
                    "truncated",
                    "remainingRange",
                ),
            ),
            stable_errors=CONTENT_READ_ERRORS + ("CONTENT_READ_FAILED",),
            verification=(
                "Return document-order non-overlapping observed literals and "
                "ranges from one revision; an empty match list is success."
            ),
        ),
        _contract(
            name="replaceContent",
            category="bodyContent",
            purpose=(
                "Replace one exact range or an explicitly counted literal "
                "query result with preflighted structured content."
            ),
            binding_role="required",
            risk="destructive",
            parameters={
                "oneOf": (
                    _object(
                        {
                            "target": _object(
                                {
                                    "kind": _const("range"),
                                    "range": NONEMPTY_CONTENT_RANGE,
                                },
                                ("kind", "range"),
                            ),
                            "replacement": {
                                "oneOf": (
                                    _object(
                                        {"kind": _const("delete")},
                                        ("kind",),
                                    ),
                                    _object(
                                        {
                                            "kind": _const("blocks"),
                                            "blocks": STRUCTURED_BLOCKS,
                                        },
                                        ("kind", "blocks"),
                                    ),
                                )
                            },
                        },
                        ("target", "replacement"),
                    ),
                    _object(
                        {
                            "target": _object(
                                {
                                    "kind": _const("query"),
                                    "query": TEXT_QUERY,
                                    "expectedMatchCount": _integer(
                                        minimum=1, maximum=200
                                    ),
                                },
                                (
                                    "kind",
                                    "query",
                                    "expectedMatchCount",
                                ),
                            ),
                            "replacement": {
                                "oneOf": (
                                    _object(
                                        {"kind": _const("delete")},
                                        ("kind",),
                                    ),
                                    _object(
                                        {
                                            "kind": _const("text"),
                                            "runs": TEXT_RUNS,
                                        },
                                        ("kind", "runs"),
                                    ),
                                )
                            },
                        },
                        ("target", "replacement"),
                    ),
                )
            },
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "matchedCount": _integer(minimum=1, maximum=200),
                    "ranges": _array(
                        CONTENT_RANGE,
                        minimum=1,
                        maximum=200,
                    ),
                },
                (
                    "revisionBefore",
                    "revisionAfter",
                    "matchedCount",
                    "ranges",
                ),
            ),
            stable_errors=CONTENT_MUTATION_ERRORS + (
                "MATCH_COUNT_MISMATCH",
                "CONTENT_RANGE_NOT_REPLACEABLE",
                "CONTENT_FORMAT_UNSUPPORTED",
                "CONTENT_WRITE_FAILED",
                "CONTENT_VERIFICATION_FAILED",
            ),
            verification=(
                "Preflight all targets and counts before the first write, "
                "mutate from the end, then read back every final fresh range."
            ),
        ),
        _contract(
            name="insertTable",
            category="embeddedContent",
            purpose=(
                "Insert one non-empty rectangular plain-text table at an "
                "explicit body anchor."
            ),
            binding_role="required",
            risk="write",
            parameters=_object(
                {
                    "anchor": BODY_ANCHOR,
                    "data": TABLE_DATA,
                    "headerRow": {"type": "boolean"},
                },
                ("anchor", "data", "headerRow"),
            ),
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "table": _object(
                        {
                            "range": NONEMPTY_CONTENT_RANGE,
                            "rowCount": _integer(minimum=1, maximum=100),
                            "columnCount": _integer(
                                minimum=1, maximum=50
                            ),
                            "headerRow": {"type": "boolean"},
                            "data": TABLE_DATA,
                        },
                        (
                            "range",
                            "rowCount",
                            "columnCount",
                            "headerRow",
                            "data",
                        ),
                    ),
                },
                ("revisionBefore", "revisionAfter", "table"),
            ),
            stable_errors=CONTENT_MUTATION_ERRORS + (
                "TABLE_ANCHOR_UNSUPPORTED",
                "TABLE_APPLY_FAILED",
                "TABLE_VERIFICATION_FAILED",
                "WORD_CAPABILITY_UNAVAILABLE",
            ),
            verification=(
                "Read back exact dimensions, repeating-header state, range, "
                "and normalized text from every created cell."
            ),
        ),
        _contract(
            name="insertImage",
            category="embeddedContent",
            purpose=(
                "Validate and embed one PNG or JPEG at an explicit body anchor "
                "with explicit placement, size, and alternative text."
            ),
            binding_role="required",
            risk="write",
            parameters=_object(
                {
                    "anchor": BODY_ANCHOR,
                    "source": _object(
                        {
                            "kind": _const("file"),
                            "path": IMAGE_PATH,
                        },
                        ("kind", "path"),
                    ),
                    "placement": IMAGE_PLACEMENT,
                    "size": IMAGE_SIZE,
                    "alternativeText": ALTERNATIVE_TEXT,
                },
                (
                    "anchor",
                    "source",
                    "placement",
                    "size",
                    "alternativeText",
                ),
            ),
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "image": OBSERVED_IMAGE,
                },
                ("revisionBefore", "revisionAfter", "image"),
            ),
            stable_errors=CONTENT_MUTATION_ERRORS + (
                "IMAGE_SOURCE_NOT_FOUND",
                "IMAGE_SOURCE_UNREADABLE",
                "IMAGE_FORMAT_UNSUPPORTED",
                "IMAGE_SOURCE_LIMIT_EXCEEDED",
                "IMAGE_PLACEMENT_UNSUPPORTED",
                "IMAGE_APPLY_FAILED",
                "IMAGE_VERIFICATION_FAILED",
                "WORD_CAPABILITY_UNAVAILABLE",
            ),
            verification=(
                "Read back embedded media hash/type, dimensions, exact anchor, "
                "placement, wrap, and alternative text before success."
            ),
        ),
        _contract(
            name="setHeaderFooter",
            category="pageStructure",
            purpose=(
                "Apply explicit header/footer operations to revision-bound "
                "sections and read back every touched story."
            ),
            binding_role="required",
            risk="write",
            parameters=_object(
                {
                    "sections": SECTION_SELECTOR,
                    "updates": _array(
                        HEADER_FOOTER_UPDATE,
                        minimum=1,
                        maximum=6,
                        **{"x-maxUtf16Text": 1048576},
                    ),
                },
                ("sections", "updates"),
                **{"x-uniqueBy": ("updates", "area", "variant")},
            ),
            result=_object(
                {
                    "selectedSectionCount": _integer(
                        minimum=1, maximum=64
                    ),
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "stories": _array(
                        OBSERVED_HEADER_FOOTER_STORY,
                        minimum=1,
                        maximum=384,
                        **{"x-maxUtf16Text": 1048576},
                    ),
                },
                (
                    "selectedSectionCount",
                    "revisionBefore",
                    "revisionAfter",
                    "stories",
                ),
            ),
            stable_errors=MUTATION_ERRORS + (
                "STALE_DOCUMENT_REVISION",
                "SECTION_NOT_FOUND",
                "HEADER_FOOTER_LINK_INVALID",
                "HEADER_FOOTER_APPLY_FAILED",
                "HEADER_FOOTER_VERIFICATION_FAILED",
                "CONTENT_LIMIT_EXCEEDED",
                "CONTENT_CHANGED_DURING_ACTION",
                "WORD_CAPABILITY_UNAVAILABLE",
            ),
            verification=(
                "Preflight every section/update, then read back enabled flags, "
                "link state, existence, and normalized text in stable order."
            ),
        ),
        _contract(
            name="setPageLayout",
            category="pageStructure",
            purpose=(
                "Apply explicit orientation and complete margin changes to "
                "revision-bound sections."
            ),
            binding_role="required",
            risk="write",
            parameters=_object(
                {
                    "sections": SECTION_SELECTOR,
                    "layout": _object(
                        {
                            "orientation": _string(
                                enum=("portrait", "landscape")
                            ),
                            "margins": MARGINS_INPUT,
                        },
                        (),
                        **{"x-atLeastOne": ("orientation", "margins")},
                    ),
                },
                ("sections", "layout"),
            ),
            result=_object(
                {
                    "selectedSectionCount": _integer(
                        minimum=1, maximum=64
                    ),
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "sections": _array(
                        _object(
                            {
                                "index": _integer(minimum=0),
                                "layout": LAYOUT_SNAPSHOT,
                            },
                            ("index", "layout"),
                        ),
                        minimum=1,
                        maximum=64,
                    ),
                },
                (
                    "selectedSectionCount",
                    "revisionBefore",
                    "revisionAfter",
                    "sections",
                ),
            ),
            stable_errors=MUTATION_ERRORS + (
                "STALE_DOCUMENT_REVISION",
                "SECTION_NOT_FOUND",
                "PAGE_LAYOUT_INVALID",
                "PAGE_LAYOUT_APPLY_FAILED",
                "PAGE_LAYOUT_VERIFICATION_FAILED",
                "CONTENT_CHANGED_DURING_ACTION",
                "CONTENT_LIMIT_EXCEEDED",
                "WORD_CAPABILITY_UNAVAILABLE",
            ),
            verification=(
                "Preflight all section dimensions, then read back normalized "
                "orientation and all four margins for every selected section."
            ),
        ),
        _contract(
            name="insertBreak",
            category="pageStructure",
            purpose=(
                "Insert one explicit page or section break at a collapsed "
                "body anchor and verify the resulting structure."
            ),
            binding_role="required",
            risk="write",
            parameters=_object(
                {
                    "anchor": BODY_ANCHOR,
                    "type": _string(
                        enum=(
                            "page",
                            "sectionNextPage",
                            "sectionContinuous",
                            "sectionEvenPage",
                            "sectionOddPage",
                        )
                    ),
                },
                ("anchor", "type"),
            ),
            result={
                "oneOf": (
                    _object(
                        {
                            "revisionBefore": CONTENT_REVISION,
                            "revisionAfter": CONTENT_REVISION,
                            "break": _object(
                                {
                                    "type": _const("page"),
                                    "range": NONEMPTY_CONTENT_RANGE,
                                    "sectionCountBefore": _integer(
                                        minimum=1
                                    ),
                                    "sectionCountAfter": _integer(
                                        minimum=1
                                    ),
                                },
                                (
                                    "type",
                                    "range",
                                    "sectionCountBefore",
                                    "sectionCountAfter",
                                ),
                            ),
                        },
                        ("revisionBefore", "revisionAfter", "break"),
                    ),
                    _object(
                        {
                            "revisionBefore": CONTENT_REVISION,
                            "revisionAfter": CONTENT_REVISION,
                            "break": _object(
                                {
                                    "type": _string(
                                        enum=(
                                            "sectionNextPage",
                                            "sectionContinuous",
                                            "sectionEvenPage",
                                            "sectionOddPage",
                                        )
                                    ),
                                    "range": NONEMPTY_CONTENT_RANGE,
                                    "sectionCountBefore": _integer(
                                        minimum=1
                                    ),
                                    "sectionCountAfter": _integer(
                                        minimum=1
                                    ),
                                    "followingSectionIndex": _integer(
                                        minimum=1
                                    ),
                                },
                                (
                                    "type",
                                    "range",
                                    "sectionCountBefore",
                                    "sectionCountAfter",
                                    "followingSectionIndex",
                                ),
                            ),
                        },
                        ("revisionBefore", "revisionAfter", "break"),
                    ),
                )
            },
            stable_errors=CONTENT_MUTATION_ERRORS + (
                "BREAK_ANCHOR_UNSUPPORTED",
                "BREAK_APPLY_FAILED",
                "BREAK_VERIFICATION_FAILED",
                "WORD_CAPABILITY_UNAVAILABLE",
            ),
            verification=(
                "Use a duplicated collapsed range, then read back the page "
                "marker or exact section count/start type at a fresh revision."
            ),
        ),
        _contract(
            name="save",
            category="persistence",
            purpose=(
                "Persist the bound .docx through its existing locator and "
                "verify the saved artifact."
            ),
            binding_role="required",
            risk="write",
            parameters=_object({}, ()),
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "artifact": DOCX_ARTIFACT,
                    "documentState": SAVED_DOCUMENT_STATE,
                },
                (
                    "revisionBefore",
                    "revisionAfter",
                    "artifact",
                    "documentState",
                ),
            ),
            stable_errors=MUTATION_ERRORS + (
                "PERSISTENCE_LOCATOR_REQUIRED",
                "OUTPUT_ACCESS_DENIED",
                "OUTPUT_WRITE_FAILED",
                "OUTPUT_VERIFICATION_FAILED",
                "DOCUMENT_CHANGED_DURING_ACTION",
            ),
            verification=(
                "Prove the same backing identity, saved state, .docx format, "
                "ordinary non-empty artifact, and live exact Binding."
            ),
        ),
        _contract(
            name="saveAs",
            category="persistence",
            purpose=(
                "Persist the bound live document as .docx at one explicit "
                "user-authorized output locator."
            ),
            binding_role="required",
            risk="destructive",
            parameters=_object(
                {
                    "outputPath": DOCX_PATH,
                    "overwritePolicy": _string(enum=("failIfExists", "renameIfExists")),
                },
                ("outputPath", "overwritePolicy"),
            ),
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "artifact": DOCX_ARTIFACT,
                    "documentState": SAVED_DOCUMENT_STATE,
                    "replacedExisting": {"type": "boolean"},
                    "outputResolution": _object(
                        {"requestedPath": DOCX_PATH,
                         "attempts": _integer(minimum=1, maximum=MAX_OUTPUT_ATTEMPTS),
                         "renamed": {"type": "boolean"}},
                        ("requestedPath", "attempts", "renamed"),
                    ),
                },
                (
                    "revisionBefore",
                    "revisionAfter",
                    "artifact",
                    "documentState",
                    "replacedExisting",
                ),
            ),
            stable_errors=MUTATION_ERRORS + (
                "OUTPUT_ALREADY_EXISTS",
                "OUTPUT_NAME_EXHAUSTED",
                "OUTPUT_MATCHES_BOUND_DOCUMENT",
                "OUTPUT_PARENT_NOT_FOUND",
                "OUTPUT_PATH_INVALID",
                "OUTPUT_IN_USE",
                "OUTPUT_ACCESS_DENIED",
                "OUTPUT_WRITE_FAILED",
                "OUTPUT_VERIFICATION_FAILED",
                "DOCUMENT_LEASE_CONFLICT",
                "DOCUMENT_QUARANTINED",
                "DOCUMENT_CHANGED_DURING_ACTION",
            ),
            verification=(
                "Prove explicit .docx format, artifact, saved state, same live "
                "document, destination identity, and gap-free Lease migration."
            ),
        ),
        _contract(
            name="exportPdf",
            category="persistence",
            purpose=(
                "Export the complete bound Word document to one explicit "
                "user-authorized PDF artifact."
            ),
            binding_role="required",
            risk="destructive",
            parameters=_object(
                {
                    "outputPath": PDF_PATH,
                    "overwritePolicy": OVERWRITE_POLICY,
                },
                ("outputPath", "overwritePolicy"),
            ),
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "artifact": PDF_ARTIFACT,
                    "documentStateBefore": DOCUMENT_STATE,
                    "documentStateAfter": DOCUMENT_STATE,
                    "replacedExisting": {"type": "boolean"},
                    "outputResolution": _object(
                        {"requestedPath": PDF_PATH,
                         "attempts": _integer(minimum=1, maximum=MAX_OUTPUT_ATTEMPTS),
                         "renamed": {"type": "boolean"}},
                        ("requestedPath", "attempts", "renamed"),
                    ),
                },
                (
                    "revisionBefore",
                    "revisionAfter",
                    "artifact",
                    "documentStateBefore",
                    "documentStateAfter",
                    "replacedExisting",
                ),
            ),
            stable_errors=BINDING_ERRORS + (
                "OUTPUT_ALREADY_EXISTS",
                "OUTPUT_NAME_EXHAUSTED",
                "OUTPUT_PARENT_NOT_FOUND",
                "OUTPUT_PATH_INVALID",
                "OUTPUT_IN_USE",
                "OUTPUT_ACCESS_DENIED",
                "OUTPUT_WRITE_FAILED",
                "OUTPUT_VERIFICATION_FAILED",
                "DOCUMENT_CHANGED_DURING_ACTION",
            ),
            verification=(
                "Prove a non-empty readable PDF for one stable content revision "
                "without saving, retargeting, or changing the Word Binding."
            ),
        ),
    ),
)

WORD_TARGET_ACTION_INDEX = WORD_TARGET_CONTRACT_SET.action_index()
