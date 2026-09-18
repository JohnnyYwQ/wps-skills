"""One executable registration per Word Action: contract and implementation."""

from dataclasses import dataclass
from types import MappingProxyType
from typing import Callable, Optional

from wps_skills.core.action_runtime import ActionContract, ApplicationContractSet
from wps_skills.word.contracts import WORD_TARGET_CONTRACT_SET, WORD_FORMAT_VALIDATORS
from wps_skills.word.actions.handlers import _handler


@dataclass(frozen=True)
class ActionDefinition:
    contract: ActionContract
    handler: Callable
    prepare: Optional[Callable] = None
    observe: Optional[Callable] = None
    operation: Optional[str] = None


def _required(name, operation):
    return ActionDefinition(WORD_TARGET_CONTRACT_SET.resolve(name), _handler(operation), operation=operation)


# Admission is explicit. Discovery, validation, packaged schemas and dispatch
# are all derived from these entries, never from the target portfolio alone.
WORD_ACTIONS = MappingProxyType({
    "createDocument": ActionDefinition(
        contract=WORD_TARGET_CONTRACT_SET.resolve("createDocument"),
        prepare=lambda backend, params, context: backend.prepare_create_document(context),
        handler=lambda backend, preparation, params, context: backend.create_document(preparation, context),
        observe=lambda adapter, acquisition, params: adapter._created(acquisition),
    ),
    "openDocument": ActionDefinition(
        contract=WORD_TARGET_CONTRACT_SET.resolve("openDocument"),
        prepare=lambda backend, params, context: backend.prepare_open_document(params["path"], context),
        handler=lambda backend, preparation, params, context: backend.open_document(preparation, params["path"], context),
        observe=lambda adapter, acquisition, params: adapter._opened(acquisition, params["path"]),
    ),
    "writeContent": _required("writeContent", "insert_structured_body_content"),
    "inspectDocument": _required("inspectDocument", "read_revision_coherent_snapshot"),
    "findContent": _required("findContent", "find_literal_body_content"),
    "replaceContent": _required("replaceContent", "replace_body_content"),
    "insertTable": _required("insertTable", "insert_plain_text_table"),
    "insertImage": _required("insertImage", "insert_embedded_image"),
    "setHeaderFooter": _required("setHeaderFooter", "update_header_footer_stories"),
    "setPageLayout": _required("setPageLayout", "update_page_layout"),
    "insertBreak": _required("insertBreak", "insert_body_break"),
    "save": _required("save", "save_existing_artifact"),
    "saveAs": _required("saveAs", "save_as_artifact"),
    "exportPdf": _required("exportPdf", "export_pdf_artifact"),
})

WORD_HANDLERS = MappingProxyType({
    name: entry.handler for name, entry in WORD_ACTIONS.items()
    if entry.contract.binding_role == "required"
})


WORD_PRODUCTION_CONTRACT_SET = ApplicationContractSet(
    application="word",
    contracts=tuple(entry.contract for entry in WORD_ACTIONS.values()),
    format_validators=WORD_FORMAT_VALIDATORS,
)
WORD_PRODUCTION_ACTION_INDEX = WORD_PRODUCTION_CONTRACT_SET.action_index()


def content_actions():
    """The subset Agents can compose inside a Word Task's content steps."""
    return {
        name: entry for name, entry in WORD_ACTIONS.items()
        if entry.contract.binding_role == "required" and entry.contract.category != "persistence"
    }
