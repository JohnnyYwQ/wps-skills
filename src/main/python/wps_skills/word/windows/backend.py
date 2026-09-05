"""Windows Word Backend over one Session-owned local automation bridge."""

from dataclasses import dataclass
from typing import Any, Mapping, Optional, Protocol, runtime_checkable

from wps_skills.core.action_session import ControllerContext
from wps_skills.windows.bridge_types import BackendActionFailure, WindowsDocument
from wps_skills.word.adapter import (
    WordBackendAcquisition,
    WordBackendOperation,
    WordBackendPreparation,
    WordDefiniteEstablishFailure,
    WordUnprovableEstablishFailure,
)


@runtime_checkable
class WindowsWordBridge(Protocol):
    """Private synchronous seam to the Session-owned PowerShell bridge."""

    def execute(
        self,
        operation: str,
        arguments: Mapping[str, Any],
        context: Optional[ControllerContext],
    ) -> Mapping[str, Any]:
        ...

    def close(self) -> bool:
        ...


@dataclass(frozen=True)
class WindowsWordPreparation:
    preparation_id: str
    coordination_identity: str
    authorized_path: Optional[str]

    def __post_init__(self) -> None:
        if not isinstance(self.preparation_id, str) or not self.preparation_id:
            raise ValueError("bridge preparation id must be non-empty")
        if (
            not isinstance(self.coordination_identity, str)
            or not self.coordination_identity
        ):
            raise ValueError("coordination identity must be non-empty")


def _require_mapping(value, description):
    if not isinstance(value, Mapping):
        raise TypeError(f"{description} must be an object")
    return value


def _require_exact_keys(value, expected, description):
    value = _require_mapping(value, description)
    if set(value) != set(expected):
        raise TypeError(f"{description} has invalid fields")
    return value


def _require_string(value, description):
    if not isinstance(value, str) or not value:
        raise TypeError(f"{description} must be a non-empty string")
    return value


def _require_boolean(value, description):
    if not isinstance(value, bool):
        raise TypeError(f"{description} must be Boolean")
    return value


class WindowsWordBackend:
    """Validate bridge facts and preserve one exact bridge document reference."""

    _CREATE_FIELDS = frozenset({
        "documentId",
        "revision",
        "persistenceState",
        "readOnly",
    })
    _OPEN_FIELDS = _CREATE_FIELDS | frozenset({
        "artifactFormat",
        "artifactSizeBytes",
    })

    def __init__(self, *, bridge: WindowsWordBridge):
        if not isinstance(bridge, WindowsWordBridge):
            raise ValueError("bridge must satisfy the Windows Word bridge seam")
        self._bridge = bridge
        self._preparation = None
        self._document = None
        self._persistence_path = None

    def _execute(self, operation, arguments, context):
        result = self._bridge.execute(operation, arguments, context)
        return _require_mapping(result, "Windows Word bridge result")

    def _begin_acquisition(self):
        if self._document is not None:
            raise RuntimeError("Windows Word Backend is already bound")

    def _prepare(self, operation, arguments, context, *, path=None):
        self._begin_acquisition()
        try:
            result = _require_exact_keys(
                self._execute(operation, arguments, context),
                {"preparationId", "coordinationIdentity"},
                "Windows Word preparation result",
            )
        except BackendActionFailure as failure:
            if (
                failure.outcome == "failed"
                and failure.binding_disposition == "unchanged"
            ):
                self._preparation = None
                raise WordDefiniteEstablishFailure(
                    code=failure.code,
                    message=failure.message,
                ) from failure
            raise WordUnprovableEstablishFailure(
                outcome=failure.outcome,
                code=failure.code,
                message=failure.message,
            ) from failure
        state = WindowsWordPreparation(
            preparation_id=_require_string(
                result["preparationId"],
                "bridge preparation id",
            ),
            coordination_identity=_require_string(
                result["coordinationIdentity"],
                "coordination identity",
            ),
            authorized_path=path,
        )
        preparation = WordBackendPreparation(
            coordination_identity=state.coordination_identity,
            state=state,
        )
        self._preparation = preparation
        return preparation

    def prepare_create_document(self, context):
        return self._prepare(
            "prepare_new_document",
            {},
            context,
        )

    def prepare_open_document(self, path, context):
        if not isinstance(path, str) or not path:
            raise ValueError("open path must be non-empty")
        return self._prepare(
            "prepare_existing_document",
            {"path": path},
            context,
            path=path,
        )

    def _require_preparation(self, preparation, *, path):
        if preparation is not self._preparation:
            raise ValueError(
                "Windows Word Backend received another preparation"
            )
        if not isinstance(preparation, WordBackendPreparation) or not isinstance(
            preparation.state,
            WindowsWordPreparation,
        ):
            raise ValueError("Windows Word Backend preparation is invalid")
        if preparation.state.authorized_path != path:
            raise ValueError("Windows Word preparation path does not match")
        return preparation.state

    @staticmethod
    def _acquisition(result, *, authorized_path, opened):
        fields = (
            WindowsWordBackend._OPEN_FIELDS
            if opened
            else WindowsWordBackend._CREATE_FIELDS
        )
        result = _require_exact_keys(
            result,
            fields,
            "Windows Word acquisition result",
        )
        document = WindowsDocument(
            bridge_document_id=_require_string(
                result["documentId"],
                "bridge document id",
            ),
            authorized_path=authorized_path,
        )
        revision = _require_string(result["revision"], "Content Revision")
        persistence_state = _require_string(
            result["persistenceState"],
            "persistence state",
        )
        read_only = _require_boolean(result["readOnly"], "read-only state")
        artifact_format = None
        artifact_size_bytes = None
        if opened:
            artifact_format = result["artifactFormat"]
            artifact_size_bytes = result["artifactSizeBytes"]
        return document, WordBackendAcquisition(
            document=document,
            revision=revision,
            persistence_state=persistence_state,
            read_only=read_only,
            artifact_format=artifact_format,
            artifact_size_bytes=artifact_size_bytes,
        )

    def _establish(self, operation, arguments, context, *, path=None):
        self._begin_acquisition()
        try:
            result = self._execute(operation, arguments, context)
            document, acquisition = self._acquisition(
                result,
                authorized_path=path,
                opened=path is not None,
            )
        except BackendActionFailure as failure:
            if (
                failure.outcome == "failed"
                and failure.binding_disposition == "unchanged"
            ):
                self._preparation = None
                raise WordDefiniteEstablishFailure(
                    code=failure.code,
                    message=failure.message,
                ) from failure
            raise WordUnprovableEstablishFailure(
                outcome=failure.outcome,
                code=failure.code,
                message=failure.message,
            ) from failure
        self._document = document
        self._persistence_path = path
        return acquisition

    def create_document(self, preparation, context):
        state = self._require_preparation(preparation, path=None)
        return self._establish(
            "acquire_new_document",
            {"preparationId": state.preparation_id},
            context,
        )

    def open_document(self, preparation, path, context):
        if not isinstance(path, str) or not path:
            raise ValueError("open path must be non-empty")
        state = self._require_preparation(preparation, path=path)
        return self._establish(
            "acquire_existing_document",
            {"preparationId": state.preparation_id},
            context,
            path=path,
        )

    def _require_bound_document(self, document):
        if document is not self._document:
            raise ValueError(
                "Windows Word Backend received another document"
            )
        return document

    def is_document_live(self, document):
        document = self._require_bound_document(document)
        result = self._execute(
            "probe_bound_document",
            {"documentId": document.bridge_document_id},
            None,
        )
        result = _require_exact_keys(
            result,
            {"live"},
            "Windows Word liveness result",
        )
        return _require_boolean(result["live"], "document liveness")

    def invoke(self, document, operation, context):
        document = self._require_bound_document(document)
        if not isinstance(operation, WordBackendOperation):
            raise ValueError("Windows Word Backend requires an operation")
        arguments = {
            "documentId": document.bridge_document_id,
            "operationArguments": dict(operation.arguments),
        }
        if self._persistence_path is not None:
            arguments["authorizedPath"] = self._persistence_path
        result = self._execute(operation.name, arguments, context)
        if operation.name == 'save_as_artifact':
            if result.get('artifact', {}).get('path') != operation.arguments['outputPath']:
                raise ValueError('Save As returned another output path')
            self._persistence_path = operation.arguments['outputPath']
        return result

    def close(self):
        return self._bridge.close()
