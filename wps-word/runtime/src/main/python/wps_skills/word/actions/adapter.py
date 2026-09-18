"""Core-facing Word Adapter and exact document binding."""

from types import MappingProxyType
from typing import Any, Mapping
from wps_skills.core.action_runtime import (
    AcquiredDocument,
    ActionError,
    ControllerCommand,
    ControllerResult,
    DefiniteEstablishFailure,
    PreparedDocumentAcquisition,
    UnprovableEstablishFailure,
)
from wps_skills.word.contracts import WORD_TARGET_CONTRACT_SET
from wps_skills.word.backend.acquisition import (
    WordBackendAcquisition,
)
from wps_skills.word.backend.preparation import (
    WordBackendPreparation,
)
from wps_skills.word.backend.exceptions import (
    WordDefiniteEstablishFailure,
    WordUnprovableEstablishFailure,
)
from wps_skills.word.backend.protocol import (
    WordActionHandler,
    WordBackend,
)


class WordAdapter:
    """Core-facing Word Adapter; platform behavior stays behind WordBackend."""

    application = "word"
    contracts = WORD_TARGET_CONTRACT_SET

    def __init__(
        self,
        *,
        backend: WordBackend,
        handlers: Mapping[str, WordActionHandler],
        contracts=WORD_TARGET_CONTRACT_SET,
    ):
        if not isinstance(backend, WordBackend):
            raise ValueError("backend must satisfy the Word Backend seam")
        if not isinstance(handlers, Mapping):
            raise ValueError("Word handlers must be a mapping")
        if getattr(contracts, "application", None) != "word":
            raise ValueError("Word contracts must belong to word")
        required_actions = frozenset(
            contract.name
            for contract in contracts.contracts
            if contract.binding_role == "required"
        )
        invalid = [
            name
            for name, handler in handlers.items()
            if (
                not isinstance(name, str)
                or not name
                or name not in required_actions
                or not callable(handler)
            )
        ]
        if invalid:
            raise ValueError("invalid Word Action handler mapping")
        missing = required_actions.difference(handlers)
        if contracts is not WORD_TARGET_CONTRACT_SET and missing:
            raise ValueError("registered Word contracts require every handler")
        self.contracts = contracts
        self._backend = backend
        self._handlers = MappingProxyType(dict(handlers))

    @staticmethod
    def _require_word_command(command: ControllerCommand) -> None:
        if not isinstance(command, ControllerCommand):
            raise ValueError("Word Adapter requires a Controller Command")
        if command.address.app != "word":
            raise ValueError("Word Adapter received another application")

    @staticmethod
    def _require_acquisition(acquisition) -> WordBackendAcquisition:
        if not isinstance(acquisition, WordBackendAcquisition):
            raise TypeError("Word Backend returned an invalid acquisition")
        return acquisition

    @classmethod
    def _created(cls, acquisition) -> AcquiredDocument:
        acquisition = cls._require_acquisition(acquisition)
        if (
            acquisition.persistence_state != "unsaved"
            or acquisition.read_only
            or acquisition.artifact_format is not None
            or acquisition.artifact_size_bytes is not None
        ):
            raise TypeError("Word Backend returned invalid create observations")
        return AcquiredDocument(
            document=acquisition.document,
            data={
                "revision": acquisition.revision,
                "documentState": {
                    "persistenceState": acquisition.persistence_state,
                    "readOnly": acquisition.read_only,
                },
            },
        )

    @classmethod
    def _opened(cls, acquisition, requested_path) -> AcquiredDocument:
        acquisition = cls._require_acquisition(acquisition)
        if (
            acquisition.persistence_state not in {"saved", "modified"}
            or acquisition.artifact_format != "docx"
            or acquisition.artifact_size_bytes is None
        ):
            raise TypeError("Word Backend returned invalid open observations")
        return AcquiredDocument(
            document=acquisition.document,
            data={
                "revision": acquisition.revision,
                "artifact": {
                    "path": requested_path,
                    "format": acquisition.artifact_format,
                    "sizeBytes": acquisition.artifact_size_bytes,
                },
                "documentState": {
                    "persistenceState": acquisition.persistence_state,
                    "readOnly": acquisition.read_only,
                },
            },
        )

    @staticmethod
    def _translate_establish_failure(exc):
        if isinstance(exc, WordDefiniteEstablishFailure):
            raise DefiniteEstablishFailure(
                code=exc.code,
                message=exc.message,
            ) from exc
        if isinstance(exc, WordUnprovableEstablishFailure):
            raise UnprovableEstablishFailure(
                outcome=exc.outcome,
                code=exc.code,
                message=exc.message,
                partial_document=exc.partial_document,
            ) from exc
        raise exc

    def prepare_establish(
        self,
        command: ControllerCommand,
    ) -> PreparedDocumentAcquisition:
        self._require_word_command(command)
        try:
            from wps_skills.word.actions.registry import WORD_ACTIONS
            entry = WORD_ACTIONS.get(command.address.action)
            if entry is None or entry.prepare is None:
                raise DefiniteEstablishFailure(
                    code="INVALID_PARAMS",
                    message="Action is not a Word establish Action",
                )
            preparation = entry.prepare(self._backend, command.params, command.context)
        except (WordDefiniteEstablishFailure, WordUnprovableEstablishFailure) as exc:
            self._translate_establish_failure(exc)
        if not isinstance(preparation, WordBackendPreparation):
            raise TypeError("Word Backend returned an invalid preparation")
        return PreparedDocumentAcquisition(
            coordination_identity=preparation.coordination_identity,
            application_state=preparation,
        )

    def establish(
        self,
        prepared: PreparedDocumentAcquisition,
        command: ControllerCommand,
    ) -> AcquiredDocument:
        self._require_word_command(command)
        if not isinstance(prepared, PreparedDocumentAcquisition) or not isinstance(
            prepared.application_state,
            WordBackendPreparation,
        ):
            raise TypeError("Word Adapter requires its prepared acquisition")
        preparation = prepared.application_state
        if preparation.coordination_identity != prepared.coordination_identity:
            raise TypeError("Word acquisition identities do not match")
        try:
            from wps_skills.word.actions.registry import WORD_ACTIONS
            entry = WORD_ACTIONS.get(command.address.action)
            if entry is None or entry.observe is None:
                raise DefiniteEstablishFailure(
                    code="INVALID_PARAMS",
                    message="Action is not a Word establish Action",
                )
            acquisition = entry.handler(
                self._backend, preparation, command.params, command.context,
            )
            acquired = entry.observe(self, acquisition, command.params)
        except (WordDefiniteEstablishFailure, WordUnprovableEstablishFailure) as exc:
            self._translate_establish_failure(exc)
        return acquired

    def is_live(self, document: Any) -> bool:
        live = self._backend.is_document_live(document)
        if not isinstance(live, bool):
            raise TypeError("Word Backend liveness observation must be Boolean")
        return live

    def handle(
        self,
        document: Any,
        command: ControllerCommand,
    ) -> ControllerResult:
        self._require_word_command(command)
        handler = self._handlers.get(command.address.action)
        if handler is None:
            return ControllerResult.failed(
                error=ActionError(
                    code="WORD_CAPABILITY_UNAVAILABLE",
                    message=(
                        f"Word Action is not implemented: "
                        f"{command.address.action}"
                    ),
                ),
                controller_state="usable",
                binding_disposition="unchanged",
            )
        return handler(
            self._backend,
            document,
            command.params,
            command.context,
        )

    def handle_none(self, command: ControllerCommand) -> ControllerResult:
        self._require_word_command(command)
        raise ValueError("The Word Application Contract Set has no none Action")

    def close(self) -> bool:
        return self._backend.close()
