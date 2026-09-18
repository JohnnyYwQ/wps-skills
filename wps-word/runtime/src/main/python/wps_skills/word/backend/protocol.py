"""Platform-independent Word Backend protocol and handler type."""

from typing import Any, Callable, Mapping, Protocol, runtime_checkable
from wps_skills.core.action_runtime import ControllerContext, ControllerResult
from wps_skills.word.backend.operation import (
    WordBackendOperation,
)
from wps_skills.word.backend.acquisition import (
    WordBackendAcquisition,
)
from wps_skills.word.backend.preparation import (
    WordBackendPreparation,
)


@runtime_checkable
class WordBackend(Protocol):
    """Platform seam used inside the Word Adapter implementation."""

    def prepare_create_document(
        self,
        context: ControllerContext,
    ) -> WordBackendPreparation:
        ...

    def prepare_open_document(
        self,
        path: str,
        context: ControllerContext,
    ) -> WordBackendPreparation:
        ...

    def create_document(
        self,
        preparation: WordBackendPreparation,
        context: ControllerContext,
    ) -> WordBackendAcquisition:
        ...

    def open_document(
        self,
        preparation: WordBackendPreparation,
        path: str,
        context: ControllerContext,
    ) -> WordBackendAcquisition:
        ...

    def is_document_live(self, document: Any) -> bool:
        ...

    def invoke(
        self,
        document: Any,
        operation: WordBackendOperation,
        context: ControllerContext,
    ) -> Mapping[str, Any]:
        ...

    def close(self) -> bool:
        ...


WordActionHandler = Callable[
    [WordBackend, Any, Mapping[str, Any], ControllerContext],
    ControllerResult,
]
