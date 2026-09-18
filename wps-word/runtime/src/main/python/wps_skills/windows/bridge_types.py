"""Application-independent facts and failures at the Windows bridge seam."""

from dataclasses import dataclass
from typing import Optional


class BackendActionFailure(Exception):
    """Closed operation failure reported by a WPS Backend."""

    def __init__(
        self,
        *,
        outcome: str,
        code: str,
        message: str,
        binding_disposition: str,
    ):
        super().__init__(message)
        if outcome not in {"failed", "unknown"}:
            raise ValueError("invalid WPS Action failure outcome")
        if not isinstance(code, str) or not code:
            raise ValueError("WPS Action failure code must be non-empty")
        if not isinstance(message, str) or not message:
            raise ValueError("WPS Action failure message must be non-empty")
        if binding_disposition not in {
            "unchanged",
            "lost",
            "unprovable",
        }:
            raise ValueError("invalid WPS Action binding disposition")
        if outcome == "unknown" and binding_disposition == "lost":
            raise ValueError(
                "an unknown WPS Action cannot prove binding loss"
            )
        self.outcome = outcome
        self.code = code
        self.message = message
        self.binding_disposition = binding_disposition


@dataclass(frozen=True)
class WindowsDocument:
    """Opaque in-process reference to the bridge's one exact COM document."""

    bridge_document_id: str
    authorized_path: Optional[str]

    def __post_init__(self) -> None:
        if (
            not isinstance(self.bridge_document_id, str)
            or not self.bridge_document_id
        ):
            raise ValueError("bridge document id must be non-empty")
        if self.authorized_path is not None and (
            not isinstance(self.authorized_path, str)
            or not self.authorized_path
        ):
            raise ValueError("authorized path must be non-empty when present")


