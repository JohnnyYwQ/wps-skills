"""Acquired Word document and cleanup ownership."""

from dataclasses import dataclass
from typing import Any, Optional


@dataclass(frozen=True)
class WordBackendAcquisition:
    """One exact document plus platform observations needed by establish."""

    document: Any
    revision: str
    persistence_state: str
    read_only: bool
    artifact_format: Optional[str] = None
    artifact_size_bytes: Optional[int] = None

    def __post_init__(self) -> None:
        if self.document is None:
            raise ValueError("Word Backend acquisition requires a document")
        if not isinstance(self.revision, str) or not self.revision:
            raise ValueError("Word Backend acquisition requires a revision")
        if self.persistence_state not in {"unsaved", "saved", "modified"}:
            raise ValueError("invalid Word Backend persistence state")
        if not isinstance(self.read_only, bool):
            raise ValueError("Word Backend read-only state must be Boolean")
        if (
            self.artifact_format is not None
            and self.artifact_format != "docx"
        ):
            raise ValueError("Word establish artifacts must use docx")
        if (
            self.artifact_size_bytes is not None
            and (
                not isinstance(self.artifact_size_bytes, int)
                or isinstance(self.artifact_size_bytes, bool)
                or self.artifact_size_bytes < 1
            )
        ):
            raise ValueError("Word establish artifact size must be positive")
