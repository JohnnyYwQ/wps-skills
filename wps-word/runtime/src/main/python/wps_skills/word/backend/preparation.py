"""Prepared Word document acquisition."""

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class WordBackendPreparation:
    """Opaque platform preparation plus its cross-process identity."""

    coordination_identity: str
    state: Any

    def __post_init__(self) -> None:
        if (
            not isinstance(self.coordination_identity, str)
            or not self.coordination_identity
        ):
            raise ValueError("Word coordination identity must be non-empty")
