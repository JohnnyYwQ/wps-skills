"""Definite and unprovable Word acquisition failures."""

from typing import Any


class WordDefiniteEstablishFailure(Exception):
    """Backend proved that establish produced no document or binding effect."""

    def __init__(self, *, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message


class WordUnprovableEstablishFailure(Exception):
    """Backend cannot prove the complete establish effect."""

    def __init__(
        self,
        *,
        outcome: str,
        code: str,
        message: str,
        partial_document: Any = None,
    ):
        super().__init__(message)
        if outcome not in {"failed", "unknown"}:
            raise ValueError("invalid Word establish outcome")
        self.outcome = outcome
        self.code = code
        self.message = message
        self.partial_document = partial_document
