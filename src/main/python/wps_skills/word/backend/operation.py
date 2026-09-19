"""Immutable operations sent to the Word Backend."""

from dataclasses import dataclass
from typing import Any, Mapping
from wps_skills.word.contracts import WORD_TARGET_CONTRACT_SET
from wps_skills.core.action_runtime import _freeze_json


_WORD_ACTION_NAMES = frozenset(WORD_TARGET_CONTRACT_SET.action_names)


@dataclass(frozen=True)
class WordBackendOperation:
    """Backend-local operation produced by a Word Action handler."""

    name: str
    arguments: Mapping[str, Any]

    def __post_init__(self) -> None:
        if not isinstance(self.name, str) or not self.name:
            raise ValueError("Word Backend operation name must be non-empty")
        if self.name in _WORD_ACTION_NAMES:
            raise ValueError(
                "Word Backend operation must not reuse a public Action name"
            )
        if not isinstance(self.arguments, Mapping):
            raise ValueError("Word Backend operation arguments must be an object")
        object.__setattr__(self, "arguments", _freeze_json(self.arguments))
