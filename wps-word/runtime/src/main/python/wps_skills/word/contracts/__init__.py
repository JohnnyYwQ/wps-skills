"""Word contract discovery; definitions and executable admission stay separate."""

from wps_skills.word.contracts.definitions import WORD_TARGET_CONTRACT_SET, WORD_TARGET_ACTION_INDEX
from wps_skills.word.contracts.formats import WORD_FORMAT_VALIDATORS


def __getattr__(name):
    if name in {"WORD_PRODUCTION_CONTRACT_SET", "WORD_PRODUCTION_ACTION_INDEX"}:
        from wps_skills.word.actions import registry
        return getattr(registry, name)
    raise AttributeError(name)
