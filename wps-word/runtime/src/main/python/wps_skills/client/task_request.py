"""Bounded, data-only Task Requests and references to earlier Action Responses."""

import re


MAX_STEPS = 128
IDENTIFIER = re.compile(r"[A-Za-z0-9][A-Za-z0-9_-]{0,79}\Z")


class TaskRequestError(ValueError):
    def __init__(self, message, *, step_id=None, code="INVALID_TASK_REQUEST"):
        super().__init__(message)
        self.step_id = step_id
        self.code = code


def _object(value, required, optional=()):
    if not isinstance(value, dict) or not set(required) <= set(value) or set(value) - set(required) - set(optional):
        raise TaskRequestError("Expected object fields: " + ", ".join(required))


def _identifier(value):
    if not isinstance(value, str) or not IDENTIFIER.fullmatch(value):
        raise TaskRequestError("IDs must be 1-80 ASCII letters, digits, underscores or hyphens, starting with a letter or digit")


def _reference(value, available):
    _object(value, ("step", "path"))
    if not isinstance(value["step"], str) or value["step"] not in available:
        raise TaskRequestError("A result reference must name an available earlier step")
    path = value["path"]
    if (not isinstance(path, list) or not 1 <= len(path) <= 32
            or any(not isinstance(key, str) and (type(key) is not int or key < 0) for key in path)):
        raise TaskRequestError("Reference path must contain 1-32 string keys or nonnegative integer indexes")


def validate_value(value, available, depth=0):
    if depth > 32:
        raise TaskRequestError("Task values may be nested at most 32 levels")
    has_reference = False
    if isinstance(value, dict):
        if "$ref" in value:
            _object(value, ("$ref",))
            _reference(value["$ref"], available)
            return True
        for child in value.values():
            has_reference = validate_value(child, available, depth + 1) or has_reference
    elif isinstance(value, list):
        for child in value:
            has_reference = validate_value(child, available, depth + 1) or has_reference
    return has_reference


def result_steps(result):
    """All execution records in dispatch order, including fixed lifecycle work."""
    if "document" in result:
        completion = result["completion"]
        return [result["document"], *result["steps"], *(
            step for step in (completion["save"], completion["pdf"]) if step is not None
        )]
    return result["steps"]


def resolve(value, responses):
    if isinstance(value, dict):
        if "$ref" in value:
            ref = value["$ref"]
            result = responses[ref["step"]]
            try:
                for key in ref["path"]:
                    if isinstance(result, dict) and isinstance(key, str):
                        result = result[key]
                    elif isinstance(result, list) and type(key) is int:
                        result = result[key]
                    else:
                        raise KeyError(key)
            except (KeyError, IndexError):
                raise TaskRequestError("Result reference does not exist: " + str(ref), code="TASK_REFERENCE_UNAVAILABLE")
            return result
        return {key: resolve(child, responses) for key, child in value.items()}
    if isinstance(value, list):
        return [resolve(child, responses) for child in value]
    return value
