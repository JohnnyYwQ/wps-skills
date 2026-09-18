"""Strict JSON decoding and atomic durable file publication."""

import json
import os
from pathlib import Path
import tempfile
import time

def _unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError("duplicate JSON key")
        result[key] = value
    return result


def _nonfinite(value):
    raise ValueError("non-finite JSON number")

def decode(value):
    result = json.loads(value, object_pairs_hook=_unique_object, parse_constant=_nonfinite)
    json.dumps(result, allow_nan=False)
    return result


def _read(path):
    return decode(_retry_windows_file_operation(path.read_text, encoding="utf-8-sig"))


def _retry_windows_file_operation(operation, *args, **kwargs):
    # A Windows reader can briefly hold a non-delete-sharing handle.
    # Retry only the filesystem operation, never an Action.
    deadline = time.monotonic() + 1
    while True:
        try:
            return operation(*args, **kwargs)
        except PermissionError:
            if os.name != "nt" or time.monotonic() >= deadline:
                raise
            time.sleep(0.01)


def _write(path, value, *, once=False):
    """Publish complete JSON atomically; once=True never replaces a request."""
    encoded = json.dumps(value, ensure_ascii=True, allow_nan=False) + "\n"
    descriptor, name = tempfile.mkstemp(prefix=".write-", dir=str(path.parent))
    temporary = Path(name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
        if once:
            os.link(str(temporary), str(path))
        else:
            _retry_windows_file_operation(os.replace, str(temporary), str(path))
    finally:
        # once=True publishes a hard link: a request reader can also block
        # deletion of this temporary name for the same underlying file.
        _retry_windows_file_operation(temporary.unlink, missing_ok=True)


