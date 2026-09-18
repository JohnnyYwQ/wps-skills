"""Bounded output-name policy shared by Word handlers and contract validation."""

import ntpath


MAX_OUTPUT_ATTEMPTS = 20


def output_candidate(path, attempt):
    """Preserve the exact directory and extension; attempt zero is the request."""
    if attempt == 0:
        return path
    stem, extension = ntpath.splitext(path)
    return "{} ({}){}".format(stem, attempt, extension)
