"""Word path and text format validators."""

import re
import unicodedata


def _host_absolute_path(value, suffixes):
    if any(
        unicodedata.category(character) in {"Cc", "Cs"}
        for character in value
    ):
        return False
    drive_absolute = re.match(r"^[A-Za-z]:[\\/]", value) is not None
    unc_absolute = re.match(
        r"^\\\\[^\\/]+[\\/][^\\/]+[\\/].+$",
        value,
    ) is not None
    posix_absolute = value.startswith("/") and not value.startswith("//")
    if not (drive_absolute or unc_absolute or posix_absolute):
        return False
    if (drive_absolute or unc_absolute) and not _windows_path_is_valid(value):
        return False
    return value.casefold().endswith(suffixes)


def _windows_path_is_valid(value):
    tail = value[3:] if re.match(r"^[A-Za-z]:[\\/]", value) else value[2:]
    components = re.split(r"[\\/]", tail)
    if not components or any(not component for component in components):
        return False
    reserved = re.compile(
        r"^(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?$",
        re.IGNORECASE,
    )
    return all(
        re.search(r'[<>:"|?*]', component) is None
        and not component.endswith((" ", "."))
        and reserved.fullmatch(component) is None
        for component in components
    )


def _word_run_text(value):
    return bool(value) and all(
        unicodedata.category(character) not in {"Cc", "Cs"}
        and character not in {"\u2028", "\u2029", "\f", "\ufffc"}
        for character in value
    )


def _word_story_text(value):
    return all(
        character in {"\t", "\n"}
        or (
            unicodedata.category(character) not in {"Cc", "Cs"}
            and character not in {"\u2028", "\u2029", "\ufffc"}
        )
        for character in value
    ) and "\r" not in value


def _word_normalized_text(value):
    return all(
        character in {"\t", "\n", "\f", "\u2028", "\ufffc"}
        or (
            unicodedata.category(character) not in {"Cc", "Cs"}
            and character != "\u2029"
        )
        for character in value
    ) and "\r" not in value


def _word_query_text(value):
    forbidden = {
        "\t", "\n", "\r", "\u2028", "\u2029", "\f", "\ufffc"
    }
    return bool(value) and not any(
        character in forbidden
        or unicodedata.category(character) in {"Cc", "Cs"}
        for character in value
    )


WORD_FORMAT_VALIDATORS = {
    "absoluteDocxPath": lambda value: _host_absolute_path(
        value,
        (".docx",),
    ),
    "absolutePdfPath": lambda value: _host_absolute_path(
        value,
        (".pdf",),
    ),
    "absoluteImagePath": lambda value: _host_absolute_path(
        value,
        (".png", ".jpg", ".jpeg"),
    ),
    "wordRunText": _word_run_text,
    "wordStoryText": _word_story_text,
    "wordNormalizedText": _word_normalized_text,
    "wordQueryText": _word_query_text,
}
