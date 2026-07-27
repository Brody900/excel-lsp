"""Safe public labels for workbook external-link targets."""

from __future__ import annotations

import re
from pathlib import PurePosixPath
from urllib.parse import unquote, urlsplit

_MAX_EXTERNAL_TARGET_LENGTH = 4_096
_MAX_EXTERNAL_DECODE_ROUNDS = 16
_NEUTRAL_EXTERNAL_LINK = "[external-workbook]"
_URI_SCHEME = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*:")
_WINDOWS_DRIVE = re.compile(r"^[A-Za-z]:[\\/]")
_HTTP_URI = re.compile(r"^https?://", re.IGNORECASE)
_FILE_URI = re.compile(r"^file://", re.IGNORECASE)
_WORKBOOK_SUFFIXES = (
    ".xlsx",
    ".xlsm",
    ".xltx",
    ".xltm",
    ".xlsb",
    ".xls",
    ".xlt",
    ".csv",
    ".tsv",
    ".ods",
    ".fods",
)


def external_link_label(target: str) -> str:
    """Return a bounded workbook label without credentials or path details."""
    decoded_target = _bounded_unquote(target.strip())
    if decoded_target is None:
        return _NEUTRAL_EXTERNAL_LINK
    path = decoded_target
    try:
        if _WINDOWS_DRIVE.match(decoded_target):
            path = decoded_target
        elif decoded_target.startswith(("//", "\\\\")):
            path = urlsplit(f"https:{decoded_target.replace(chr(92), '/')}").path
        elif _HTTP_URI.match(decoded_target):
            parsed = urlsplit(decoded_target)
            if parsed.scheme.casefold() not in {"http", "https"} or not parsed.hostname:
                return _NEUTRAL_EXTERNAL_LINK
            path = parsed.path
        elif _FILE_URI.match(decoded_target):
            parsed = urlsplit(decoded_target)
            if parsed.scheme.casefold() != "file" or "@" in parsed.netloc:
                return _NEUTRAL_EXTERNAL_LINK
            path = parsed.path
        elif _URI_SCHEME.match(decoded_target) or "://" in decoded_target:
            return _NEUTRAL_EXTERNAL_LINK
    except ValueError:
        return _NEUTRAL_EXTERNAL_LINK

    decoded = _bounded_unquote(path)
    if decoded is None:
        return _NEUTRAL_EXTERNAL_LINK
    for delimiter in ("?", "#", ";"):
        decoded = decoded.split(delimiter, 1)[0]
    normalized = decoded.replace("\\", "/").rstrip("/")
    name = PurePosixPath(normalized).name.strip()
    if name.startswith("[") and name.endswith("]"):
        name = name[1:-1].strip()
    if not _is_safe_workbook_name(name):
        return _NEUTRAL_EXTERNAL_LINK
    return f"[{name}]"


def _bounded_unquote(value: str) -> str | None:
    if len(value) > _MAX_EXTERNAL_TARGET_LENGTH:
        return None
    decoded = value
    for _round in range(_MAX_EXTERNAL_DECODE_ROUNDS):
        next_decoded = unquote(decoded)
        if next_decoded == decoded:
            return decoded
        decoded = next_decoded
    return decoded if unquote(decoded) == decoded else None


def _is_safe_workbook_name(name: str) -> bool:
    if not name or name in {".", ".."} or len(name) > 255:
        return False
    if any(ord(character) < 32 for character in name):
        return False
    if any(
        character in name
        for character in (
            "/",
            "\\",
            ":",
            "@",
            "?",
            "#",
            ";",
            "&",
            "=",
            "+",
            "[",
            "]",
            "%",
        )
    ):
        return False
    folded = name.casefold()
    for suffix in _WORKBOOK_SUFFIXES:
        if folded.endswith(suffix):
            return bool(name[: -len(suffix)].strip(" ."))
    return False


__all__ = ["external_link_label"]
