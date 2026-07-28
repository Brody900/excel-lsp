"""Typed diagnostic records and P5-owned diagnostic construction."""

from __future__ import annotations

import math
import os
import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from types import MappingProxyType
from typing import Literal, cast
from urllib.parse import unquote, urlsplit

from excel_lsp.core.external_links import external_link_label
from excel_lsp.core.models import Rect
from excel_lsp.core.parse.coordinates import column_label
from excel_lsp.core.symbols import cell_symbol_id, formula_block_symbol_id

DiagnosticSeverity = Literal["error", "warn", "info"]

DIAGNOSTIC_SEVERITIES: Mapping[str, frozenset[DiagnosticSeverity]] = MappingProxyType(
    {
        "E_ERRVAL": frozenset({"error"}),
        "E_CIRCULAR": frozenset({"error"}),
        # A missing local target is an error. Remote or otherwise unverifiable
        # targets retain the same semantic code but are warnings, per HANDOFF 5.6.
        "E_BROKEN_XLINK": frozenset({"error", "warn"}),
        "W_POSSIBLE_CIRCULAR": frozenset({"warn"}),
        "W_INCONSISTENT_FORMULA": frozenset({"warn"}),
        "W_UNKNOWN_NAME": frozenset({"warn"}),
        "W_PARSE": frozenset({"warn"}),
        "W_LARGE_SHEET": frozenset({"warn"}),
        "W_REGEX_TIMEOUT": frozenset({"warn"}),
        "I_DYNAMIC_REF": frozenset({"info"}),
        "I_VOLATILE": frozenset({"info"}),
        "I_STALE": frozenset({"info"}),
    }
)

P5_PERSISTED_CODES = frozenset({"E_ERRVAL", "E_BROKEN_XLINK", "I_VOLATILE"})
P5_DEFERRED_CODES = frozenset({"W_REGEX_TIMEOUT", "I_STALE"})

_SEVERITY_ORDER: Mapping[DiagnosticSeverity, int] = MappingProxyType(
    {"error": 0, "warn": 1, "info": 2}
)
_RECOGNIZED_ERROR_MESSAGES: Mapping[str, str] = MappingProxyType(
    {
        "#REF!": "invalid cell or range reference",
        "#DIV/0!": "division by zero",
        "#N/A": "value is not available",
        "#VALUE!": "invalid value or operand type",
        "#NAME?": "unrecognized formula name",
        "#NUM!": "invalid numeric result",
        "#SPILL!": "dynamic-array result cannot spill",
        "#CALC!": "calculation engine error",
        "#BLOCKED!": "calculation was blocked",
    }
)
_MAX_ERROR_VALUE_DISPLAY = 120
_MAX_EXTERNAL_TARGET_LENGTH = 4_096
_MAX_RELATED_DEPTH = 64
_WINDOWS_DRIVE = re.compile(r"^[A-Za-z]:[\\/]")
_WINDOWS_FILE_URI_PATH = re.compile(r"^/[A-Za-z]:/")


def _empty_related() -> Mapping[str, object]:
    return MappingProxyType({})


def _freeze(value: object, *, depth: int = 0) -> object:
    if isinstance(value, Mapping):
        if depth >= _MAX_RELATED_DEPTH:
            raise TypeError("diagnostic related data exceeds the nesting limit")
        result: dict[str, object] = {}
        for key, item in cast(Mapping[object, object], value).items():
            if not isinstance(key, str):
                raise TypeError("diagnostic related keys must be strings")
            result[key] = _freeze(item, depth=depth + 1)
        return MappingProxyType(result)
    if isinstance(value, (list, tuple)):
        if depth >= _MAX_RELATED_DEPTH:
            raise TypeError("diagnostic related data exceeds the nesting limit")
        return tuple(_freeze(item, depth=depth + 1) for item in cast(Sequence[object], value))
    if isinstance(value, float) and not math.isfinite(value):
        raise TypeError("diagnostic related floats must be finite")
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    raise TypeError(
        f"diagnostic related values must be JSON-compatible, got {type(value).__name__}"
    )


@dataclass(frozen=True, slots=True)
class Diagnostic:
    """One public diagnostic with immutable, JSON-compatible related data."""

    severity: DiagnosticSeverity
    code: str
    sheet: str
    row: int | None
    col: int | None
    ref: str
    message: str
    related: Mapping[str, object] = field(default_factory=_empty_related)

    def __post_init__(self) -> None:
        allowed = DIAGNOSTIC_SEVERITIES.get(self.code)
        if allowed is None:
            raise ValueError(f"unknown diagnostic code: {self.code}")
        if self.severity not in allowed:
            raise ValueError(f"diagnostic {self.code} does not allow severity {self.severity}")
        if not self.sheet or not self.ref or not self.message:
            raise ValueError("diagnostic sheet, ref, and message must not be empty")
        if (self.row is None) != (self.col is None):
            raise ValueError("diagnostic row and column must both be present or absent")
        if self.row is not None and (self.row < 1 or self.col is None or self.col < 1):
            raise ValueError("diagnostic coordinates must be positive")
        frozen = _freeze(self.related)
        if not isinstance(frozen, Mapping):  # pragma: no cover - field type boundary
            raise TypeError("diagnostic related data must be a mapping")
        object.__setattr__(self, "related", frozen)


@dataclass(frozen=True, slots=True)
class DiagnosticReport:
    """Deterministic filtered diagnostics plus pre-limit aggregate counts."""

    diagnostics: tuple[Diagnostic, ...]
    total: int
    counts_by_severity: Mapping[str, int]
    counts_by_code: Mapping[str, int]
    truncated: bool

    def __post_init__(self) -> None:
        if self.total < len(self.diagnostics):
            raise ValueError("diagnostic total cannot be smaller than returned diagnostics")
        if self.truncated != (self.total > len(self.diagnostics)):
            raise ValueError("diagnostic truncation flag must match total and returned counts")
        for counts in (self.counts_by_severity, self.counts_by_code):
            if any(type(key) is not str for key in counts):
                raise ValueError("diagnostic count keys must be strings")
            if any(type(value) is not int or value < 0 for value in counts.values()):
                raise ValueError("diagnostic counts must use string keys and nonnegative integers")
            if sum(counts.values()) != self.total:
                raise ValueError("diagnostic counts must each sum to the report total")
        object.__setattr__(
            self,
            "counts_by_severity",
            MappingProxyType(dict(self.counts_by_severity)),
        )
        object.__setattr__(
            self,
            "counts_by_code",
            MappingProxyType(dict(self.counts_by_code)),
        )


def diagnostic_sort_key(diagnostic: Diagnostic) -> tuple[object, ...]:
    """Return the shared public ordering used by persistence and queries."""
    return (
        _SEVERITY_ORDER[diagnostic.severity],
        diagnostic.sheet.casefold(),
        -1 if diagnostic.row is None else diagnostic.row,
        -1 if diagnostic.col is None else diagnostic.col,
        diagnostic.code,
        diagnostic.ref,
        diagnostic.message,
    )


def error_value_diagnostic(
    sheet: str,
    row: int,
    col: int,
    cell_ref: str,
    value: object,
) -> Diagnostic:
    """Create E_ERRVAL solely from an OOXML error-typed stored cell."""
    raw_value = "" if value is None else str(value)
    display = raw_value
    if len(display) > _MAX_ERROR_VALUE_DISPLAY:
        display = f"{display[: _MAX_ERROR_VALUE_DISPLAY - 1]}…"
    explanation = _RECOGNIZED_ERROR_MESSAGES.get(raw_value)
    if explanation is None:
        message = f"Cell contains an Excel error value ({display or 'unknown error'})."
    else:
        message = f"Cell contains {raw_value}: {explanation}."
    return Diagnostic(
        severity="error",
        code="E_ERRVAL",
        sheet=sheet,
        row=row,
        col=col,
        ref=cell_symbol_id(sheet, cell_ref),
        message=message,
        related={"errorValue": display},
    )


def volatile_formula_diagnostic(
    sheet: str,
    block_n: int,
    row: int,
    col: int,
    cell_ref: str,
) -> Diagnostic:
    """Create one I_VOLATILE finding per volatile formula block."""
    return Diagnostic(
        severity="info",
        code="I_VOLATILE",
        sheet=sheet,
        row=row,
        col=col,
        ref=cell_symbol_id(sheet, cell_ref),
        message="Formula block contains a volatile function and recalculates with Excel.",
        related={"block": formula_block_symbol_id(sheet, block_n)},
    )


def regex_timeout_diagnostic(
    sheet: str,
    ref: str,
    *,
    deadline_ms: int,
) -> Diagnostic:
    """Construct the transient P7 find-guard warning using the P5 catalog."""
    if type(deadline_ms) is not int or deadline_ms < 1:
        raise ValueError("regex deadline must be a positive integer")
    return Diagnostic(
        severity="warn",
        code="W_REGEX_TIMEOUT",
        sheet=sheet,
        row=None,
        col=None,
        ref=ref,
        message="Regular-expression search reached its safety deadline.",
        related={"deadlineMs": deadline_ms},
    )


def stale_range_diagnostic(
    sheet: str,
    rect: Rect,
    *,
    since: str,
) -> Diagnostic:
    """Construct the P6 staleness finding for one persisted stale rectangle."""
    if not since:
        raise ValueError("staleness timestamp must not be empty")
    first = f"{column_label(rect.col_min)}{rect.row_min}"
    last = f"{column_label(rect.col_max)}{rect.row_max}"
    local_ref = first if first == last else f"{first}:{last}"
    return Diagnostic(
        severity="info",
        code="I_STALE",
        sheet=sheet,
        row=rect.row_min,
        col=rect.col_min,
        ref=f"{sheet}!{local_ref}",
        message="Values in this range may be stale until Excel recalculates and saves.",
        related={"range": local_ref, "since": since},
    )


def external_link_diagnostic(
    workbook_path: str | Path,
    sheet: str,
    link_index: int,
    target: str,
) -> Diagnostic | None:
    """Check one external-link relationship without exposing its raw target."""
    label = external_link_label(target)
    status, reason = _external_target_status(Path(workbook_path), target)
    if status == "ok":
        return None
    if status == "missing":
        severity: DiagnosticSeverity = "error"
        message = f"External workbook {label} was not found on disk."
    elif status == "remote":
        severity = "warn"
        message = f"External workbook {label} is remote and could not be checked locally."
    else:
        severity = "warn"
        message = f"External workbook {label} could not be checked safely."
    return Diagnostic(
        severity=severity,
        code="E_BROKEN_XLINK",
        sheet=sheet,
        row=None,
        col=None,
        ref=f"external:{label}",
        message=message,
        related={
            "linkIndex": link_index,
            "status": status,
            "target": label,
            "reason": reason,
        },
    )


def external_link_health_snapshot(
    workbook_path: str | Path,
    external_links: Mapping[int, str],
) -> tuple[tuple[int, str, str], ...]:
    """Return a path-free snapshot whose changes invalidate link diagnostics."""
    workbook = Path(workbook_path)
    return tuple(
        (link_index, *_external_target_status(workbook, target))
        for link_index, target in sorted(external_links.items())
    )


def _external_target_status(workbook_path: Path, target: str) -> tuple[str, str]:
    stripped = target.strip()
    if not stripped or len(stripped) > _MAX_EXTERNAL_TARGET_LENGTH:
        return "uncheckable", "invalid-target"
    if stripped.startswith(("\\\\", "//")):
        return "remote", "network-path"
    if _WINDOWS_DRIVE.match(stripped):
        if os.name != "nt":
            return "uncheckable", "foreign-local-path"
        candidate_text = unquote(stripped)
    else:
        try:
            parsed = urlsplit(stripped)
        except ValueError:
            return "uncheckable", "invalid-uri"
        scheme = parsed.scheme.casefold()
        if scheme and scheme != "file":
            return "remote", "non-file-uri"
        if scheme == "file":
            if "@" in parsed.netloc:
                return "uncheckable", "credentialed-file-uri"
            if parsed.netloc:
                return "remote", "network-file-uri"
            candidate_text = unquote(parsed.path)
            if _WINDOWS_FILE_URI_PATH.match(candidate_text):
                if os.name != "nt":
                    return "uncheckable", "foreign-local-path"
                candidate_text = candidate_text[1:]
        else:
            candidate_text = unquote(parsed.path)
    if not candidate_text or "\x00" in candidate_text:
        return "uncheckable", "invalid-path"
    if candidate_text.startswith(("\\\\", "//")):
        return "remote", "network-path"
    try:
        candidate = Path(candidate_text.replace("/", os.sep))
        if not candidate.is_absolute():
            candidate = workbook_path.expanduser().resolve().parent / candidate
        if str(candidate).startswith(("\\\\", "//")):
            return "remote", "network-path"
        return ("ok", "local-file") if candidate.is_file() else ("missing", "local-file")
    except (OSError, RuntimeError, ValueError):
        return "uncheckable", "path-check-failed"


__all__ = [
    "DIAGNOSTIC_SEVERITIES",
    "P5_DEFERRED_CODES",
    "P5_PERSISTED_CODES",
    "Diagnostic",
    "DiagnosticReport",
    "DiagnosticSeverity",
    "diagnostic_sort_key",
    "error_value_diagnostic",
    "external_link_diagnostic",
    "external_link_health_snapshot",
    "regex_timeout_diagnostic",
    "stale_range_diagnostic",
    "volatile_formula_diagnostic",
]
