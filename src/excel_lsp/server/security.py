"""Filesystem-scope enforcement for public tool paths."""

from __future__ import annotations

import os
from pathlib import Path

from excel_lsp.core.errors import ErrorCode, ExcelLSPError


def resolve_workbook_path(path: str | Path) -> Path:
    """Resolve a workbook path and enforce ``EXCEL_LSP_ROOT`` when configured."""
    try:
        workbook = Path(path).expanduser().resolve(strict=False)
    except (OSError, RuntimeError, ValueError) as exc:
        raise ExcelLSPError(ErrorCode.NOT_FOUND, f"Workbook path is invalid: {path!s}") from exc

    configured = os.environ.get("EXCEL_LSP_ROOT")
    if configured is None:
        return workbook
    roots = tuple(
        root
        for item in configured.split(os.pathsep)
        if item.strip()
        for root in (Path(item.strip()).expanduser().resolve(strict=False),)
        if root.is_dir()
    )
    if not roots or not any(_is_within(workbook, root) for root in roots):
        raise ExcelLSPError(
            ErrorCode.PATH_DENIED,
            f"Workbook path is outside EXCEL_LSP_ROOT: {workbook}",
            hint="Choose a workbook beneath an allowed root or update EXCEL_LSP_ROOT.",
        )
    return workbook


def _is_within(path: Path, root: Path) -> bool:
    try:
        return os.path.commonpath(
            (os.path.normcase(str(path)), os.path.normcase(str(root)))
        ) == os.path.normcase(str(root))
    except ValueError:
        return False


__all__ = ["resolve_workbook_path"]
