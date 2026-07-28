"""Shared bounded tool implementation for MCP and the debugging CLI."""

from __future__ import annotations

import json
import time
from collections import Counter
from collections.abc import Callable, Iterable, Mapping, Sequence
from pathlib import Path
from typing import Any, Literal, cast

import regex

from excel_lsp.core.diagnostics import DiagnosticSeverity
from excel_lsp.core.edit import CellEdit
from excel_lsp.core.edit import set_column_formula as core_set_column_formula
from excel_lsp.core.edit import write_cells as core_write_cells
from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.formulas.a1 import CellRef
from excel_lsp.core.formulas.r1c1 import to_r1c1
from excel_lsp.core.graph.models import GraphArea, GraphTarget, TraceNode
from excel_lsp.core.index.lifecycle import ensure_fresh, index_workbook
from excel_lsp.core.index.store import IndexStore
from excel_lsp.core.models import IndexUpdate, Rect
from excel_lsp.core.parse.coordinates import make_cell_ref, parse_cell_ref, parse_rect
from excel_lsp.core.symbols import (
    cell_symbol_id,
    column_symbol_id,
    defined_name_symbol_id,
    formula_block_symbol_id,
    region_symbol_id,
    sheet_symbol_id,
)
from excel_lsp.core.values import normalize_value
from excel_lsp.core.workbook_map import build_workbook_map
from excel_lsp.server.cursors import decode_cursor, encode_cursor, parameter_hash
from excel_lsp.server.models import WriteCellInput
from excel_lsp.server.security import resolve_workbook_path

RESPONSE_CHARACTER_CAP = 8_000
RAW_VALUE_CAP = 200
_SYMBOL_KINDS = frozenset({"sheets", "regions", "columns", "names", "fblocks", "cells"})
_FIND_KINDS = frozenset({"values", "headers", "formulas", "names"})


class ToolService:
    """Stateless facade whose methods correspond one-for-one with MCP tools."""

    def open_workbook(
        self,
        path: str,
        *,
        progress: Callable[[int, int, str], None] | None = None,
    ) -> dict[str, Any]:
        workbook, update = self._fresh(path, progress=progress)
        payload = cast(dict[str, Any], build_workbook_map(workbook))
        payload["reindexed"] = update.changed
        return _fit_payload(
            payload,
            list_keys=("sheetList", "names", "externalLinks", "hints"),
        )

    def refresh(
        self,
        path: str,
        recalculated: bool = False,
        *,
        progress: Callable[[int, int, str], None] | None = None,
    ) -> dict[str, Any]:
        workbook = resolve_workbook_path(path)
        update = index_workbook(workbook, recalculated=recalculated, progress=progress)
        payload = cast(dict[str, Any], build_workbook_map(workbook))
        payload.update(
            {
                "reindexed": update.changed,
                "reindexedSheets": list(update.reindexed_sheets),
            }
        )
        return _fit_payload(
            payload,
            list_keys=("sheetList", "names", "externalLinks", "hints", "reindexedSheets"),
        )

    def sheet_names(self, path: str) -> tuple[str, ...]:
        """Return current sheet names for progress notifications."""
        _, update = self._fresh(path)
        with IndexStore(update.index_path) as store:
            rows = store.connection.execute("SELECT name FROM sheets ORDER BY id").fetchall()
        return tuple(str(row["name"]) for row in rows)

    def list_symbols(
        self,
        path: str,
        query: str = "",
        kinds: Sequence[str] | None = None,
    ) -> dict[str, Any]:
        _, update = self._fresh(path)
        selected = _validated_selection(kinds, _SYMBOL_KINDS, "symbol kind")
        needle = query.casefold()
        symbols: list[dict[str, str]] = []
        total = 0

        def add_symbol(symbol: dict[str, str]) -> None:
            nonlocal total
            if needle and not (
                needle in symbol["id"].casefold() or needle in symbol["description"].casefold()
            ):
                return
            total += 1
            symbols.append(symbol)
            if len(symbols) > 200:
                symbols.sort(key=lambda item: (item["id"].casefold(), item["id"]))
                del symbols[100:]

        with IndexStore(update.index_path) as store:
            connection = store.connection
            if "sheets" in selected:
                for row in connection.execute(
                    "SELECT name, kind, visibility, max_row, max_col FROM sheets ORDER BY id"
                ):
                    add_symbol(
                        {
                            "id": sheet_symbol_id(str(row["name"])),
                            "kind": "sheet",
                            "description": (
                                f"{row['kind']} {row['visibility']} "
                                f"{row['max_row']}x{row['max_col']}"
                            ),
                        }
                    )
            if "regions" in selected:
                for row in connection.execute(
                    "SELECT s.name, r.n, r.kind, r.row_min, r.row_max, r.col_min, r.col_max "
                    "FROM regions r JOIN sheets s ON s.id=r.sheet_id ORDER BY s.id,r.n"
                ):
                    add_symbol(
                        {
                            "id": region_symbol_id(str(row["name"]), int(row["n"])),
                            "kind": "region",
                            "description": f"{row['kind']} {_rect_ref(_row_rect(row))}",
                        }
                    )
            if "columns" in selected:
                for row in connection.execute(
                    "SELECT s.name, r.n, c.norm_header, c.header, c.dtype "
                    "FROM columns c JOIN regions r ON r.id=c.region_id "
                    "JOIN sheets s ON s.id=r.sheet_id ORDER BY s.id,r.n,c.idx"
                ):
                    add_symbol(
                        {
                            "id": column_symbol_id(
                                str(row["name"]), int(row["n"]), str(row["norm_header"])
                            ),
                            "kind": "column",
                            "description": f"{_short(str(row['header']), 96)} ({row['dtype']})",
                        }
                    )
            if "names" in selected:
                for row in connection.execute(
                    "SELECT d.name, d.refers_to, d.kind, s.name scope FROM defined_names d "
                    "LEFT JOIN sheets s ON s.id=d.scope_sheet_id "
                    "ORDER BY COALESCE(s.id,0),d.name,d.id"
                ):
                    add_symbol(
                        {
                            "id": defined_name_symbol_id(
                                str(row["name"]),
                                scope_sheet=None if row["scope"] is None else str(row["scope"]),
                            ),
                            "kind": "name",
                            "description": f"{row['kind']} {_short(str(row['refers_to']), 120)}",
                        }
                    )
            if "fblocks" in selected:
                for row in connection.execute(
                    "SELECT s.name, f.n, f.r1c1, f.row_min, f.row_max, f.col_min, f.col_max "
                    "FROM fblocks f JOIN sheets s ON s.id=f.sheet_id ORDER BY s.id,f.n"
                ):
                    add_symbol(
                        {
                            "id": formula_block_symbol_id(str(row["name"]), int(row["n"])),
                            "kind": "fblock",
                            "description": (
                                f"{_rect_ref(_row_rect(row))} {_short(str(row['r1c1']), 100)}"
                            ),
                        }
                    )
            if "cells" in selected:
                for row in connection.execute(
                    "SELECT s.name, c.ref, c.value_type, c.formula FROM cells c "
                    "JOIN sheets s ON s.id=c.sheet_id ORDER BY s.id,c.row,c.col"
                ):
                    add_symbol(
                        {
                            "id": cell_symbol_id(str(row["name"]), str(row["ref"])),
                            "kind": "cell",
                            "description": (
                                "formula" if row["formula"] is not None else str(row["value_type"])
                            ),
                        }
                    )

        symbols.sort(key=lambda item: (item["id"].casefold(), item["id"]))
        payload: dict[str, Any] = {
            "symbols": symbols[:100],
            "total": total,
            "truncated": total > 100,
            "reindexed": update.changed,
            "nextSteps": ["Use a symbol id with schema, trace, or profile tools."],
        }
        return _fit_payload(payload, list_keys=("symbols",))

    def get_region_schema(self, path: str, region_id: str) -> dict[str, Any]:
        _, update = self._fresh(path)
        with IndexStore(update.index_path) as store:
            connection = store.connection
            row = _region_row(connection, region_id)
            columns = connection.execute(
                "SELECT idx,header,norm_header,dtype,nonnull,distinct_est,formula_block_id "
                "FROM columns WHERE region_id=? ORDER BY idx",
                (int(row["id"]),),
            ).fetchall()
            ncols = len(columns)
            sample_rows = max(0, min(3, 180 // ncols)) if ncols else 0
            body_start = int(row["row_min"]) + int(row["header_rows"])
            samples = (
                _read_grid(
                    connection,
                    int(row["sheet_id"]),
                    Rect(
                        body_start,
                        min(int(row["row_max"]), body_start + sample_rows - 1),
                        int(row["col_min"]),
                        int(row["col_max"]),
                    ),
                )
                if sample_rows and body_start <= int(row["row_max"])
                else []
            )
            validations = _validation_summaries(connection, row, columns)
            fblocks = [
                {
                    "id": formula_block_symbol_id(str(row["sheet"]), int(item["n"])),
                    "range": _rect_ref(_row_rect(item)),
                    "r1c1": _short(str(item["r1c1"]), 240),
                    "volatile": bool(item["volatile"]),
                    "opaque": bool(item["opaque"]),
                }
                for item in connection.execute(
                    "SELECT n,r1c1,row_min,row_max,col_min,col_max,volatile,opaque "
                    "FROM fblocks WHERE sheet_id=? AND row_max>=? AND row_min<=? "
                    "AND col_max>=? AND col_min<=? ORDER BY n",
                    (
                        int(row["sheet_id"]),
                        int(row["row_min"]),
                        int(row["row_max"]),
                        int(row["col_min"]),
                        int(row["col_max"]),
                    ),
                )
            ]
            stale = _area_is_stale(connection, str(row["sheet"]), _row_rect(row))
        rendered_columns = [
            {
                "id": column_symbol_id(str(row["sheet"]), int(row["n"]), str(item["norm_header"])),
                "header": _short(str(item["header"]), 160),
                "dtype": str(item["dtype"]),
                "nonnull": int(item["nonnull"]),
                "distinct": int(item["distinct_est"]),
                "validation": validations.get(int(item["idx"]), []),
            }
            for item in columns
        ]
        payload = {
            "region": region_id,
            "sheet": str(row["sheet"]),
            "range": _rect_ref(_row_rect(row)),
            "kind": str(row["kind"]),
            "headerRows": int(row["header_rows"]),
            "columns": rendered_columns,
            "samples": samples,
            "formulaBlocks": fblocks,
            "conf": float(row["confidence"]),
            "stale": stale,
            "reindexed": update.changed,
            "truncated": False,
            "nextSteps": ["Trace formula blocks; use read_range only for necessary values."],
        }
        if ncols and sample_rows == 0:
            payload["samplesOmitted"] = f"{ncols} columns; use read_range"
        return _fit_payload(payload, list_keys=("columns", "formulaBlocks", "samples"))

    def read_range(
        self,
        path: str,
        ref: str,
        cursor: str | None = None,
        max_cells: int = RAW_VALUE_CAP,
    ) -> dict[str, Any]:
        if type(max_cells) is not int or not 1 <= max_cells <= RAW_VALUE_CAP:
            raise ExcelLSPError(
                ErrorCode.INVALID_VALUE,
                f"max_cells must be an integer from 1 through {RAW_VALUE_CAP}.",
            )
        workbook, update = self._fresh(path)
        with IndexStore(update.index_path) as store:
            sheet, rect = _resolve_area(store.connection, ref)
            generation = store.generation
            params = {"path": str(workbook), "ref": ref, "max_cells": max_cells}
            digest = parameter_hash(params)
            offset = (
                0
                if cursor is None
                else decode_cursor(
                    cursor,
                    tool="read_range",
                    params_hash=digest,
                    generation=generation,
                )
            )
            total = (rect.row_max - rect.row_min + 1) * (rect.col_max - rect.col_min + 1)
            if offset > total:
                raise ExcelLSPError(
                    ErrorCode.STALE_CURSOR,
                    "The pagination cursor offset is outside the requested range.",
                    hint="Re-issue the original query without a cursor.",
                )
            page_size = min(max_cells, total - offset)
            while True:
                values, page_bounds, value_truncated = _read_grid_page(
                    store.connection, sheet, rect, offset, page_size
                )
                next_offset = offset + page_size
                truncated = next_offset < total
                next_cursor = (
                    encode_cursor(
                        tool="read_range",
                        params_hash=digest,
                        offset=next_offset,
                        generation=generation,
                    )
                    if truncated
                    else None
                )
                payload = {
                    "sheet": sheet,
                    "range": _rect_ref(rect),
                    "page": {"start": page_bounds[0], "end": page_bounds[1]},
                    "values": values,
                    "offset": offset,
                    "totalCells": total,
                    "truncated": truncated,
                    "cursor": next_cursor,
                    "stale": _area_is_stale(store.connection, sheet, rect),
                    "reindexed": update.changed,
                }
                if value_truncated:
                    payload["valueTruncated"] = True
                if _serialized_length(payload) <= RESPONSE_CHARACTER_CAP:
                    return payload
                if page_size == 1:
                    return _fit_single_range_value(payload)
                page_size = max(1, page_size // 2)

    def find(
        self,
        path: str,
        pattern: str,
        search_in: Sequence[str] | None = None,
        sheet: str | None = None,
        max_results: int = 50,
    ) -> dict[str, Any]:
        if not pattern or len(pattern) > 256:
            raise ExcelLSPError(
                ErrorCode.INVALID_VALUE,
                "find pattern must contain between 1 and 256 characters.",
            )
        if type(max_results) is not int or not 1 <= max_results <= 50:
            raise ExcelLSPError(ErrorCode.INVALID_VALUE, "find max must be from 1 through 50.")
        selected = _validated_selection(search_in, _FIND_KINDS, "find target")
        _, update = self._fresh(path)
        try:
            expression = regex.compile(pattern)
        except regex.error as exc:
            raise ExcelLSPError(
                ErrorCode.INVALID_VALUE,
                f"Invalid regular expression: {_short(str(exc), 160)}",
            ) from exc
        matches: list[dict[str, str]] = []
        deadline = time.monotonic() + 2.0
        timed_out = False
        value_stale: bool | None = None
        with IndexStore(update.index_path) as store:
            if "values" in selected:
                value_stale = _search_scope_is_stale(store.connection, sheet)
            subjects = _find_subjects(store.connection, selected, sheet)
            for ref_value, kind, subject in subjects:
                if time.monotonic() >= deadline:
                    timed_out = True
                    break
                bounded = subject[:1000]
                remaining = max(0.001, deadline - time.monotonic())
                try:
                    match = expression.search(bounded, timeout=remaining)
                except TimeoutError:
                    timed_out = True
                    break
                if match is None:
                    continue
                matches.append(
                    {"ref": ref_value, "kind": kind, "snippet": _snippet(bounded, match)}
                )
                if len(matches) >= max_results:
                    break
        payload = {
            "matches": matches,
            "truncated": timed_out or len(matches) >= max_results,
            "reindexed": update.changed,
        }
        if value_stale is not None:
            payload["stale"] = value_stale
        if timed_out:
            payload["warnings"] = [
                {
                    "code": "W_REGEX_TIMEOUT",
                    "message": "Regex search reached the 2-second safety deadline.",
                }
            ]
        return _fit_payload(payload, list_keys=("matches",))

    def trace_precedents(
        self, path: str, ref_or_symbol: str, depth: int = 2, max_nodes: int = 200
    ) -> dict[str, Any]:
        return self._trace(path, ref_or_symbol, "precedents", depth, max_nodes)

    def trace_dependents(
        self, path: str, ref_or_symbol: str, depth: int = 2, max_nodes: int = 200
    ) -> dict[str, Any]:
        return self._trace(path, ref_or_symbol, "dependents", depth, max_nodes)

    def trace_path(
        self,
        path: str,
        from_ref_or_symbol: str,
        to_ref_or_symbol: str,
        max_paths: int = 3,
        max_depth: int = 12,
    ) -> dict[str, Any]:
        _, update = self._fresh(path)
        with IndexStore(update.index_path) as store:
            source = _resolve_graph_target(store.connection, from_ref_or_symbol)
            destination = _resolve_graph_target(store.connection, to_ref_or_symbol)
            try:
                result = store.dependency_graph.trace_path(
                    source, destination, max_paths=max_paths, max_depth=max_depth
                )
            except ValueError as exc:
                raise ExcelLSPError(ErrorCode.INVALID_VALUE, str(exc)) from exc
        payload = {
            "connected": result.connected,
            "paths": [
                [{"symbol": step.symbol, "via": step.via} for step in path_items]
                for path_items in result.paths
            ],
            "truncated": result.truncated,
            "reindexed": update.changed,
        }
        return _fit_payload(payload, list_keys=("paths",))

    def explain_formula(self, path: str, ref: str) -> dict[str, Any]:
        _, update = self._fresh(path)
        with IndexStore(update.index_path) as store:
            sheet, rect = _resolve_area(store.connection, ref, require_cell=True)
            row = store.connection.execute(
                "SELECT c.formula,c.formula_kind,c.ref FROM cells c "
                "JOIN sheets s ON s.id=c.sheet_id "
                "WHERE s.name=? AND c.row=? AND c.col=?",
                (sheet, rect.row_min, rect.col_min),
            ).fetchone()
            if row is None or row["formula"] is None:
                raise ExcelLSPError(
                    ErrorCode.INVALID_REF, f"Cell does not contain a formula: {ref}"
                )
            block = store.connection.execute(
                "SELECT f.n,f.r1c1,f.row_min,f.row_max,f.col_min,f.col_max,f.volatile,f.opaque "
                "FROM fblocks f JOIN sheets s ON s.id=f.sheet_id WHERE s.name=? "
                "AND f.row_min<=? AND f.row_max>=? AND f.col_min<=? AND f.col_max>=? "
                "ORDER BY f.n LIMIT 1",
                (sheet, rect.row_min, rect.row_min, rect.col_min, rect.col_min),
            ).fetchone()
            diagnostic_report = store.get_diagnostics(sheet=sheet, max_results=100)
            cell_id = cell_symbol_id(sheet, str(row["ref"]))
            diagnostics = [
                _diagnostic_dict(item)
                for item in diagnostic_report.diagnostics
                if item.ref == cell_id
            ]
            via_rows = store.connection.execute(
                "SELECT DISTINCT e.via FROM edges e JOIN fblocks f ON e.src_kind='fblock' "
                "AND f.id=e.src_id JOIN sheets s ON s.id=f.sheet_id WHERE s.name=? "
                "AND f.row_min<=? AND f.row_max>=? AND f.col_min<=? AND f.col_max>=? "
                "ORDER BY e.via",
                (sheet, rect.row_min, rect.row_min, rect.col_min, rect.col_min),
            ).fetchall()
        formula = str(row["formula"])
        payload = {
            "ref": cell_id,
            "a1": formula,
            "r1c1": str(block["r1c1"])
            if block is not None
            else to_r1c1(formula, CellRef(rect.row_min, rect.col_min)),
            "block": None if block is None else formula_block_symbol_id(sheet, int(block["n"])),
            "extent": None if block is None else _rect_ref(_row_rect(block)),
            "resolvedNames": [
                str(item["via"]) for item in via_rows if "name" in str(item["via"]).casefold()
            ],
            "structuredRefs": [
                str(item["via"]) for item in via_rows if "structured" in str(item["via"]).casefold()
            ],
            "volatile": False if block is None else bool(block["volatile"]),
            "opaque": False if block is None else bool(block["opaque"]),
            "diagnostics": diagnostics,
            "reindexed": update.changed,
        }
        return _fit_payload(payload, list_keys=("diagnostics", "resolvedNames", "structuredRefs"))

    def get_diagnostics(
        self,
        path: str,
        sheet: str | None = None,
        severity: str | None = None,
        code: str | None = None,
        max_results: int = 100,
    ) -> dict[str, Any]:
        _, update = self._fresh(path)
        with IndexStore(update.index_path) as store:
            try:
                report = store.get_diagnostics(
                    sheet=sheet,
                    severity=cast(DiagnosticSeverity | None, severity),
                    code=code,
                    max_results=max_results,
                )
            except ValueError as exc:
                raise ExcelLSPError(ErrorCode.INVALID_VALUE, str(exc)) from exc
        payload = {
            "diagnostics": [_diagnostic_dict(item) for item in report.diagnostics],
            "total": report.total,
            "counts": {
                "severity": dict(report.counts_by_severity),
                "code": dict(report.counts_by_code),
            },
            "truncated": report.truncated,
            "reindexed": update.changed,
        }
        return _fit_payload(payload, list_keys=("diagnostics",))

    def profile_column(self, path: str, col_symbol_or_ref: str) -> dict[str, Any]:
        _, update = self._fresh(path)
        with IndexStore(update.index_path) as store:
            sheet, rect, distinct_est = _resolve_column(store, col_symbol_or_ref)
            rows = store.connection.execute(
                "SELECT value,value_type,formula FROM cells c JOIN sheets s ON s.id=c.sheet_id "
                "WHERE s.name=? AND c.row BETWEEN ? AND ? AND c.col BETWEEN ? AND ? "
                "ORDER BY c.row,c.col",
                (sheet, rect.row_min, rect.row_max, rect.col_min, rect.col_max),
            ).fetchall()
            stale = _area_is_stale(store.connection, sheet, rect)
        values = [normalize_value(row["value"]) for row in rows if row["value"] is not None]
        numeric = [
            float(value)
            for value in values
            if isinstance(value, (int, float)) and not isinstance(value, bool)
        ]
        nonnull_rows = [row for row in rows if row["value"] is not None]
        formula_count = sum(row["formula"] is not None for row in rows)
        missing_formula_values = sum(
            row["formula"] is not None and row["value"] is None for row in rows
        )
        payload: dict[str, Any] = {
            "column": col_symbol_or_ref,
            "sheet": sheet,
            "range": _rect_ref(rect),
            "count": (rect.row_max - rect.row_min + 1) * (rect.col_max - rect.col_min + 1),
            "nonnull": len(values),
            "distinct": distinct_est if distinct_est is not None else len(set(map(str, values))),
            "cachedValues": not formula_count or missing_formula_values == 0,
            "stale": stale,
            "reindexed": update.changed,
        }
        if values and all(row["value_type"] == "number" for row in nonnull_rows):
            payload.update(
                {
                    "sum": sum(numeric),
                    "mean": sum(numeric) / len(numeric),
                    "min": min(numeric),
                    "max": max(numeric),
                }
            )
        else:
            counts = Counter(_short(str(value), 120) for value in values)
            payload["topValues"] = [
                {"value": value, "count": count}
                for value, count in sorted(counts.items(), key=lambda item: (-item[1], item[0]))[:5]
            ]
        if not payload["cachedValues"]:
            payload["hint"] = "Open in Excel, recalculate, save, then refresh(recalculated=true)."
        return _fit_payload(payload, list_keys=("topValues",))

    def write_cells(self, path: str, cells: Sequence[WriteCellInput]) -> dict[str, Any]:
        workbook = resolve_workbook_path(path)
        if not 1 <= len(cells) <= 500:
            raise ExcelLSPError(ErrorCode.INVALID_VALUE, "cells must contain 1 through 500 edits.")
        update = ensure_fresh(workbook)
        with IndexStore(update.index_path) as store:
            default_sheet = _default_sheet(store.connection)
        edits: list[CellEdit] = []
        results: list[dict[str, Any]] = []
        for item in cells:
            try:
                has_value = "value" in item.model_fields_set
                has_formula = "formula" in item.model_fields_set
                if has_value == has_formula:
                    raise ExcelLSPError(
                        ErrorCode.INVALID_VALUE,
                        "Exactly one of value or formula must be supplied.",
                    )
                sheet, local = _split_cell_ref(item.ref, default_sheet)
                parse_cell_ref(local)
                if has_formula:
                    if not item.formula or not item.formula.startswith("="):
                        raise ExcelLSPError(
                            ErrorCode.INVALID_VALUE,
                            "Formula must be a nonempty string beginning with '='.",
                        )
                    edits.append(CellEdit.formula(sheet, local, item.formula))
                else:
                    edits.append(CellEdit.value(sheet, local, item.value))
                results.append({"ref": item.ref, "ok": True})
            except ValueError as exc:
                results.append(
                    {
                        "ref": item.ref,
                        "ok": False,
                        "error": {
                            "code": ErrorCode.INVALID_REF.value,
                            "message": str(exc),
                        },
                    }
                )
            except ExcelLSPError as exc:
                results.append({"ref": item.ref, "ok": False, **exc.as_dict()})
        if not edits:
            return _fit_payload(
                {
                    "results": results,
                    "resultsTotal": len(results),
                    "staleBlocks": 0,
                    "reindexed": update.changed,
                },
                list_keys=("results",),
            )
        try:
            result = core_write_cells(workbook, edits)
        except ExcelLSPError as exc:
            for item in results:
                if item["ok"]:
                    item["ok"] = False
                    item["error"] = exc.as_dict()["error"]
            return _fit_payload(
                {
                    "results": results,
                    "resultsTotal": len(results),
                    "staleBlocks": 0,
                    "reindexed": update.changed,
                },
                list_keys=("results",),
            )
        return _fit_payload(
            {
                "results": results,
                "resultsTotal": len(results),
                "staleBlocks": result.stale_blocks,
                "generation": result.generation,
                "reindexed": update.changed,
            },
            list_keys=("results",),
        )

    def set_column_formula(
        self,
        path: str,
        col_symbol: str,
        formula: str,
        overwrite: bool = False,
    ) -> dict[str, Any]:
        workbook = resolve_workbook_path(path)
        update = ensure_fresh(workbook)
        result = core_set_column_formula(workbook, col_symbol, formula, overwrite=overwrite)
        return _fit_payload(
            {
                "formulaBlock": result.formula_block,
                "cellsWritten": result.cells_written,
                "staleBlocks": result.edit.stale_blocks,
                "generation": result.edit.generation,
                "reindexed": update.changed,
            }
        )

    def _fresh(
        self,
        path: str,
        *,
        progress: Callable[[int, int, str], None] | None = None,
    ) -> tuple[Path, IndexUpdate]:
        workbook = resolve_workbook_path(path)
        return workbook, ensure_fresh(workbook, progress=progress)

    def _trace(
        self,
        path: str,
        ref_or_symbol: str,
        direction: Literal["precedents", "dependents"],
        depth: int,
        max_nodes: int,
    ) -> dict[str, Any]:
        _, update = self._fresh(path)
        with IndexStore(update.index_path) as store:
            target = _resolve_graph_target(store.connection, ref_or_symbol)
            try:
                graph = store.dependency_graph
                result = (
                    graph.trace_precedents(target, depth=depth, max_nodes=max_nodes)
                    if direction == "precedents"
                    else graph.trace_dependents(target, depth=depth, max_nodes=max_nodes)
                )
            except ValueError as exc:
                raise ExcelLSPError(ErrorCode.INVALID_VALUE, str(exc)) from exc
        payload = {
            "direction": direction,
            "tree": _trace_node_dict(result.root),
            "nodeCount": result.node_count,
            "edgeCount": result.edge_count,
            "truncated": result.truncated,
            "reindexed": update.changed,
        }
        return _fit_trace_payload(payload)


def _validated_selection(
    requested: Sequence[str] | None, allowed: frozenset[str], label: str
) -> frozenset[str]:
    if requested is None or len(requested) == 0:
        return allowed
    result = frozenset(requested)
    unknown = sorted(result - allowed)
    if unknown:
        raise ExcelLSPError(
            ErrorCode.INVALID_VALUE,
            f"Unknown {label}{'s' if len(unknown) != 1 else ''}: {', '.join(unknown)}.",
        )
    return result


def _row_rect(row: Mapping[str, object]) -> Rect:
    return Rect(
        int(cast(Any, row["row_min"])),
        int(cast(Any, row["row_max"])),
        int(cast(Any, row["col_min"])),
        int(cast(Any, row["col_max"])),
    )


def _rect_ref(rect: Rect) -> str:
    start = make_cell_ref(rect.row_min, rect.col_min)
    end = make_cell_ref(rect.row_max, rect.col_max)
    return start if start == end else f"{start}:{end}"


def _default_sheet(connection: Any) -> str:
    rows = connection.execute("SELECT name FROM sheets ORDER BY id LIMIT 2").fetchall()
    if not rows:
        raise ExcelLSPError(ErrorCode.NOT_FOUND, "Workbook has no indexed worksheets.")
    return str(rows[0]["name"])


def _split_qualified(value: str, default_sheet: str) -> tuple[str, str]:
    text = value.strip()
    if "!" not in text:
        return default_sheet, text
    sheet, local = text.rsplit("!", 1)
    sheet = sheet.strip()
    if len(sheet) >= 2 and sheet[0] == sheet[-1] == "'":
        sheet = sheet[1:-1].replace("''", "'")
    if not sheet or not local:
        raise ValueError(f"invalid qualified reference: {value!r}")
    return sheet, local


def _split_cell_ref(value: str, default_sheet: str) -> tuple[str, str]:
    text = value[5:] if value.startswith("cell:") else value
    return _split_qualified(text, default_sheet)


def _resolve_area(connection: Any, value: str, *, require_cell: bool = False) -> tuple[str, Rect]:
    default_sheet = _default_sheet(connection)
    if value.startswith("cell:"):
        sheet, local = _split_cell_ref(value, default_sheet)
        rect = parse_rect(local)
    elif value.startswith("sheet:"):
        sheet = value[6:]
        row = connection.execute(
            "SELECT max_row,max_col FROM sheets WHERE name=?", (sheet,)
        ).fetchone()
        if row is None:
            raise ExcelLSPError(ErrorCode.UNKNOWN_SYMBOL, f"Unknown symbol: {value}")
        rect = Rect(1, max(1, int(row["max_row"])), 1, max(1, int(row["max_col"])))
    elif value.startswith("fblock:"):
        pieces = value.split(":", 2)
        if len(pieces) != 3:
            raise ExcelLSPError(ErrorCode.UNKNOWN_SYMBOL, f"Unknown symbol: {value}")
        try:
            ordinal = int(pieces[2])
        except ValueError as exc:
            raise ExcelLSPError(ErrorCode.UNKNOWN_SYMBOL, f"Unknown symbol: {value}") from exc
        row = connection.execute(
            "SELECT f.row_min,f.row_max,f.col_min,f.col_max FROM fblocks f "
            "JOIN sheets s ON s.id=f.sheet_id WHERE s.name=? AND f.n=?",
            (pieces[1], ordinal),
        ).fetchone()
        if row is None:
            raise ExcelLSPError(ErrorCode.UNKNOWN_SYMBOL, f"Unknown symbol: {value}")
        sheet, rect = pieces[1], _row_rect(row)
    elif value.startswith("region:"):
        row = _region_row(connection, value)
        sheet, rect = str(row["sheet"]), _row_rect(row)
    elif value.startswith("col:"):
        row = _column_row(connection, value)
        sheet = str(row["sheet"])
        rect = Rect(
            int(row["row_min"]) + int(row["header_rows"]),
            int(row["row_max"]) - int(row["totals_rows"]),
            int(row["column"]),
            int(row["column"]),
        )
    elif value.startswith("name:"):
        row = _name_area_row(connection, value)
        sheet, rect = str(row["sheet"]), _row_rect(row)
    else:
        try:
            sheet, local = _split_qualified(value, default_sheet)
            rect = parse_rect(local)
        except ValueError as exc:
            raise ExcelLSPError(ErrorCode.INVALID_REF, f"Invalid reference: {value!r}") from exc
    known = connection.execute("SELECT 1 FROM sheets WHERE name=?", (sheet,)).fetchone()
    if known is None:
        raise ExcelLSPError(ErrorCode.INVALID_REF, f"Unknown worksheet: {sheet!r}")
    if require_cell and (rect.row_min != rect.row_max or rect.col_min != rect.col_max):
        raise ExcelLSPError(ErrorCode.INVALID_REF, f"Expected one cell reference: {value!r}")
    return sheet, rect


def _resolve_graph_target(connection: Any, value: str) -> GraphArea | GraphTarget:
    sheet, rect = _resolve_area(connection, value)
    row = connection.execute("SELECT id FROM sheets WHERE name=?", (sheet,)).fetchone()
    if row is None:
        raise ExcelLSPError(ErrorCode.INVALID_REF, f"Unknown worksheet: {sheet!r}")
    area = GraphArea(int(row["id"]), sheet, rect)
    if value.startswith("fblock:"):
        return GraphTarget("fblock", value, area.ref, area)
    if rect.row_min == rect.row_max and rect.col_min == rect.col_max:
        local = make_cell_ref(rect.row_min, rect.col_min)
        return GraphTarget("cell", cell_symbol_id(sheet, local), area.ref, area)
    return area


def _region_row(connection: Any, symbol: str) -> Any:
    pieces = symbol.split(":", 2)
    if len(pieces) != 3 or pieces[0] != "region":
        raise ExcelLSPError(ErrorCode.UNKNOWN_SYMBOL, f"Unknown region symbol: {symbol}")
    try:
        ordinal = int(pieces[2])
    except ValueError as exc:
        raise ExcelLSPError(ErrorCode.UNKNOWN_SYMBOL, f"Unknown region symbol: {symbol}") from exc
    row = connection.execute(
        "SELECT r.*,s.name sheet FROM regions r JOIN sheets s ON s.id=r.sheet_id "
        "WHERE s.name=? AND r.n=?",
        (pieces[1], ordinal),
    ).fetchone()
    if row is None:
        raise ExcelLSPError(ErrorCode.UNKNOWN_SYMBOL, f"Unknown region symbol: {symbol}")
    return row


def _column_row(connection: Any, symbol: str) -> Any:
    pieces = symbol.split(":", 3)
    if len(pieces) != 4 or pieces[0] != "col":
        raise ExcelLSPError(ErrorCode.UNKNOWN_SYMBOL, f"Unknown column symbol: {symbol}")
    try:
        ordinal = int(pieces[2])
    except ValueError as exc:
        raise ExcelLSPError(ErrorCode.UNKNOWN_SYMBOL, f"Unknown column symbol: {symbol}") from exc
    row = connection.execute(
        "SELECT s.name sheet,r.row_min,r.row_max,r.header_rows,"
        "COALESCE(lo.totals_rows,0) totals_rows,r.col_min+c.idx column,c.distinct_est "
        "FROM columns c JOIN regions r ON r.id=c.region_id JOIN sheets s ON s.id=r.sheet_id "
        "LEFT JOIN list_objects lo ON lo.sheet_id=r.sheet_id AND lo.name=r.list_object_name "
        "WHERE s.name=? AND r.n=? AND c.norm_header=?",
        (pieces[1], ordinal, pieces[3]),
    ).fetchone()
    if row is None:
        raise ExcelLSPError(ErrorCode.UNKNOWN_SYMBOL, f"Unknown column symbol: {symbol}")
    return row


def _name_area_row(connection: Any, symbol: str) -> Any:
    raw = symbol[5:]
    scope: str | None = None
    name = raw
    if "!" in raw:
        scope, name = raw.split("!", 1)
    row = connection.execute(
        "SELECT s.name sheet,a.row_min,a.row_max,a.col_min,a.col_max FROM defined_names d "
        "JOIN name_areas a ON a.name_id=d.id JOIN sheets s ON s.id=a.sheet_id "
        "LEFT JOIN sheets scope ON scope.id=d.scope_sheet_id WHERE d.name=? "
        "AND ((? IS NULL AND d.scope_sheet_id IS NULL) OR scope.name=?) ORDER BY a.id LIMIT 1",
        (name, scope, scope),
    ).fetchone()
    if row is None:
        raise ExcelLSPError(ErrorCode.UNKNOWN_SYMBOL, f"Unknown range symbol: {symbol}")
    return row


def _read_grid(connection: Any, sheet_id: int, rect: Rect) -> list[list[Any]]:
    cells = {
        (int(row["row"]), int(row["col"])): normalize_value(row["value"])
        for row in connection.execute(
            "SELECT row,col,value FROM cells WHERE sheet_id=? AND row BETWEEN ? AND ? "
            "AND col BETWEEN ? AND ? ORDER BY row,col",
            (sheet_id, rect.row_min, rect.row_max, rect.col_min, rect.col_max),
        )
    }
    return [
        [cells.get((row, col)) for col in range(rect.col_min, rect.col_max + 1)]
        for row in range(rect.row_min, rect.row_max + 1)
    ]


def _read_grid_page(
    connection: Any, sheet: str, rect: Rect, offset: int, count: int
) -> tuple[list[list[Any]], tuple[str, str], bool]:
    if count == 0:
        start = make_cell_ref(rect.row_min, rect.col_min)
        return [], (start, start), False
    width = rect.col_max - rect.col_min + 1
    coordinates: list[tuple[int, int]] = []
    for position in range(offset, offset + count):
        row_offset, col_offset = divmod(position, width)
        coordinates.append((rect.row_min + row_offset, rect.col_min + col_offset))
    first_row, first_col = coordinates[0]
    last_row, last_col = coordinates[-1]
    rows = connection.execute(
        "SELECT c.row,c.col,c.value FROM cells c JOIN sheets s ON s.id=c.sheet_id "
        "WHERE s.name=? AND c.row BETWEEN ? AND ? AND c.col BETWEEN ? AND ?",
        (sheet, first_row, last_row, rect.col_min, rect.col_max),
    ).fetchall()
    stored = {(int(row["row"]), int(row["col"])): normalize_value(row["value"]) for row in rows}
    result: list[list[Any]] = []
    current_row = -1
    value_truncated = False
    for row, col in coordinates:
        if row != current_row:
            result.append([])
            current_row = row
        value = stored.get((row, col))
        if isinstance(value, str) and len(value) > 4_000:
            value = _short(value, 4_000)
            value_truncated = True
        result[-1].append(value)
    return (
        result,
        (make_cell_ref(first_row, first_col), make_cell_ref(last_row, last_col)),
        value_truncated,
    )


def _validation_summaries(
    connection: Any, region: Any, columns: Sequence[Any]
) -> dict[int, list[dict[str, Any]]]:
    rows = connection.execute(
        "SELECT col_min,col_max,vtype,operator,formula1,formula2,allow_blank "
        "FROM validations WHERE sheet_id=? AND row_max>=? AND row_min<=? "
        "AND col_max>=? AND col_min<=? ORDER BY id",
        (
            int(region["sheet_id"]),
            int(region["row_min"]),
            int(region["row_max"]),
            int(region["col_min"]),
            int(region["col_max"]),
        ),
    ).fetchall()
    result: dict[int, list[dict[str, Any]]] = {int(item["idx"]): [] for item in columns}
    for row in rows:
        first_index = max(0, int(row["col_min"]) - int(region["col_min"]))
        last_index = min(len(columns) - 1, int(row["col_max"]) - int(region["col_min"]))
        summary = {
            "type": row["vtype"],
            "operator": row["operator"],
            "formula1": _short(str(row["formula1"]), 120) if row["formula1"] else None,
            "formula2": _short(str(row["formula2"]), 120) if row["formula2"] else None,
            "allowBlank": bool(row["allow_blank"]),
        }
        for index in range(first_index, last_index + 1):
            result[index].append(summary)
    return result


def _find_subjects(
    connection: Any, selected: frozenset[str], sheet: str | None
) -> Iterable[tuple[str, str, str]]:
    sheet_exists = (
        sheet is None
        or connection.execute("SELECT 1 FROM sheets WHERE name=?", (sheet,)).fetchone() is not None
    )
    if not sheet_exists:
        raise ExcelLSPError(ErrorCode.INVALID_REF, f"Unknown worksheet: {sheet!r}")
    sheet_clause = "" if sheet is None else " WHERE s.name=?"
    parameters = () if sheet is None else (sheet,)
    if "values" in selected or "formulas" in selected:
        rows = connection.execute(
            "SELECT s.name,c.ref,c.value,c.formula FROM cells c JOIN sheets s ON s.id=c.sheet_id"
            + sheet_clause
            + " ORDER BY s.id,c.row,c.col",
            parameters,
        )
        for row in rows:
            ref_value = cell_symbol_id(str(row["name"]), str(row["ref"]))
            if "values" in selected and row["value"] is not None:
                yield ref_value, "value", str(normalize_value(row["value"]))
            if "formulas" in selected and row["formula"] is not None:
                yield ref_value, "formula", str(row["formula"])
    if "headers" in selected:
        query = (
            "SELECT s.name,r.n,c.norm_header,c.header FROM columns c "
            "JOIN regions r ON r.id=c.region_id JOIN sheets s ON s.id=r.sheet_id"
            + sheet_clause
            + " ORDER BY s.id,r.n,c.idx"
        )
        for row in connection.execute(query, parameters):
            yield (
                column_symbol_id(str(row["name"]), int(row["n"]), str(row["norm_header"])),
                "header",
                str(row["header"]),
            )
    if "names" in selected:
        query = (
            "SELECT d.name,d.refers_to,s.name scope FROM defined_names d "
            "LEFT JOIN sheets s ON s.id=d.scope_sheet_id"
        )
        name_parameters: tuple[object, ...] = ()
        if sheet is not None:
            query += " WHERE s.name=? OR s.name IS NULL"
            name_parameters = (sheet,)
        query += " ORDER BY COALESCE(s.id,0),d.name,d.id"
        for row in connection.execute(query, name_parameters):
            yield (
                defined_name_symbol_id(
                    str(row["name"]),
                    scope_sheet=None if row["scope"] is None else str(row["scope"]),
                ),
                "name",
                f"{row['name']} {row['refers_to']}",
            )


def _snippet(subject: str, match: Any) -> str:
    start = max(0, match.start() - 30)
    end = min(len(subject), max(match.end(), match.start() + 1) + 30)
    prefix = "…" if start else ""
    suffix = "…" if end < len(subject) else ""
    return _short(f"{prefix}{subject[start:end]}{suffix}", 80)


def _trace_node_dict(node: TraceNode) -> dict[str, Any]:
    target = node.target
    return {
        "kind": target.kind,
        "symbol": target.symbol,
        "ref": target.ref,
        "via": node.via,
        "childCount": node.child_count,
        "children": [_trace_node_dict(child) for child in node.children],
    }


def _diagnostic_dict(item: Any) -> dict[str, Any]:
    return {
        "severity": item.severity,
        "code": item.code,
        "sheet": item.sheet,
        "row": item.row,
        "col": item.col,
        "ref": item.ref,
        "message": _short(item.message, 500),
        "related": dict(item.related),
    }


def _resolve_column(store: IndexStore, value: str) -> tuple[str, Rect, int | None]:
    if value.startswith("col:"):
        row = _column_row(store.connection, value)
        rect = Rect(
            int(row["row_min"]) + int(row["header_rows"]),
            int(row["row_max"]) - int(row["totals_rows"]),
            int(row["column"]),
            int(row["column"]),
        )
        return str(row["sheet"]), rect, int(row["distinct_est"])
    sheet, rect = _resolve_area(store.connection, value)
    if rect.col_min != rect.col_max:
        raise ExcelLSPError(ErrorCode.INVALID_REF, "Column profile requires one column.")
    candidates = store.connection.execute(
        "SELECT r.row_min+r.header_rows body_min,"
        "r.row_max-COALESCE(lo.totals_rows,0) body_max,c.distinct_est "
        "FROM columns c JOIN regions r ON r.id=c.region_id "
        "JOIN sheets s ON s.id=r.sheet_id "
        "LEFT JOIN list_objects lo ON lo.sheet_id=r.sheet_id AND lo.name=r.list_object_name "
        "WHERE s.name=? AND r.col_min+c.idx=? AND r.row_max>=? AND r.row_min<=? "
        "ORDER BY r.n",
        (sheet, rect.col_min, rect.row_min, rect.row_max),
    ).fetchall()
    if len(candidates) == 1:
        candidate = candidates[0]
        body_min = max(rect.row_min, int(candidate["body_min"]))
        body_max = min(rect.row_max, int(candidate["body_max"]))
        if body_min <= body_max:
            semantic_rect = Rect(body_min, body_max, rect.col_min, rect.col_max)
            complete_body = body_min == int(candidate["body_min"]) and body_max == int(
                candidate["body_max"]
            )
            return (
                sheet,
                semantic_rect,
                int(candidate["distinct_est"]) if complete_body else None,
            )
    return sheet, rect, None


def _area_is_stale(connection: Any, sheet: str, rect: Rect) -> bool:
    return (
        connection.execute(
            "SELECT 1 FROM staleness st JOIN sheets s ON s.id=st.sheet_id WHERE s.name=? "
            "AND st.row_max>=? AND st.row_min<=? AND st.col_max>=? AND st.col_min<=? LIMIT 1",
            (sheet, rect.row_min, rect.row_max, rect.col_min, rect.col_max),
        ).fetchone()
        is not None
    )


def _search_scope_is_stale(connection: Any, sheet: str | None) -> bool:
    if sheet is None:
        row = connection.execute("SELECT 1 FROM staleness LIMIT 1").fetchone()
    else:
        row = connection.execute(
            "SELECT 1 FROM staleness st JOIN sheets s ON s.id=st.sheet_id WHERE s.name=? LIMIT 1",
            (sheet,),
        ).fetchone()
    return row is not None


def _short(value: str, limit: int) -> str:
    if len(value) <= limit:
        return value
    return value[: max(0, limit - 1)] + "…"


def _serialized_length(payload: Mapping[str, Any]) -> int:
    return len(json.dumps(payload, ensure_ascii=False, indent=2, default=str))


def _fit_single_range_value(payload: dict[str, Any]) -> dict[str, Any]:
    """Shorten one string value by serialized size without changing page semantics."""
    values = payload.get("values")
    if (
        not isinstance(values, list)
        or len(values) != 1
        or not isinstance(values[0], list)
        or len(values[0]) != 1
        or not isinstance(values[0][0], str)
    ):
        raise ExcelLSPError(
            ErrorCode.INTERNAL, "A single-cell range response could not be serialized safely."
        )
    original = values[0][0]
    payload["valueTruncated"] = True
    low = 0
    high = len(original)
    while low < high:
        midpoint = (low + high + 1) // 2
        values[0][0] = original[:midpoint] + "…"
        if _serialized_length(payload) <= RESPONSE_CHARACTER_CAP:
            low = midpoint
        else:
            high = midpoint - 1
    values[0][0] = original[:low] + "…"
    if _serialized_length(payload) > RESPONSE_CHARACTER_CAP:
        raise ExcelLSPError(
            ErrorCode.INTERNAL, "A single-cell range response could not be serialized safely."
        )
    return payload


def _fit_payload(payload: dict[str, Any], *, list_keys: Sequence[str] = ()) -> dict[str, Any]:
    """Deterministically reduce bounded-list detail until the 8k contract holds."""
    if _serialized_length(payload) <= RESPONSE_CHARACTER_CAP:
        return payload
    payload["truncated"] = True
    for key in list_keys:
        value = payload.get(key)
        if not isinstance(value, list):
            continue
        while value and _serialized_length(payload) > RESPONSE_CHARACTER_CAP:
            if len(value) == 1:
                break
            del value[(len(value) + 1) // 2 :]
    if _serialized_length(payload) <= RESPONSE_CHARACTER_CAP:
        return payload
    _truncate_strings(payload, 256)
    if _serialized_length(payload) <= RESPONSE_CHARACTER_CAP:
        return payload
    _truncate_strings(payload, 80)
    if _serialized_length(payload) > RESPONSE_CHARACTER_CAP:
        for key in list_keys:
            value = payload.get(key)
            if isinstance(value, list):
                value.clear()
    if _serialized_length(payload) > RESPONSE_CHARACTER_CAP:
        raise ExcelLSPError(
            ErrorCode.INTERNAL, "A bounded tool response could not be serialized safely."
        )
    return payload


def fit_tool_envelope(payload: dict[str, Any]) -> dict[str, Any]:
    """Apply the serialized response cap to every transport envelope."""
    if "error" not in payload:
        return _fit_payload(payload)
    if _serialized_length(payload) <= RESPONSE_CHARACTER_CAP:
        return payload
    payload["truncated"] = True
    _truncate_strings(payload, 256)
    if _serialized_length(payload) <= RESPONSE_CHARACTER_CAP:
        return payload
    error = payload.get("error")
    if isinstance(error, dict) and "details" in error:
        error["details"] = {"truncated": True}
    _truncate_strings(payload, 80)
    if _serialized_length(payload) <= RESPONSE_CHARACTER_CAP:
        return payload
    safe_error: dict[str, Any] = {
        "code": _short(str(error.get("code", ErrorCode.INTERNAL.value)), 64)
        if isinstance(error, dict)
        else ErrorCode.INTERNAL.value,
        "message": _short(str(error.get("message", "Excel LSP returned an oversized error.")), 512)
        if isinstance(error, dict)
        else "Excel LSP returned an oversized error.",
    }
    if isinstance(error, dict) and error.get("hint"):
        safe_error["hint"] = _short(str(error["hint"]), 256)
    return {"error": safe_error, "truncated": True}


def _truncate_strings(value: Any, limit: int) -> None:
    if isinstance(value, dict):
        for key, item in tuple(value.items()):
            if isinstance(item, str):
                value[key] = _short(item, limit)
            else:
                _truncate_strings(item, limit)
    elif isinstance(value, list):
        for index, item in enumerate(value):
            if isinstance(item, str):
                value[index] = _short(item, limit)
            else:
                _truncate_strings(item, limit)


def _fit_trace_payload(payload: dict[str, Any]) -> dict[str, Any]:
    if _serialized_length(payload) <= RESPONSE_CHARACTER_CAP:
        return payload
    payload["truncated"] = True
    tree = payload["tree"]
    while _serialized_length(payload) > RESPONSE_CHARACTER_CAP:
        leaves = _trace_parents_with_children(tree)
        if not leaves:
            break
        parent = leaves[-1]
        parent["children"].pop()
    return _fit_payload(payload)


def _trace_parents_with_children(node: dict[str, Any]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for child in node.get("children", []):
        result.extend(_trace_parents_with_children(child))
    if node.get("children"):
        result.append(node)
    return result


__all__ = ["RAW_VALUE_CAP", "RESPONSE_CHARACTER_CAP", "ToolService", "fit_tool_envelope"]
