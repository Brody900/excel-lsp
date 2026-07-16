"""Compact, deterministic workbook-map projection over the semantic index."""

from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from urllib.parse import unquote, urlsplit

from excel_lsp.core.index import IndexStore, ensure_fresh
from excel_lsp.core.parse.coordinates import make_cell_ref
from excel_lsp.core.symbols import defined_name_symbol_id, region_symbol_id

MAP_CHARACTER_CAP = 8_000
MAP_REGION_LIMIT = 8
MAP_NAME_LIMIT = 20
MAP_EXTERNAL_LINK_LIMIT = 10
# These source ceilings are proofs from the frozen 8,000-character response cap,
# not extra public degradation rules. Even with one-character fields, the compact
# JSON envelope cannot contain 201 sheet items, 81 region items, or 513 column
# items. Round-robin regions plus a 16-column base preserve every smaller render
# tier; the extra-column allowance preserves any undegraded map that could fit.
_SHEET_SOURCE_LIMIT = 200
_REGION_SOURCE_LIMIT = 80
_BASE_COLUMN_SOURCE_LIMIT = 16
_EXTRA_COLUMN_SOURCE_LIMIT = 512
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

_HINTS = (
    "Use get_region_schema for columns+samples",
    "trace_dependents for impact analysis",
)


@dataclass(frozen=True, slots=True)
class _RegionSource:
    n: int
    range_ref: str
    kind: str
    header_rows: int
    columns: tuple[tuple[str, str], ...]
    column_count: int
    rows: int
    formula_blocks: int
    confidence: float


@dataclass(frozen=True, slots=True)
class _SheetSource:
    name: str
    visibility: str
    dimensions: str
    regions: tuple[_RegionSource, ...]
    region_count: int


@dataclass(frozen=True, slots=True)
class _NameSource:
    symbol_id: str
    refers_to: str


@dataclass(frozen=True, slots=True)
class _MapSource:
    workbook: str
    indexed_at: str
    has_vba: bool
    sheets: tuple[_SheetSource, ...]
    sheet_count: int
    visibility_counts: tuple[tuple[str, int], ...]
    names: tuple[_NameSource, ...]
    name_count: int
    external_links: tuple[str, ...]
    external_link_count: int
    diagnostic_errors: int
    diagnostic_warnings: int


@dataclass(frozen=True, slots=True)
class _RenderLimits:
    regions: int
    columns: int | None
    names: int
    external_links: int
    detailed_sheets: int | None = None


def build_workbook_map(
    path: str | Path,
    *,
    index_dir: str | Path | None = None,
    character_cap: int = MAP_CHARACTER_CAP,
) -> dict[str, object]:
    """Refresh ``path`` and return its bounded semantic map."""
    if character_cap < 1:
        raise ValueError("character_cap must be positive")
    if character_cap > MAP_CHARACTER_CAP:
        raise ValueError(f"character_cap cannot exceed {MAP_CHARACTER_CAP}")
    workbook = Path(path).expanduser().resolve()
    update = ensure_fresh(workbook, index_dir=index_dir)
    with IndexStore(update.index_path) as store:
        source = _load_source(workbook, store)
    return _bounded_render(source, character_cap=character_cap)


def serialize_workbook_map(workbook_map: dict[str, object]) -> str:
    """Serialize a map exactly as response-cap checks measure it."""
    return json.dumps(workbook_map, ensure_ascii=False, separators=(",", ":"))


def _load_source(
    workbook: Path,
    store: IndexStore,
) -> _MapSource:
    connection = store.connection

    meta = {
        str(row["key"]): str(row["value"])
        for row in connection.execute(
            """
            SELECT key, value FROM meta
            WHERE key IN ('indexed_at', 'has_vba')
            """
        ).fetchall()
    }
    stats = connection.execute(
        """
        SELECT
            (SELECT COUNT(*) FROM sheets) AS sheet_count,
            (SELECT COUNT(*) FROM sheets WHERE visibility = 'visible') AS visible_count,
            (SELECT COUNT(*) FROM sheets WHERE visibility = 'hidden') AS hidden_count,
            (SELECT COUNT(*) FROM sheets WHERE visibility = 'veryHidden') AS very_hidden_count,
            (SELECT COUNT(*) FROM defined_names) AS name_count,
            (SELECT COUNT(*) FROM diagnostics WHERE severity = 'error') AS diagnostic_errors,
            (SELECT COUNT(*) FROM diagnostics WHERE severity = 'warn') AS diagnostic_warnings
        """
    ).fetchone()
    assert stats is not None

    sheet_rows = connection.execute(
        """
        WITH region_counts AS (
            SELECT sheet_id, COUNT(*) AS region_count
            FROM regions GROUP BY sheet_id
        ),
        ranked AS (
            SELECT
                s.id, s.name, s.visibility, s.max_row, s.max_col,
                COALESCE(region_counts.region_count, 0) AS region_count,
                ROW_NUMBER() OVER (
                    ORDER BY CASE WHEN s.visibility = 'visible' THEN 1 ELSE 0 END, s.id
                ) AS detail_rank
            FROM sheets AS s
            LEFT JOIN region_counts ON region_counts.sheet_id = s.id
        )
        SELECT id, name, visibility, max_row, max_col, region_count
        FROM ranked
        WHERE detail_rank <= ?
        ORDER BY id
        """,
        (_SHEET_SOURCE_LIMIT,),
    ).fetchall()

    region_rows = connection.execute(
        """
        WITH candidate_sheets AS (
            SELECT id
            FROM (
                SELECT
                    id,
                    ROW_NUMBER() OVER (
                        ORDER BY CASE WHEN visibility = 'visible' THEN 1 ELSE 0 END, id
                    ) AS detail_rank
                FROM sheets
            )
            WHERE detail_rank <= ?
        ),
        ranked_regions AS (
            SELECT
                r.id, r.sheet_id, r.n, r.row_min, r.row_max, r.col_min, r.col_max,
                r.header_rows, r.kind, r.confidence,
                ROW_NUMBER() OVER (
                    PARTITION BY r.sheet_id
                    ORDER BY
                        ((r.row_max - r.row_min + 1) * (r.col_max - r.col_min + 1)) DESC,
                        r.row_min, r.col_min, r.n
                ) AS render_rank
            FROM regions AS r
            JOIN candidate_sheets ON candidate_sheets.id = r.sheet_id
        ),
        prioritized_regions AS (
            SELECT
                *,
                ROW_NUMBER() OVER (ORDER BY render_rank, sheet_id) AS source_rank
            FROM ranked_regions
            WHERE render_rank <= ?
        ),
        selected_regions AS (
            SELECT *
            FROM prioritized_regions
            WHERE source_rank <= ?
        ),
        column_counts AS (
            SELECT c.region_id, COUNT(*) AS column_count
            FROM columns AS c
            JOIN selected_regions AS selected ON selected.id = c.region_id
            GROUP BY c.region_id
        ),
        fblock_counts AS (
            SELECT selected.id AS region_id, COUNT(f.id) AS formula_blocks
            FROM selected_regions AS selected
            LEFT JOIN fblocks AS f
              ON f.sheet_id = selected.sheet_id
             AND f.row_max >= selected.row_min AND f.row_min <= selected.row_max
             AND f.col_max >= selected.col_min AND f.col_min <= selected.col_max
            GROUP BY selected.id
        )
        SELECT
            selected.id, selected.sheet_id, selected.n,
            selected.row_min, selected.row_max, selected.col_min, selected.col_max,
            selected.header_rows, selected.kind, selected.confidence,
            COALESCE(column_counts.column_count, 0) AS column_count,
            COALESCE(fblock_counts.formula_blocks, 0) AS formula_blocks
        FROM selected_regions AS selected
        LEFT JOIN column_counts ON column_counts.region_id = selected.id
        LEFT JOIN fblock_counts ON fblock_counts.region_id = selected.id
        ORDER BY selected.sheet_id, selected.render_rank
        """,
        (_SHEET_SOURCE_LIMIT, MAP_REGION_LIMIT, _REGION_SOURCE_LIMIT),
    ).fetchall()

    column_rows = connection.execute(
        """
        WITH candidate_sheets AS (
            SELECT id
            FROM (
                SELECT
                    id,
                    ROW_NUMBER() OVER (
                        ORDER BY CASE WHEN visibility = 'visible' THEN 1 ELSE 0 END, id
                    ) AS detail_rank
                FROM sheets
            )
            WHERE detail_rank <= ?
        ),
        ranked_regions AS (
            SELECT
                r.id, r.sheet_id,
                ROW_NUMBER() OVER (
                    PARTITION BY r.sheet_id
                    ORDER BY
                        ((r.row_max - r.row_min + 1) * (r.col_max - r.col_min + 1)) DESC,
                        r.row_min, r.col_min, r.n
                ) AS render_rank
            FROM regions AS r
            JOIN candidate_sheets ON candidate_sheets.id = r.sheet_id
        ),
        prioritized_regions AS (
            SELECT
                *,
                ROW_NUMBER() OVER (ORDER BY render_rank, sheet_id) AS source_rank
            FROM ranked_regions
            WHERE render_rank <= ?
        ),
        selected_regions AS (
            SELECT id, sheet_id, render_rank
            FROM prioritized_regions
            WHERE source_rank <= ?
        ),
        ranked_columns AS (
            SELECT
                c.region_id, selected.sheet_id, selected.render_rank,
                c.idx, c.header, c.dtype,
                ROW_NUMBER() OVER (
                    PARTITION BY c.region_id ORDER BY c.idx
                ) AS column_rank
            FROM columns AS c
            JOIN selected_regions AS selected ON selected.id = c.region_id
        ),
        extra_columns AS (
            SELECT region_id, idx
            FROM ranked_columns
            WHERE column_rank > ?
            ORDER BY sheet_id, render_rank, column_rank
            LIMIT ?
        )
        SELECT c.region_id, c.idx, c.header, c.dtype
        FROM ranked_columns AS c
        LEFT JOIN extra_columns AS extra
          ON extra.region_id = c.region_id AND extra.idx = c.idx
        WHERE c.column_rank <= ? OR extra.region_id IS NOT NULL
        ORDER BY c.region_id, c.idx
        """,
        (
            _SHEET_SOURCE_LIMIT,
            MAP_REGION_LIMIT,
            _REGION_SOURCE_LIMIT,
            _BASE_COLUMN_SOURCE_LIMIT,
            _EXTRA_COLUMN_SOURCE_LIMIT,
            _BASE_COLUMN_SOURCE_LIMIT,
        ),
    ).fetchall()

    name_rows = connection.execute(
        """
        SELECT d.name, scope.name AS scope_name, d.refers_to
        FROM defined_names AS d
        LEFT JOIN sheets AS scope ON scope.id = d.scope_sheet_id
        ORDER BY COALESCE(scope.name, ''), d.name, d.refers_to
        LIMIT ?
        """,
        (MAP_NAME_LIMIT,),
    ).fetchall()

    link_rows = connection.execute(
        """
        WITH stored AS (
            SELECT COALESCE(
                (SELECT value FROM meta WHERE key = 'external_links'),
                '{}'
            ) AS payload
        ),
        valid_object AS (
            SELECT
                CASE
                    WHEN json_valid(payload)
                    THEN CASE
                        WHEN json_type(payload) = 'object' THEN payload
                        ELSE '{}'
                    END
                    ELSE '{}'
                END AS payload
            FROM stored
        ),
        raw_links AS (
            SELECT CAST(link.key AS INTEGER) AS link_index, CAST(link.value AS TEXT) AS target
            FROM valid_object
            JOIN json_each(valid_object.payload) AS link
            WHERE link.type = 'text'
              AND CAST(link.key AS INTEGER) > 0
              AND CAST(CAST(link.key AS INTEGER) AS TEXT) = link.key
        ),
        ranked AS (
            SELECT
                link_index,
                target,
                ROW_NUMBER() OVER (ORDER BY link_index) AS render_rank,
                COUNT(*) OVER () AS total
            FROM raw_links
        )
        SELECT target, total, render_rank
        FROM ranked
        WHERE render_rank <= ?
        UNION ALL
        SELECT NULL, 0, 0
        WHERE NOT EXISTS (SELECT 1 FROM raw_links)
        ORDER BY render_rank
        """,
        (MAP_EXTERNAL_LINK_LIMIT,),
    ).fetchall()

    columns_by_region: defaultdict[int, list[tuple[str, str]]] = defaultdict(list)
    for row in column_rows:
        columns_by_region[int(row["region_id"])].append((str(row["header"]), str(row["dtype"])))
    regions_by_sheet: defaultdict[int, list[_RegionSource]] = defaultdict(list)
    for row in region_rows:
        region_id = int(row["id"])
        row_min = int(row["row_min"])
        row_max = int(row["row_max"])
        col_min = int(row["col_min"])
        col_max = int(row["col_max"])
        header_rows = int(row["header_rows"])
        regions_by_sheet[int(row["sheet_id"])].append(
            _RegionSource(
                n=int(row["n"]),
                range_ref=_range_ref(row_min, row_max, col_min, col_max),
                kind=str(row["kind"]),
                header_rows=header_rows,
                columns=tuple(columns_by_region[region_id]),
                column_count=int(row["column_count"]),
                rows=max(0, row_max - row_min + 1 - header_rows),
                formula_blocks=int(row["formula_blocks"]),
                confidence=float(row["confidence"]),
            )
        )
    sheets = tuple(
        _SheetSource(
            name=str(row["name"]),
            visibility=str(row["visibility"]),
            dimensions=(
                _range_ref(1, int(row["max_row"]), 1, int(row["max_col"]))
                if int(row["max_row"]) and int(row["max_col"])
                else "A1"
            ),
            regions=tuple(regions_by_sheet[int(row["id"])]),
            region_count=int(row["region_count"]),
        )
        for row in sheet_rows
    )
    names = tuple(
        _NameSource(
            symbol_id=defined_name_symbol_id(
                str(row["name"]),
                scope_sheet=None if row["scope_name"] is None else str(row["scope_name"]),
            ),
            refers_to=str(row["refers_to"]),
        )
        for row in name_rows
    )
    external_links = tuple(
        _external_link_label(str(row["target"])) for row in link_rows if row["target"] is not None
    )
    external_link_count = int(link_rows[0]["total"]) if link_rows else 0
    return _MapSource(
        workbook=workbook.name,
        indexed_at=meta.get("indexed_at", ""),
        has_vba=meta.get("has_vba", "0") == "1",
        sheets=sheets,
        sheet_count=int(stats["sheet_count"]),
        visibility_counts=(
            ("visible", int(stats["visible_count"])),
            ("hidden", int(stats["hidden_count"])),
            ("veryHidden", int(stats["very_hidden_count"])),
        ),
        names=names,
        name_count=int(stats["name_count"]),
        external_links=external_links,
        external_link_count=external_link_count,
        diagnostic_errors=int(stats["diagnostic_errors"]),
        diagnostic_warnings=int(stats["diagnostic_warnings"]),
    )


def _bounded_render(source: _MapSource, *, character_cap: int) -> dict[str, object]:
    attempts = (
        _RenderLimits(MAP_REGION_LIMIT, None, MAP_NAME_LIMIT, MAP_EXTERNAL_LINK_LIMIT),
        _RenderLimits(8, 16, 20, 10),
        _RenderLimits(8, 8, 20, 10),
        _RenderLimits(4, 8, 20, 10),
        _RenderLimits(2, 6, 20, 10),
        _RenderLimits(1, 4, 10, 5),
        _RenderLimits(0, 0, 5, 3),
        _RenderLimits(0, 0, 0, 0, 40),
        _RenderLimits(0, 0, 0, 0, 20),
        _RenderLimits(0, 0, 0, 0, 10),
        _RenderLimits(0, 0, 0, 0, 0),
    )
    for limits in attempts:
        if not _source_is_complete_for(source, limits):
            continue
        rendered = _render(source, limits)
        if len(serialize_workbook_map(rendered)) <= character_cap:
            return rendered
    raise ValueError("character cap is too small for the minimal workbook map")


def _render(source: _MapSource, limits: _RenderLimits) -> dict[str, object]:
    detailed_indices = _detailed_sheet_indices(source.sheets, limits.detailed_sheets)
    sheet_list: list[dict[str, object]] = []
    for index, sheet in enumerate(source.sheets):
        if index not in detailed_indices:
            continue
        rendered_sheet: dict[str, object] = {
            "sheet": sheet.name,
            "dims": sheet.dimensions,
            "regions": _render_regions(sheet, limits),
        }
        if sheet.visibility != "visible":
            rendered_sheet["vis"] = sheet.visibility
        sheet_list.append(rendered_sheet)

    names = [{"id": item.symbol_id, "ref": item.refers_to} for item in source.names[: limits.names]]
    links = list(source.external_links[: limits.external_links])
    result: dict[str, object] = {
        "workbook": source.workbook,
        "sheets": source.sheet_count,
        "indexed_at": source.indexed_at,
        "stale": False,
        "hasVBA": source.has_vba,
        "sheetList": sheet_list,
        "names": names,
        "namesMore": source.name_count - len(names),
        "externalLinks": links,
        "externalLinksMore": source.external_link_count - len(links),
        "diagCounts": {
            "error": source.diagnostic_errors,
            "warn": source.diagnostic_warnings,
        },
        "hints": list(_HINTS),
    }
    omitted_sheets = source.sheet_count - len(sheet_list)
    if omitted_sheets:
        result["sheetListMore"] = omitted_sheets
        rendered_visibility = Counter(
            sheet.visibility
            for index, sheet in enumerate(source.sheets)
            if index in detailed_indices
        )
        omitted_by_visibility = {
            visibility: total - rendered_visibility[visibility]
            for visibility, total in source.visibility_counts
            if total > rendered_visibility[visibility]
        }
        result["sheetListMoreByVis"] = omitted_by_visibility
    return result


def _render_regions(sheet: _SheetSource, limits: _RenderLimits) -> list[dict[str, object]]:
    rendered: list[dict[str, object]] = []
    for region in sheet.regions[: limits.regions]:
        columns = region.columns
        if limits.columns is not None:
            columns = columns[: limits.columns]
        item: dict[str, object] = {
            "id": region_symbol_id(sheet.name, region.n),
            "range": region.range_ref,
            "kind": region.kind,
            "headerRows": region.header_rows,
            "cols": [{"h": header, "t": dtype} for header, dtype in columns],
            "rows": region.rows,
            "fblocks": region.formula_blocks,
            "conf": round(region.confidence, 2),
        }
        omitted_columns = region.column_count - len(columns)
        if omitted_columns:
            item["colsMore"] = omitted_columns
        rendered.append(item)
    omitted_regions = sheet.region_count - len(rendered)
    if omitted_regions:
        rendered.append({"more": omitted_regions})
    return rendered


def _source_is_complete_for(source: _MapSource, limits: _RenderLimits) -> bool:
    if limits.detailed_sheets is None and len(source.sheets) != source.sheet_count:
        return False
    detailed_indices = _detailed_sheet_indices(source.sheets, limits.detailed_sheets)
    for index, sheet in enumerate(source.sheets):
        if index not in detailed_indices:
            continue
        required_regions = min(sheet.region_count, limits.regions)
        if len(sheet.regions) < required_regions:
            return False
        for region in sheet.regions[:required_regions]:
            required_columns = (
                region.column_count
                if limits.columns is None
                else min(region.column_count, limits.columns)
            )
            if len(region.columns) < required_columns:
                return False
    return True


def _detailed_sheet_indices(
    sheets: tuple[_SheetSource, ...],
    limit: int | None,
) -> frozenset[int]:
    if limit is None or limit >= len(sheets):
        return frozenset(range(len(sheets)))
    prioritized = sorted(
        range(len(sheets)),
        key=lambda index: (sheets[index].visibility == "visible", index),
    )
    return frozenset(prioritized[:limit])


def _range_ref(row_min: int, row_max: int, col_min: int, col_max: int) -> str:
    start = make_cell_ref(row_min, col_min)
    end = make_cell_ref(row_max, col_max)
    return start if start == end else f"{start}:{end}"


def _external_link_label(target: str) -> str:
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


__all__ = [
    "MAP_CHARACTER_CAP",
    "MAP_EXTERNAL_LINK_LIMIT",
    "MAP_NAME_LIMIT",
    "MAP_REGION_LIMIT",
    "build_workbook_map",
    "serialize_workbook_map",
]
