"""Focused full/incremental indexing and freshness lifecycle tests."""

from __future__ import annotations

import hashlib
import json
import os
from collections.abc import Callable
from pathlib import Path
from types import MappingProxyType, TracebackType
from typing import ClassVar, Self, cast

import pytest

import excel_lsp.core.index.lifecycle as lifecycle
from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.index import IndexStore, ensure_fresh, index_workbook, resolve_index_path
from excel_lsp.core.models import (
    CellRecord,
    DefinedName,
    NameArea,
    PackageHashes,
    Rect,
    SheetDescriptor,
    SheetKind,
    SheetParseSummary,
    WorkbookMetadata,
)
from excel_lsp.core.parse.styles import DEFAULT_STYLE_CATALOG


class FakeOOXMLParser:
    """Narrow parser contract fake; production imports the real OOXMLParser."""

    parse_calls: ClassVar[list[str]] = []
    styles = DEFAULT_STYLE_CATALOG

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self.document = cast(dict[str, object], json.loads(self.path.read_text(encoding="utf-8")))
        raw_sheets = cast(list[dict[str, object]], self.document["sheets"])
        descriptors = tuple(
            SheetDescriptor(
                order=order,
                name=cast(str, raw_sheet["name"]),
                sheet_id=cast(int, raw_sheet.get("sheet_id", order + 100)),
                rel_id=f"rId{order + 1}",
                xml_part=cast(str, raw_sheet["part"]),
                kind=cast(SheetKind, raw_sheet.get("kind", "worksheet")),
            )
            for order, raw_sheet in enumerate(raw_sheets)
        )
        self.metadata = WorkbookMetadata(
            path=str(self.path),
            date1904=cast(bool, self.document.get("date1904", False)),
            sheets=descriptors,
            defined_names=(
                DefinedName(
                    name="ScopedInput",
                    refers_to="Alpha!$A$1",
                    scope_sheet_order=0,
                    kind="range",
                    is_builtin=False,
                    areas=(NameArea("Alpha", Rect(1, 1, 1, 1)),),
                ),
            ),
            external_links=MappingProxyType({1: "../budget.xlsx"}),
            has_vba=True,
        )
        parts: dict[str, str] = {
            "xl/workbook.xml": _content_hash(self.document["workbook"]),
            "xl/_rels/workbook.xml.rels": _content_hash(self.document["rels"]),
            "xl/sharedStrings.xml": _content_hash(self.document["shared"]),
            "xl/styles.xml": _content_hash(self.document["styles"]),
        }
        for raw_sheet in raw_sheets:
            parts[cast(str, raw_sheet["part"])] = _content_hash(raw_sheet)
        self.hashes = PackageHashes(
            whole_file=hashlib.sha256(self.path.read_bytes()).hexdigest(),
            parts=parts,
        )

    def __enter__(self) -> Self:
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        del exc_type, exc_value, traceback

    def parse_sheet(
        self,
        sheet: SheetDescriptor,
        on_cell: Callable[[CellRecord], None],
    ) -> SheetParseSummary:
        self.parse_calls.append(sheet.name)
        raw_sheets = cast(list[dict[str, object]], self.document["sheets"])
        raw_sheet = next(item for item in raw_sheets if item["name"] == sheet.name)
        if sheet.kind != "worksheet":
            return SheetParseSummary(
                descriptor=sheet,
                part_hash=self.hashes.parts[sheet.xml_part],
                max_row=0,
                max_col=0,
                cell_count=0,
            )

        max_row = 0
        max_col = 0
        raw_cells = cast(list[dict[str, object]], raw_sheet.get("cells", []))
        for raw_cell in raw_cells:
            row = cast(int, raw_cell["row"])
            col = cast(int, raw_cell["col"])
            max_row = max(max_row, row)
            max_col = max(max_col, col)
            on_cell(
                CellRecord(
                    ref=cast(str, raw_cell["ref"]),
                    row=row,
                    col=col,
                    value=cast(int | float | str | bool | None, raw_cell.get("value")),
                    value_type=cast(str, raw_cell.get("value_type", "number")),  # type: ignore[arg-type]
                    formula=cast(str | None, raw_cell.get("formula")),
                )
            )
        return SheetParseSummary(
            descriptor=sheet,
            part_hash=self.hashes.parts[sheet.xml_part],
            max_row=max_row,
            max_col=max_col,
            cell_count=len(raw_cells),
        )


class TornSheetParser(FakeOOXMLParser):
    """Simulate a concurrent save tearing worksheet streaming."""

    failures_remaining: ClassVar[int] = 0
    failures_observed: ClassVar[int] = 0

    def parse_sheet(
        self,
        sheet: SheetDescriptor,
        on_cell: Callable[[CellRecord], None],
    ) -> SheetParseSummary:
        if sheet.name == "Alpha" and self.failures_remaining:
            type(self).failures_remaining -= 1
            type(self).failures_observed += 1
            self.document["writer_revision"] = self.failures_observed
            _write_document(self.path, self.document)
            raise ExcelLSPError(ErrorCode.CORRUPT, "worksheet stream was torn")
        return super().parse_sheet(sheet, on_cell)


@pytest.fixture(autouse=True)
def _fake_parser(monkeypatch: pytest.MonkeyPatch) -> None:
    FakeOOXMLParser.parse_calls.clear()
    monkeypatch.setattr(lifecycle, "OOXMLParser", FakeOOXMLParser)


def test_full_index_then_untouched_open_is_noop_and_source_is_unchanged(
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "model.xlsx"
    _write_document(workbook, _document())
    before_bytes = workbook.read_bytes()
    before_stat = workbook.stat()

    first = index_workbook(workbook, index_dir=tmp_path / "indexes")

    assert first.changed is True
    assert first.generation == 1
    assert first.reindexed_sheets == ("Alpha", "Beta", "Chart")
    assert workbook.read_bytes() == before_bytes
    after_stat = workbook.stat()
    assert (after_stat.st_mtime_ns, after_stat.st_size) == (
        before_stat.st_mtime_ns,
        before_stat.st_size,
    )
    with IndexStore(first.index_path) as store:
        assert [
            tuple(row)
            for row in store.connection.execute(
                "SELECT id, name, kind FROM sheets ORDER BY id"
            ).fetchall()
        ] == [
            (1, "Alpha", "worksheet"),
            (2, "Beta", "worksheet"),
            (3, "Chart", "chartsheet"),
        ]
        assert tuple(
            store.connection.execute(
                "SELECT value, formula FROM cells WHERE sheet_id = 2"
            ).fetchone()
        ) == (3, "=Alpha!A1+2")
        assert (
            store.connection.execute(
                "SELECT scope_sheet_id FROM defined_names WHERE name = 'ScopedInput'"
            ).fetchone()[0]
            == 1
        )
        assert store.get_meta("has_vba") == "1"
        assert json.loads(store.get_meta("external_links") or "{}") == {"1": "../budget.xlsx"}

    calls_after_full = tuple(FakeOOXMLParser.parse_calls)
    second = ensure_fresh(workbook, index_dir=tmp_path / "indexes")
    assert second.changed is False
    assert second.generation == first.generation
    assert second.reindexed_sheets == ()
    assert tuple(FakeOOXMLParser.parse_calls) == calls_after_full


def test_one_sheet_change_is_incremental_and_generation_is_monotonic(tmp_path: Path) -> None:
    workbook = tmp_path / "incremental.xlsx"
    document = _document()
    _write_document(workbook, document)
    first = index_workbook(workbook, index_dir=tmp_path / "indexes")
    FakeOOXMLParser.parse_calls.clear()

    alpha = cast(list[dict[str, object]], document["sheets"])[0]
    cast(list[dict[str, object]], alpha["cells"])[0]["value"] = 100
    _write_document(workbook, document)
    second = index_workbook(workbook, index_dir=tmp_path / "indexes")

    assert second.reindexed_sheets == ("Alpha",)
    assert FakeOOXMLParser.parse_calls == ["Alpha"]
    assert second.generation == first.generation + 1
    with IndexStore(second.index_path) as store:
        assert (
            store.connection.execute(
                "SELECT value FROM cells WHERE sheet_id = 1 AND ref = 'A1'"
            ).fetchone()[0]
            == 100
        )
        assert (
            store.connection.execute(
                "SELECT value FROM cells WHERE sheet_id = 2 AND ref = 'A1'"
            ).fetchone()[0]
            == 3
        )


@pytest.mark.parametrize("changed_key", ["workbook", "rels", "shared", "styles"])
def test_global_package_parts_invalidate_every_sheet(
    tmp_path: Path,
    changed_key: str,
) -> None:
    workbook = tmp_path / f"global-{changed_key}.xlsx"
    document = _document()
    _write_document(workbook, document)
    index_workbook(workbook, index_dir=tmp_path / "indexes")
    FakeOOXMLParser.parse_calls.clear()

    document[changed_key] = f"{document[changed_key]}-changed"
    _write_document(workbook, document)
    update = index_workbook(workbook, index_dir=tmp_path / "indexes")

    assert update.reindexed_sheets == ("Alpha", "Beta", "Chart")
    assert FakeOOXMLParser.parse_calls == ["Alpha", "Beta", "Chart"]


def test_incremental_and_fresh_full_indexes_have_equal_canonical_exports(
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "canonical.xlsx"
    document = _document()
    _write_document(workbook, document)
    index_workbook(workbook, index_dir=tmp_path / "incremental-index")

    beta = cast(list[dict[str, object]], document["sheets"])[1]
    cast(list[dict[str, object]], beta["cells"])[0]["formula"] = "=Alpha!A1+5"
    cast(list[dict[str, object]], beta["cells"])[0]["value"] = 6
    _write_document(workbook, document)
    incremental = index_workbook(workbook, index_dir=tmp_path / "incremental-index")
    full = index_workbook(workbook, index_dir=tmp_path / "full-index")

    with IndexStore(incremental.index_path) as incremental_store:
        incremental_export = incremental_store.canonical_export()
        incremental_source_key = (
            "edge_source_rtree"
            if incremental_store.edge_store.backend == "rtree"
            else "edge_source_intervals"
        )
        incremental_trust = incremental_store.connection.execute(
            """
            SELECT dirty, mutation_epoch, clean_epoch
            FROM graph_spatial_state WHERE singleton = 1
            """
        ).fetchone()
    with IndexStore(full.index_path) as full_store:
        full_export = full_store.canonical_export()
        full_source_key = (
            "edge_source_rtree"
            if full_store.edge_store.backend == "rtree"
            else "edge_source_intervals"
        )

    assert incremental_source_key == full_source_key
    assert incremental_export["edge_ranks"]
    assert incremental_export[incremental_source_key]
    assert incremental_export["graph_spatial_state"][0][0] == 0
    assert incremental_trust is not None
    assert int(incremental_trust[0]) == 0
    assert int(incremental_trust[1]) == int(incremental_trust[2])
    assert incremental_export == full_export


def test_schema_version_rebuild_forces_full_index_even_when_stat_is_fresh(
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "schema.xlsx"
    _write_document(workbook, _document())
    first = index_workbook(workbook, index_dir=tmp_path / "indexes")
    with IndexStore(first.index_path) as store:
        store.set_meta("schema_version", "obsolete")
    FakeOOXMLParser.parse_calls.clear()

    rebuilt = index_workbook(workbook, index_dir=tmp_path / "indexes")

    assert rebuilt.reindexed_sheets == ("Alpha", "Beta", "Chart")
    assert rebuilt.generation > first.generation
    assert FakeOOXMLParser.parse_calls == ["Alpha", "Beta", "Chart"]


def test_missing_graph_analysis_version_forces_full_semantic_refresh(
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "graph-version.xlsx"
    _write_document(workbook, _document())
    first = index_workbook(workbook, index_dir=tmp_path / "indexes")
    with IndexStore(first.index_path) as store:
        store.connection.execute("DELETE FROM meta WHERE key = 'graph_analysis_version'")
    FakeOOXMLParser.parse_calls.clear()

    refreshed = ensure_fresh(workbook, index_dir=tmp_path / "indexes")

    assert refreshed.changed is True
    assert refreshed.generation == first.generation + 1
    assert refreshed.reindexed_sheets == ("Alpha", "Beta", "Chart")
    assert FakeOOXMLParser.parse_calls == ["Alpha", "Beta", "Chart"]
    with IndexStore(refreshed.index_path) as store:
        assert store.get_meta("graph_analysis_version") == "1"


@pytest.mark.parametrize(
    ("stored_gap_tol", "stored_path"),
    (
        (None, None),
        ("invalid", None),
        ("-1", None),
        ("0", "different-workbook.xlsx"),
    ),
)
def test_omitted_gap_tolerance_defaults_when_stored_configuration_is_invalid(
    tmp_path: Path,
    stored_gap_tol: str | None,
    stored_path: str | None,
) -> None:
    workbook = tmp_path / "stored-gap.xlsx"
    index_dir = tmp_path / "indexes"
    _write_document(workbook, _document())
    first = index_workbook(workbook, index_dir=index_dir, gap_tol=0)
    with IndexStore(first.index_path) as store:
        if stored_gap_tol is None:
            store.connection.execute("DELETE FROM meta WHERE key = 'region_gap_tol'")
        else:
            store.set_meta("region_gap_tol", stored_gap_tol)
        if stored_path is not None:
            store.set_meta("workbook_path", str(tmp_path / stored_path))

    refreshed = ensure_fresh(workbook, index_dir=index_dir)

    assert refreshed.changed is True
    assert refreshed.generation == first.generation + 1
    with IndexStore(refreshed.index_path) as store:
        assert store.get_meta("workbook_path") == str(workbook.resolve())
        assert store.get_meta("region_gap_tol") == "1"


def test_negative_gap_tolerance_is_rejected_before_index_creation(tmp_path: Path) -> None:
    workbook = tmp_path / "negative-gap.xlsx"
    index_dir = tmp_path / "indexes"
    _write_document(workbook, _document())
    index_path = resolve_index_path(workbook, index_dir)

    with pytest.raises(ValueError, match="nonnegative"):
        index_workbook(workbook, index_dir=index_dir, gap_tol=-1)

    assert not index_path.exists()


def test_excessive_gap_tolerance_is_rejected_before_index_creation(tmp_path: Path) -> None:
    workbook = tmp_path / "excessive-gap.xlsx"
    index_dir = tmp_path / "indexes"
    _write_document(workbook, _document())
    index_path = resolve_index_path(workbook, index_dir)

    with pytest.raises(ValueError, match="must not exceed 8"):
        index_workbook(workbook, index_dir=index_dir, gap_tol=9)

    assert not index_path.exists()


@pytest.mark.parametrize("gap_tol", (True, False, 1.5, 0.0))
def test_noninteger_gap_tolerance_is_rejected_before_index_creation(
    tmp_path: Path,
    gap_tol: object,
) -> None:
    workbook = tmp_path / "noninteger-gap.xlsx"
    index_dir = tmp_path / "indexes"
    _write_document(workbook, _document())
    index_path = resolve_index_path(workbook, index_dir)

    with pytest.raises(ValueError, match="gap_tol must be an integer"):
        index_workbook(
            workbook,
            index_dir=index_dir,
            gap_tol=cast(int, gap_tol),
        )

    assert not index_path.exists()


def test_parse_time_torn_save_retries_once_after_the_source_stat_changes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = tmp_path / "torn-once.xlsx"
    _write_document(workbook, _document())
    TornSheetParser.failures_remaining = 1
    TornSheetParser.failures_observed = 0
    monkeypatch.setattr(lifecycle, "OOXMLParser", TornSheetParser)

    update = index_workbook(workbook, index_dir=tmp_path / "indexes")

    assert TornSheetParser.failures_observed == 1
    assert update.changed is True
    assert update.reindexed_sheets == ("Alpha", "Beta", "Chart")
    with IndexStore(update.index_path) as store:
        assert store.connection.execute("SELECT COUNT(*) FROM cells").fetchone()[0] == 3


def test_hash_equal_refresh_retries_if_mtime_changes_during_hashing(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = tmp_path / "mtime-during-hash.xlsx"
    _write_document(workbook, _document())
    first = index_workbook(workbook, index_dir=tmp_path / "indexes")
    before_touch = workbook.stat()
    os.utime(
        workbook,
        ns=(before_touch.st_atime_ns, before_touch.st_mtime_ns + 2_000_000_000),
    )

    class MtimeTouchParser(FakeOOXMLParser):
        attempts = 0

        def __init__(self, path: str | Path) -> None:
            super().__init__(path)
            type(self).attempts += 1
            if type(self).attempts == 1:
                current = self.path.stat()
                os.utime(
                    self.path,
                    ns=(current.st_atime_ns, current.st_mtime_ns + 2_000_000_000),
                )

    monkeypatch.setattr(lifecycle, "OOXMLParser", MtimeTouchParser)

    refreshed = index_workbook(workbook, index_dir=tmp_path / "indexes")
    final_stat = workbook.stat()

    assert MtimeTouchParser.attempts == 2
    assert refreshed.changed is False
    assert refreshed.generation == first.generation
    assert refreshed.reindexed_sheets == ()
    with IndexStore(refreshed.index_path) as store:
        assert store.generation == first.generation
        assert store.get_meta("mtime_ns") == str(final_stat.st_mtime_ns)
        assert store.get_meta("size") == str(final_stat.st_size)


def test_repeated_parse_time_torn_save_preserves_the_existing_index(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = tmp_path / "torn-twice.xlsx"
    document = _document()
    _write_document(workbook, document)
    initial = index_workbook(workbook, index_dir=tmp_path / "indexes")
    with IndexStore(initial.index_path) as store:
        before_export = store.canonical_export()
        before_generation = store.generation

    alpha = cast(list[dict[str, object]], document["sheets"])[0]
    cast(list[dict[str, object]], alpha["cells"])[0]["value"] = 99
    _write_document(workbook, document)
    TornSheetParser.failures_remaining = 2
    TornSheetParser.failures_observed = 0
    monkeypatch.setattr(lifecycle, "OOXMLParser", TornSheetParser)

    with pytest.raises(ExcelLSPError) as caught:
        index_workbook(workbook, index_dir=tmp_path / "indexes")

    assert caught.value.code is ErrorCode.CORRUPT
    assert TornSheetParser.failures_observed == 2
    with IndexStore(initial.index_path) as store:
        assert store.generation == before_generation
        assert store.canonical_export() == before_export


def test_path_resolution_defaults_next_to_file_and_hashes_central_locations(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    first = tmp_path / "one" / "same.xlsx"
    second = tmp_path / "two" / "same.xlsx"
    central = tmp_path / "central"

    assert resolve_index_path(first) == first.resolve().with_name("same.xlsx.xlsp.db")
    first_central = resolve_index_path(first, central)
    assert first_central.parent == central.resolve()
    expected_hash = hashlib.sha256(str(first.resolve()).encode("utf-8")).hexdigest()[:8]
    assert first_central.name == f"same.{expected_hash}.xlsp.db"
    assert first_central == resolve_index_path(first, central)
    assert first_central != resolve_index_path(second, central)

    monkeypatch.setenv("EXCEL_LSP_INDEX_DIR", str(central))
    assert resolve_index_path(first) == first_central


def _document() -> dict[str, object]:
    return {
        "workbook": "workbook-v1",
        "rels": "rels-v1",
        "shared": "shared-v1",
        "styles": "styles-v1",
        "date1904": False,
        "irrelevant": "ignored-by-part-selection",
        "sheets": [
            {
                "name": "Alpha",
                "part": "xl/worksheets/sheet1.xml",
                "cells": [
                    {"ref": "A1", "row": 1, "col": 1, "value": 1},
                    {"ref": "B1", "row": 1, "col": 2, "value": "north", "value_type": "string"},
                ],
            },
            {
                "name": "Beta",
                "part": "xl/worksheets/sheet2.xml",
                "cells": [
                    {
                        "ref": "A1",
                        "row": 1,
                        "col": 1,
                        "value": 3,
                        "formula": "=Alpha!A1+2",
                    }
                ],
            },
            {
                "name": "Chart",
                "part": "xl/chartsheets/sheet3.xml",
                "kind": "chartsheet",
                "cells": [],
            },
        ],
    }


def _write_document(path: Path, document: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    previous_mtime = path.stat().st_mtime_ns if path.exists() else None
    path.write_text(json.dumps(document, sort_keys=True), encoding="utf-8")
    if previous_mtime is not None and path.stat().st_mtime_ns <= previous_mtime:
        next_mtime = previous_mtime + 1_000_000
        os.utime(path, ns=(next_mtime, next_mtime))


def _content_hash(value: object) -> str:
    payload = json.dumps(value, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()
