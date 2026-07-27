"""Formula-analysis persistence and lifecycle integration regressions."""

from __future__ import annotations

import json
import runpy
from collections.abc import Callable, Mapping, Sequence
from pathlib import Path
from types import MappingProxyType
from typing import cast
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

import excel_lsp.core.index.store as store_module
from excel_lsp.core.formulas.indexing import SheetFormulaAnalysis
from excel_lsp.core.index import IndexStore, index_workbook
from excel_lsp.core.models import (
    CellRecord,
    DefinedName,
    SheetDescriptor,
    SheetParseSummary,
    TableInfo,
    WorkbookMetadata,
)

GenerateAll = Callable[[Path], dict[str, Path]]
generate_all = cast(
    GenerateAll,
    runpy.run_path(str(Path(__file__).parents[1] / "fixtures" / "generate.py"))["generate_all"],
)

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCUMENT_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_TYPE_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@pytest.fixture(scope="module")
def formula_fixtures(tmp_path_factory: pytest.TempPathFactory) -> dict[str, Path]:
    return generate_all(tmp_path_factory.mktemp("formula-index-fixtures"))


def test_f07_persists_exact_blocks_edges_and_single_inconsistency(
    tmp_path: Path,
    formula_fixtures: dict[str, Path],
) -> None:
    update = index_workbook(formula_fixtures["F07"], index_dir=tmp_path / "indexes")

    with IndexStore(update.index_path) as store:
        exported = store.canonical_export()

    assert exported["fblocks"] == (
        ("FormulaBlocks", 0, "=RC[-2]*RC[-1]", 2, 11, 3, 3, 0, 0),
        ("FormulaBlocks", 1, "=RC[-2]+RC[-1]", 12, 12, 3, 3, 0, 0),
        ("FormulaBlocks", 2, "=RC[-2]*RC[-1]", 13, 21, 3, 3, 0, 0),
    )
    assert exported["edges"] == (
        ("fblock", "FormulaBlocks", 0, "FormulaBlocks", 2, 11, 1, 1, "ref"),
        ("fblock", "FormulaBlocks", 0, "FormulaBlocks", 2, 11, 2, 2, "ref"),
        ("fblock", "FormulaBlocks", 1, "FormulaBlocks", 12, 12, 1, 1, "ref"),
        ("fblock", "FormulaBlocks", 1, "FormulaBlocks", 12, 12, 2, 2, "ref"),
        ("fblock", "FormulaBlocks", 2, "FormulaBlocks", 13, 21, 1, 1, "ref"),
        ("fblock", "FormulaBlocks", 2, "FormulaBlocks", 13, 21, 2, 2, "ref"),
    )
    assert exported["diagnostics"] == (
        (
            "warn",
            "W_INCONSISTENT_FORMULA",
            "FormulaBlocks",
            12,
            3,
            "cell:FormulaBlocks!C12",
            "Formula differs from the dominant pattern in its contiguous run.",
            '{"dominantBlock":"fblock:FormulaBlocks:0","expectedR1C1":"=RC[-2]*RC[-1]"}',
        ),
    )


def test_f19_persists_modern_edges_without_spurious_diagnostics(
    tmp_path: Path,
    formula_fixtures: dict[str, Path],
) -> None:
    update = index_workbook(formula_fixtures["F19"], index_dir=tmp_path / "indexes")

    with IndexStore(update.index_path) as store:
        exported = store.canonical_export()

    assert tuple(
        (row[1], row[3], row[4], row[5], row[6], row[7], row[8]) for row in exported["fblocks"]
    ) == (
        (0, 1, 1, 1, 1, 0, 0),
        (1, 1, 1, 2, 2, 0, 0),
        (2, 1, 1, 3, 3, 0, 0),
        (3, 1, 1, 4, 4, 0, 0),
        (4, 1, 1, 5, 5, 0, 0),
        (5, 1, 1, 6, 6, 0, 0),
        (6, 2, 2, 7, 7, 0, 0),
    )
    assert exported["edges"] == (
        ("fblock", "Modern", 0, "Modern", 2, 4, 9, 9, "ref"),
        ("fblock", "Modern", 1, "Modern", 1, 1, 1, 1, "spill"),
        ("fblock", "Modern", 2, "Modern", 1, 1, 1, 1, "spill"),
        ("fblock", "Modern", 3, "Modern", 2, 2, 9, 9, "ref"),
        ("fblock", "Modern", 4, "Modern", 3, 3, 9, 9, "ref"),
        ("fblock", "Modern", 5, "Modern", 2, 4, 8, 8, "ref"),
        ("fblock", "Modern", 5, "Modern", 2, 4, 9, 9, "ref"),
        ("fblock", "Modern", 6, "Modern", 2, 4, 9, 9, "ref"),
    )
    codes = {str(row[1]) for row in exported["diagnostics"]}
    assert not codes.intersection({"W_UNKNOWN_NAME", "W_PARSE"})


def test_listobject_totals_are_excluded_when_linking_columns_to_blocks(
    tmp_path: Path,
) -> None:
    descriptor = _descriptor("TableData", 0)
    metadata = _metadata(descriptor)
    table = TableInfo(
        "DataTable",
        "DataTable",
        "A1:B5",
        1,
        1,
        ("Input", "Calculated"),
    )
    cells = (
        _cell("A1", 1, 1, "Input"),
        _cell("B1", 1, 2, "Calculated"),
        _cell("A2", 2, 1, 1),
        _cell("B2", 2, 2, 2, formula="=A2*2"),
        _cell("A3", 3, 1, 2),
        _cell("B3", 3, 2, 4, formula="=A3*2"),
        _cell("A4", 4, 1, 3),
        _cell("B4", 4, 2, 6, formula="=A4*2"),
        _cell("A5", 5, 1, "Total"),
        _cell("B5", 5, 2, 12, formula="=SUM(B2:B4)"),
    )

    with IndexStore(tmp_path / "totals.xlsp.db") as store:
        store.replace_sheet_catalog(metadata.sheets)
        _replace_sheet(store, descriptor, cells, tables=(table,), part_hash="table-v1")
        assert store.replace_formula_analysis(metadata) == ("TableData",)

        links = store.connection.execute(
            """
            SELECT c.header, f.n, f.row_min, f.row_max
            FROM columns AS c
            JOIN regions AS r ON r.id = c.region_id
            LEFT JOIN fblocks AS f ON f.id = c.formula_block_id
            ORDER BY c.idx
            """
        ).fetchall()
        blocks = store.connection.execute(
            "SELECT n, row_min, row_max, col_min, col_max FROM fblocks ORDER BY n"
        ).fetchall()

    assert tuple(map(tuple, links)) == (
        ("Input", None, None, None),
        ("Calculated", 0, 2, 4),
    )
    assert tuple(map(tuple, blocks)) == ((0, 2, 4, 2, 2), (1, 5, 5, 2, 2))


def test_adjacent_listobjects_share_one_block_and_keep_distinct_input_edges(
    tmp_path: Path,
) -> None:
    descriptor = _descriptor("Tables", 0)
    metadata = _metadata(descriptor)
    tables = (
        TableInfo("TableA", "TableA", "A1:B4", 1, 0, ("Input", "Result")),
        TableInfo(
            "TableB",
            "TableB",
            "C1:E4",
            1,
            0,
            ("Result", "Input", "A]B"),
        ),
    )
    cells = (
        _cell("A1", 1, 1, "Input"),
        _cell("B1", 1, 2, "Result"),
        _cell("C1", 1, 3, "Result"),
        _cell("D1", 1, 4, "Input"),
        _cell("E1", 1, 5, "A]B"),
        _cell("A2", 2, 1, 10),
        _cell("B2", 2, 2, 10, formula="=[@Input]"),
        _cell("C2", 2, 3, 20, formula="=[@Input]"),
        _cell("D2", 2, 4, 20),
        _cell("E2", 2, 5, "first"),
        _cell("A3", 3, 1, 11),
        _cell("B3", 3, 2, 11, formula="=[@Input]"),
        _cell("C3", 3, 3, 21, formula="=[@Input]"),
        _cell("D3", 3, 4, 21),
        _cell("E3", 3, 5, "second"),
        _cell("A4", 4, 1, 12),
        _cell("B4", 4, 2, 12, formula="=[@Input]"),
        _cell("C4", 4, 3, 22, formula="=[@Input]"),
        _cell("D4", 4, 4, 22),
        _cell("E4", 4, 5, "third"),
    )

    with IndexStore(tmp_path / "adjacent-tables.xlsp.db") as store:
        store.replace_sheet_catalog(metadata.sheets)
        _replace_sheet(store, descriptor, cells, tables=tables, part_hash="tables-v1")
        assert store.replace_formula_analysis(metadata) == ("Tables",)
        exported = store.canonical_export()
        stored_headers = store.connection.execute(
            """
            SELECT t.name, c.idx, c.name
            FROM list_object_columns AS c
            JOIN list_objects AS t ON t.id = c.list_object_id
            ORDER BY t.name, c.idx
            """
        ).fetchall()

    assert exported["fblocks"] == (("Tables", 0, "=[@Input]", 2, 4, 2, 3, 0, 0),)
    assert exported["edges"] == (
        ("fblock", "Tables", 0, "Tables", 2, 4, 1, 1, "structured:TableA[Input]"),
        ("fblock", "Tables", 0, "Tables", 2, 4, 4, 4, "structured:TableB[Input]"),
    )
    assert exported["diagnostics"] == ()
    assert tuple(map(tuple, stored_headers)) == (
        ("TableA", 0, "Input"),
        ("TableA", 1, "Result"),
        ("TableB", 0, "Result"),
        ("TableB", 1, "Input"),
        ("TableB", 2, "A]B"),
    )


def test_mixed_relative_structured_range_persists_the_full_extruded_edge(
    tmp_path: Path,
) -> None:
    descriptor = _descriptor("Data", 0)
    metadata = _metadata(descriptor)
    table = TableInfo(
        "SalesTable",
        "SalesTable",
        "A1:C6",
        1,
        1,
        ("Item", "Qty", "Price"),
    )
    cells = tuple(
        _cell(
            f"E{row}",
            row,
            5,
            0,
            formula=f"=SUM(A{row}:SalesTable[Qty])",
        )
        for row in range(3, 11)
    )

    with IndexStore(tmp_path / "mixed-endpoint.xlsp.db") as store:
        store.replace_sheet_catalog(metadata.sheets)
        _replace_sheet(store, descriptor, cells, tables=(table,), part_hash="mixed-v1")
        assert store.replace_formula_analysis(metadata) == ("Data",)
        exported = store.canonical_export()

    assert exported["fblocks"] == (("Data", 0, "=SUM(RC[-4]:SalesTable[Qty])", 3, 10, 5, 5, 0, 0),)
    assert exported["edges"] == (
        (
            "fblock",
            "Data",
            0,
            "Data",
            2,
            10,
            1,
            2,
            "structured:SalesTable[Qty]",
        ),
    )


def test_computed_name_endpoint_persists_precedents_without_a_false_hull(
    tmp_path: Path,
) -> None:
    descriptor = _descriptor("Data", 0)
    metadata = _metadata(
        descriptor,
        defined_names=(DefinedName("Pick", "=INDEX($A:$A,1)", None, "formula", False),),
    )
    cells = tuple(
        _cell(
            f"C{row}",
            row,
            3,
            0,
            formula=f"=SUM(Pick:B{row + 2})",
        )
        for row in range(3, 6)
    )

    with IndexStore(tmp_path / "computed-name-endpoint.xlsp.db") as store:
        store.replace_sheet_catalog(metadata.sheets)
        _replace_sheet(store, descriptor, cells, part_hash="computed-name-v1")
        assert store.replace_formula_analysis(metadata) == ("Data",)
        exported = store.canonical_export()

    assert exported["fblocks"] == (("Data", 0, "=SUM(Pick:R[+2]C[-1])", 3, 5, 3, 3, 0, 1),)
    assert set(exported["edges"]) == {
        ("fblock", "Data", 0, "Data", 1, 1_048_576, 1, 1, "ref"),
        ("fblock", "Data", 0, "Data", 5, 7, 2, 2, "ref"),
        ("fblock", "Data", 0, None, None, None, None, None, "opaque:INDEX"),
    }
    assert tuple(row[1] for row in exported["diagnostics"]) == ("I_DYNAMIC_REF",)


def test_p3_replacement_preserves_large_sheet_and_sanitizes_external_targets(
    tmp_path: Path,
) -> None:
    descriptor = _descriptor("Data", 0)
    secret_target = (
        "https://user:password@example.test/private/budget.xlsx?sig=SECRET_TOKEN#fragment"
    )
    metadata = _metadata(descriptor, external_links={1: secret_target})
    cells = (
        _cell("A1", 1, 1, 1, formula="=NOW()"),
        _cell("B1", 1, 2, 1, formula='=INDIRECT("A1")'),
        _cell("C1", 1, 3, 1, formula="=[1]Data!A1"),
        _cell("D1", 1, 4, 1, formula="=UnknownName"),
    )

    with IndexStore(tmp_path / "diagnostics.xlsp.db") as store:
        store.replace_sheet_catalog(metadata.sheets)
        _replace_sheet(store, descriptor, cells, part_hash="formula-v1")
        store.connection.executemany(
            """
            INSERT INTO diagnostics(
                severity, code, sheet_id, row, col, ref, message, related
            ) VALUES (?, ?, 1, ?, ?, ?, ?, '{}')
            """,
            (
                ("warn", "W_LARGE_SHEET", None, None, "sheet:Data", "keep"),
                ("warn", "W_PARSE", 9, 9, "cell:Data!I9", "replace"),
            ),
        )

        store.replace_formula_analysis(metadata)
        blocks = store.connection.execute(
            "SELECT n, volatile, opaque FROM fblocks ORDER BY n"
        ).fetchall()
        edges = store.connection.execute("SELECT via FROM edges ORDER BY via").fetchall()
        diagnostics = store.connection.execute(
            "SELECT code, ref, message FROM diagnostics ORDER BY code, ref"
        ).fetchall()
        serialized = json.dumps(store.canonical_export(), ensure_ascii=False)

    assert tuple(map(tuple, blocks)) == (
        (0, 1, 0),
        (1, 1, 1),
        (2, 0, 0),
        (3, 0, 1),
    )
    assert tuple(row[0] for row in edges) == (
        "external:[budget.xlsx]",
        "opaque:INDIRECT",
        "opaque:name",
    )
    assert tuple(row[0] for row in diagnostics) == (
        "I_DYNAMIC_REF",
        "W_LARGE_SHEET",
        "W_UNKNOWN_NAME",
    )
    assert "replace" not in serialized
    for secret in ("user", "password", "private", "SECRET_TOKEN", "fragment"):
        assert secret not in serialized


def test_source_only_refresh_preserves_incoming_edges_with_backend_parity(
    tmp_path: Path,
) -> None:
    results = tuple(
        _source_refresh_result(tmp_path / backend, prefer_rtree=prefer_rtree)
        for backend, prefer_rtree in (("rtree", True), ("interval", False))
    )

    assert results[0] == results[1]
    edges, left_sources, right_sources, incoming_id_preserved = results[0]
    assert edges == (
        ("fblock", "Left", 0, "Right", 1, 1, 2, 2, "ref"),
        ("fblock", "Right", 0, "Left", 1, 1, 1, 1, "ref"),
    )
    assert left_sources == ("Right",)
    assert right_sources == ("Left",)
    assert incoming_id_preserved is True


@pytest.mark.parametrize(
    ("changed_context", "expected_reindexed"),
    (
        ("table", ("SheetOne",)),
        ("name", ("SheetOne", "SheetTwo")),
        ("external", ()),
    ),
)
def test_context_changes_reanalyze_all_formula_source_sheets(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    changed_context: str,
    expected_reindexed: tuple[str, ...],
) -> None:
    workbook = tmp_path / f"context-{changed_context}.xlsx"
    initial = _package_parts()
    _write_package(workbook, initial)
    index_dir = tmp_path / "indexes"
    index_workbook(workbook, index_dir=index_dir)

    calls: list[tuple[str, ...]] = []
    original = IndexStore.replace_formula_analysis

    def spy(
        store: IndexStore,
        metadata: WorkbookMetadata,
        sheets: Sequence[SheetDescriptor] | None = None,
    ) -> tuple[str, ...]:
        selected = metadata.sheets if sheets is None else sheets
        calls.append(tuple(sheet.name for sheet in selected))
        return original(store, metadata, sheets)

    monkeypatch.setattr(IndexStore, "replace_formula_analysis", spy)
    mutated = _package_parts(
        table_name="MuchLongerDataTable" if changed_context == "table" else "DataTable",
        name_ref="$A$1:$A$2" if changed_context == "name" else "$A$2",
        external_target=(
            "../a-much-longer-second-budget.xlsx"
            if changed_context == "external"
            else "../first-budget.xlsx"
        ),
    )
    _write_package(workbook, mutated)

    update = index_workbook(workbook, index_dir=index_dir)

    assert calls == [("SheetOne", "SheetTwo")]
    assert update.reindexed_sheets == expected_reindexed


def test_external_target_only_refresh_updates_raw_context_and_persisted_edge(
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "external-context.xlsx"
    index_dir = tmp_path / "indexes"
    formula = "[1]SheetOne!A2"
    _write_package(
        workbook,
        _package_parts(
            external_target="../first-budget.xlsx",
            sheet_two_formula=formula,
        ),
    )
    first = index_workbook(workbook, index_dir=index_dir)

    with IndexStore(first.index_path) as store:
        assert json.loads(store.get_meta("external_links") or "{}") == {"1": "../first-budget.xlsx"}
        assert store.canonical_export()["edges"] == (
            (
                "fblock",
                "SheetOne",
                0,
                "SheetOne",
                2,
                2,
                1,
                1,
                "ref",
            ),
            (
                "fblock",
                "SheetTwo",
                0,
                None,
                None,
                None,
                None,
                None,
                "external:[first-budget.xlsx]",
            ),
        )

    _write_package(
        workbook,
        _package_parts(
            external_target="../second-budget.xlsx",
            sheet_two_formula=formula,
        ),
    )
    second = index_workbook(workbook, index_dir=index_dir)

    with IndexStore(second.index_path) as store:
        assert json.loads(store.get_meta("external_links") or "{}") == {
            "1": "../second-budget.xlsx"
        }
        assert store.canonical_export()["edges"][-1][-1] == ("external:[second-budget.xlsx]")

    assert second.changed is True
    assert second.reindexed_sheets == ()
    assert second.generation == first.generation + 1


def test_incremental_formula_refresh_equals_fresh_full_index(tmp_path: Path) -> None:
    workbook = tmp_path / "canonical.xlsx"
    _write_package(workbook, _package_parts())
    incremental_dir = tmp_path / "incremental"
    index_workbook(workbook, index_dir=incremental_dir)

    _write_package(workbook, _package_parts(sheet_two_formula="SheetOne!A2"))
    incremental = index_workbook(workbook, index_dir=incremental_dir)
    full = index_workbook(workbook, index_dir=tmp_path / "full")

    with IndexStore(incremental.index_path) as incremental_store:
        incremental_export = incremental_store.canonical_export()
    with IndexStore(full.index_path) as full_store:
        full_export = full_store.canonical_export()

    assert incremental.reindexed_sheets == ("SheetTwo",)
    assert incremental_export == full_export


def test_formula_analysis_failure_rolls_back_all_mutations_and_generation(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = tmp_path / "rollback.xlsx"
    _write_package(workbook, _package_parts())
    index_dir = tmp_path / "indexes"
    first = index_workbook(workbook, index_dir=index_dir)
    with IndexStore(first.index_path) as store:
        before = store.canonical_export()
        generation = store.generation

    _write_package(
        workbook,
        _package_parts(external_target="../changed-external-target-with-long-name.xlsx"),
    )
    original = store_module.analyze_sheet_formulas

    def fail_second_sheet(
        descriptor: SheetDescriptor,
        *args: object,
        **kwargs: object,
    ) -> SheetFormulaAnalysis:
        if descriptor.name == "SheetTwo":
            raise RuntimeError("injected formula-analysis failure")
        return original(descriptor, *args, **kwargs)  # type: ignore[arg-type]

    monkeypatch.setattr(store_module, "analyze_sheet_formulas", fail_second_sheet)
    with pytest.raises(RuntimeError, match="injected formula-analysis failure"):
        index_workbook(workbook, index_dir=index_dir)

    with IndexStore(first.index_path) as store:
        assert store.generation == generation
        assert store.canonical_export() == before


def _source_refresh_result(
    root: Path,
    *,
    prefer_rtree: bool,
) -> tuple[
    tuple[tuple[object, ...], ...],
    tuple[str, ...],
    tuple[str, ...],
    bool,
]:
    root.mkdir(parents=True)
    left = _descriptor("Left", 0)
    right = _descriptor("Right", 1)
    metadata = _metadata(left, right)
    with IndexStore(root / "graph.xlsp.db", prefer_rtree=prefer_rtree) as store:
        store.replace_sheet_catalog(metadata.sheets)
        _replace_sheet(
            store,
            left,
            (_cell("A1", 1, 1, 1, formula="=Right!A1"),),
            part_hash="left-v1",
        )
        _replace_sheet(
            store,
            right,
            (_cell("A1", 1, 1, 1, formula="=Left!A1"),),
            part_hash="right-v1",
        )
        store.replace_formula_analysis(metadata)
        incoming_before = int(
            store.connection.execute(
                """
                SELECT e.id FROM edges AS e
                JOIN sheets AS s ON s.id = e.src_sheet_id
                WHERE s.name = 'Right'
                """
            ).fetchone()[0]
        )

        _replace_sheet(
            store,
            left,
            (_cell("A1", 1, 1, 2, formula="=Right!B1"),),
            part_hash="left-v2",
        )
        store.replace_formula_analysis(metadata, (left,))
        incoming_after = int(
            store.connection.execute(
                """
                SELECT e.id FROM edges AS e
                JOIN sheets AS s ON s.id = e.src_sheet_id
                WHERE s.name = 'Right'
                """
            ).fetchone()[0]
        )
        edges = store.canonical_export()["edges"]
        left_sources = _point_source_sheets(store, 1, 1, 1)
        right_sources = _point_source_sheets(store, 2, 1, 2)
    return edges, left_sources, right_sources, incoming_before == incoming_after


def _point_source_sheets(
    store: IndexStore,
    sheet_id: int,
    row: int,
    col: int,
) -> tuple[str, ...]:
    edge_ids = store.edge_store.query_point(sheet_id, row, col)
    if not edge_ids:
        return ()
    placeholders = ",".join("?" for _edge_id in edge_ids)
    rows = store.connection.execute(
        f"""
        SELECT s.name
        FROM edges AS e
        JOIN sheets AS s ON s.id = e.src_sheet_id
        WHERE e.id IN ({placeholders})
        ORDER BY s.id
        """,
        edge_ids,
    ).fetchall()
    return tuple(str(item[0]) for item in rows)


def _descriptor(name: str, order: int) -> SheetDescriptor:
    return SheetDescriptor(
        order=order,
        name=name,
        sheet_id=order + 1,
        rel_id=f"rId{order + 1}",
        xml_part=f"xl/worksheets/sheet{order + 1}.xml",
        kind="worksheet",
    )


def _metadata(
    *sheets: SheetDescriptor,
    external_links: Mapping[int, str] | None = None,
    defined_names: tuple[DefinedName, ...] = (),
) -> WorkbookMetadata:
    return WorkbookMetadata(
        path="synthetic.xlsx",
        date1904=False,
        sheets=tuple(sheets),
        defined_names=defined_names,
        external_links=MappingProxyType(dict(external_links or {})),
    )


def _cell(
    ref: str,
    row: int,
    col: int,
    value: int | str,
    *,
    formula: str | None = None,
) -> CellRecord:
    return CellRecord(
        ref=ref,
        row=row,
        col=col,
        value=value,
        value_type="number" if isinstance(value, int) else "string",
        formula=formula,
    )


def _replace_sheet(
    store: IndexStore,
    descriptor: SheetDescriptor,
    cells: Sequence[CellRecord],
    *,
    tables: tuple[TableInfo, ...] = (),
    part_hash: str,
) -> SheetParseSummary:
    def parse(on_cell: Callable[[CellRecord], None]) -> SheetParseSummary:
        for cell in cells:
            on_cell(cell)
        return SheetParseSummary(
            descriptor=descriptor,
            part_hash=part_hash,
            max_row=max((cell.row for cell in cells), default=0),
            max_col=max((cell.col for cell in cells), default=0),
            cell_count=len(cells),
            tables=tables,
        )

    return store.replace_sheet(descriptor, parse)


def _write_package(path: Path, parts: Mapping[str, bytes]) -> None:
    with ZipFile(path, mode="w", compression=ZIP_DEFLATED) as archive:
        for name in sorted(parts):
            archive.writestr(name, parts[name])


def _package_parts(
    *,
    table_name: str = "DataTable",
    name_ref: str = "$A$2",
    external_target: str = "../first-budget.xlsx",
    sheet_two_formula: str = "SheetOne!B2",
) -> dict[str, bytes]:
    return {
        "xl/workbook.xml": _xml(
            f"""
            <workbook xmlns="{MAIN_NS}" xmlns:r="{DOCUMENT_REL_NS}">
              <sheets>
                <sheet name="SheetOne" sheetId="1" r:id="rIdSheetOne"/>
                <sheet name="SheetTwo" sheetId="2" r:id="rIdSheetTwo"/>
              </sheets>
              <definedNames>
                <definedName name="Input">'SheetOne'!{name_ref}</definedName>
              </definedNames>
              <externalReferences>
                <externalReference r:id="rIdExternal"/>
              </externalReferences>
            </workbook>
            """
        ),
        "xl/_rels/workbook.xml.rels": _xml(
            f"""
            <Relationships xmlns="{PACKAGE_REL_NS}">
              <Relationship Id="rIdSheetOne" Type="{REL_TYPE_BASE}/worksheet"
                Target="worksheets/sheet1.xml"/>
              <Relationship Id="rIdSheetTwo" Type="{REL_TYPE_BASE}/worksheet"
                Target="worksheets/sheet2.xml"/>
              <Relationship Id="rIdExternal" Type="{REL_TYPE_BASE}/externalLink"
                Target="externalLinks/externalLink1.xml"/>
            </Relationships>
            """
        ),
        "xl/worksheets/sheet1.xml": _xml(
            f"""
            <worksheet xmlns="{MAIN_NS}" xmlns:r="{DOCUMENT_REL_NS}">
              <sheetData>
                <row r="1">
                  <c r="A1" t="inlineStr"><is><t>Input</t></is></c>
                  <c r="B1" t="inlineStr"><is><t>Calculated</t></is></c>
                </row>
                <row r="2">
                  <c r="A2"><v>2</v></c>
                  <c r="B2"><f>A2*2</f><v>4</v></c>
                </row>
              </sheetData>
              <tableParts count="1"><tablePart r:id="rIdTable"/></tableParts>
            </worksheet>
            """
        ),
        "xl/worksheets/sheet2.xml": _xml(
            f"""
            <worksheet xmlns="{MAIN_NS}">
              <sheetData>
                <row r="1"><c r="A1"><f>{sheet_two_formula}</f><v>4</v></c></row>
              </sheetData>
            </worksheet>
            """
        ),
        "xl/worksheets/_rels/sheet1.xml.rels": _xml(
            f"""
            <Relationships xmlns="{PACKAGE_REL_NS}">
              <Relationship Id="rIdTable" Type="{REL_TYPE_BASE}/table"
                Target="../tables/table1.xml"/>
            </Relationships>
            """
        ),
        "xl/tables/table1.xml": _xml(
            f"""
            <table xmlns="{MAIN_NS}" name="{table_name}" displayName="{table_name}"
              ref="A1:B2" headerRowCount="1" totalsRowCount="0">
              <tableColumns count="2">
                <tableColumn id="1" name="Input"/>
                <tableColumn id="2" name="Calculated"/>
              </tableColumns>
            </table>
            """
        ),
        "xl/externalLinks/externalLink1.xml": _xml(
            f"""
            <externalLink xmlns="{MAIN_NS}" xmlns:r="{DOCUMENT_REL_NS}">
              <externalBook r:id="rIdBook"/>
            </externalLink>
            """
        ),
        "xl/externalLinks/_rels/externalLink1.xml.rels": _xml(
            f"""
            <Relationships xmlns="{PACKAGE_REL_NS}">
              <Relationship Id="rIdBook" Type="{REL_TYPE_BASE}/externalLinkPath"
                Target="{external_target}" TargetMode="External"/>
            </Relationships>
            """
        ),
    }


def _xml(source: str) -> bytes:
    return "\n".join(line.strip() for line in source.strip().splitlines()).encode("utf-8")
