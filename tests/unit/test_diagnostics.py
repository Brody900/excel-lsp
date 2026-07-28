from __future__ import annotations

import json
import runpy
from collections.abc import Callable, Mapping, Sequence
from pathlib import Path
from types import MappingProxyType
from typing import cast

import pytest

from excel_lsp.core.diagnostics import (
    DIAGNOSTIC_SEVERITIES,
    P5_DEFERRED_CODES,
    Diagnostic,
    DiagnosticReport,
    DiagnosticSeverity,
    error_value_diagnostic,
    external_link_diagnostic,
    regex_timeout_diagnostic,
    stale_range_diagnostic,
)
from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.index.lifecycle import index_workbook
from excel_lsp.core.index.store import IndexStore
from excel_lsp.core.models import (
    CellRecord,
    Rect,
    SheetDescriptor,
    SheetParseSummary,
    WorkbookMetadata,
)

GenerateAll = Callable[[Path], dict[str, Path]]
generate_all = cast(
    GenerateAll,
    runpy.run_path(str(Path(__file__).parents[1] / "fixtures" / "generate.py"))["generate_all"],
)


@pytest.fixture(scope="module")
def p5_indexes(tmp_path_factory: pytest.TempPathFactory) -> dict[str, Path]:
    root = tmp_path_factory.mktemp("p5-diagnostics")
    fixtures = generate_all(root / "fixtures")
    indexes: dict[str, Path] = {}
    for fixture_id in ("F08", "F10", "F11", "F18"):
        update = index_workbook(fixtures[fixture_id], index_dir=root / "indexes")
        indexes[fixture_id] = Path(update.index_path)
    return indexes


def _snapshot(diagnostic: Diagnostic) -> dict[str, object]:
    return {
        "severity": diagnostic.severity,
        "code": diagnostic.code,
        "sheet": diagnostic.sheet,
        "row": diagnostic.row,
        "col": diagnostic.col,
        "ref": diagnostic.ref,
        "message": diagnostic.message,
        "related": dict(diagnostic.related),
    }


def test_p5_fixture_diagnostics_match_the_committed_golden(
    p5_indexes: Mapping[str, Path],
) -> None:
    golden = json.loads(
        (Path(__file__).parents[1] / "golden" / "p5-diagnostics.json").read_text(encoding="utf-8")
    )
    actual: dict[str, list[dict[str, object]]] = {}
    for fixture_id, index_path in p5_indexes.items():
        with IndexStore(index_path) as store:
            actual[fixture_id] = [_snapshot(item) for item in store.get_diagnostics().diagnostics]
    assert actual == golden


def test_catalog_is_complete_and_defers_only_later_owning_phase_findings() -> None:
    assert tuple(DIAGNOSTIC_SEVERITIES) == (
        "E_ERRVAL",
        "E_CIRCULAR",
        "E_BROKEN_XLINK",
        "W_POSSIBLE_CIRCULAR",
        "W_INCONSISTENT_FORMULA",
        "W_UNKNOWN_NAME",
        "W_PARSE",
        "W_LARGE_SHEET",
        "W_REGEX_TIMEOUT",
        "I_DYNAMIC_REF",
        "I_VOLATILE",
        "I_STALE",
    )
    assert {"W_REGEX_TIMEOUT", "I_STALE"} == P5_DEFERRED_CODES


def test_later_phase_transient_diagnostic_constructors_are_already_typed() -> None:
    timeout = regex_timeout_diagnostic("Data", "sheet:Data", deadline_ms=50)
    stale = stale_range_diagnostic(
        "Data",
        Rect(2, 4, 2, 3),
        since="2026-07-28T00:00:00Z",
    )

    assert (timeout.severity, timeout.code, timeout.related) == (
        "warn",
        "W_REGEX_TIMEOUT",
        {"deadlineMs": 50},
    )
    assert (stale.severity, stale.code, stale.ref, stale.related) == (
        "info",
        "I_STALE",
        "Data!B2:C4",
        {"range": "B2:C4", "since": "2026-07-28T00:00:00Z"},
    )
    with pytest.raises(ValueError, match="deadline"):
        regex_timeout_diagnostic("Data", "sheet:Data", deadline_ms=0)
    with pytest.raises(ValueError, match="timestamp"):
        stale_range_diagnostic("Data", Rect(1, 1, 1, 1), since="")


def test_error_diagnostic_bounds_unrecognized_ooxml_values() -> None:
    diagnostic = error_value_diagnostic("Errors", 2, 2, "B2", "X" * 1_000)

    display = diagnostic.related["errorValue"]
    assert isinstance(display, str)
    assert len(display) == 120
    assert display.endswith("…")
    assert len(diagnostic.message) < 200


def test_diagnostic_report_rejects_aggregate_counts_that_cannot_describe_results() -> None:
    valid_severities = {"error": 1, "warn": 0, "info": 0}

    with pytest.raises(ValueError, match="sum"):
        DiagnosticReport((), 1, valid_severities, {}, True)
    with pytest.raises(ValueError, match="keys"):
        DiagnosticReport((), 1, valid_severities, {1: 1}, True)  # type: ignore[dict-item]
    with pytest.raises(ValueError, match="smaller"):
        DiagnosticReport((_parse_diagnostic(),), 0, {"warn": 0}, {}, False)
    with pytest.raises(ValueError, match="truncation"):
        DiagnosticReport((), 1, {"warn": 1}, {"W_PARSE": 1}, False)
    with pytest.raises(ValueError, match="nonnegative"):
        DiagnosticReport((), 0, {"warn": -1}, {"W_PARSE": 1}, False)


def test_diagnostic_related_data_is_deeply_immutable_and_json_shaped() -> None:
    diagnostic = _parse_diagnostic(related={"details": [{"token": "?"}]})

    details = cast(tuple[Mapping[str, object], ...], diagnostic.related["details"])
    assert details == ({"token": "?"},)
    with pytest.raises(TypeError):
        details[0]["token"] = "changed"  # type: ignore[index]
    with pytest.raises(TypeError, match="keys"):
        _parse_diagnostic(related={1: "bad"})  # type: ignore[dict-item]
    with pytest.raises(TypeError, match="JSON-compatible"):
        _parse_diagnostic(related={"bad": object()})
    for nonfinite in (float("nan"), float("inf"), float("-inf")):
        with pytest.raises(TypeError, match="finite"):
            _parse_diagnostic(related={"bad": nonfinite})
    nested: object = 0
    for _level in range(65):
        nested = {"child": nested}
    with pytest.raises(TypeError, match="nesting limit"):
        _parse_diagnostic(related={"bad": nested})


def test_diagnostic_identity_and_location_invariants_reject_invalid_records() -> None:
    with pytest.raises(ValueError, match="unknown diagnostic code"):
        _parse_diagnostic(code="E_NOT_REAL")
    with pytest.raises(ValueError, match="does not allow severity"):
        _parse_diagnostic(severity="error")
    with pytest.raises(ValueError, match="must not be empty"):
        _parse_diagnostic(sheet="")
    with pytest.raises(ValueError, match="both be present"):
        _parse_diagnostic(row=1, col=None)
    with pytest.raises(ValueError, match="positive"):
        _parse_diagnostic(row=0, col=1)


def test_f08_uses_ooxml_error_type_not_a_text_whitelist(
    p5_indexes: Mapping[str, Path],
) -> None:
    with IndexStore(p5_indexes["F08"]) as store:
        report = store.get_diagnostics()

    assert report.total == 10
    assert report.counts_by_severity == {"error": 10, "warn": 0, "info": 0}
    assert report.counts_by_code == {"E_ERRVAL": 10}
    assert [item.ref for item in report.diagnostics] == [
        f"cell:Errors!B{row}" for row in range(2, 12)
    ]
    assert [item.related["errorValue"] for item in report.diagnostics] == [
        "#REF!",
        "#DIV/0!",
        "#N/A",
        "#VALUE!",
        "#NAME?",
        "#NUM!",
        "#SPILL!",
        "#CALC!",
        "#BLOCKED!",
        "#FIELD!",
    ]
    assert "unknown" not in report.diagnostics[-1].message.casefold()
    assert "Excel error value" in report.diagnostics[-1].message


def test_f10_reports_missing_local_external_link_without_path_disclosure(
    p5_indexes: Mapping[str, Path],
) -> None:
    with IndexStore(p5_indexes["F10"]) as store:
        report = store.get_diagnostics()

    assert report.total == 1
    diagnostic = report.diagnostics[0]
    assert (diagnostic.severity, diagnostic.code, diagnostic.sheet) == (
        "error",
        "E_BROKEN_XLINK",
        "External",
    )
    assert diagnostic.ref == "external:[linked-budget.xlsx]"
    assert diagnostic.related == {
        "linkIndex": 1,
        "reason": "local-file",
        "status": "missing",
        "target": "[linked-budget.xlsx]",
    }
    serialized = json.dumps(_snapshot(diagnostic), ensure_ascii=False)
    assert "missing/" not in serialized


def test_f11_and_f18_emit_one_volatile_finding_per_block(
    p5_indexes: Mapping[str, Path],
) -> None:
    with IndexStore(p5_indexes["F11"]) as store:
        dynamic = store.get_diagnostics()
    with IndexStore(p5_indexes["F18"]) as store:
        volatile = store.get_diagnostics()

    assert [(item.code, item.ref) for item in dynamic.diagnostics] == [
        ("I_DYNAMIC_REF", "cell:DynamicRefs!B2"),
        ("I_VOLATILE", "cell:DynamicRefs!B2"),
        ("I_DYNAMIC_REF", "cell:DynamicRefs!C2"),
        ("I_VOLATILE", "cell:DynamicRefs!C2"),
    ]
    assert [item.related["block"] for item in dynamic.diagnostics] == [
        "fblock:DynamicRefs:0",
        "fblock:DynamicRefs:0",
        "fblock:DynamicRefs:1",
        "fblock:DynamicRefs:1",
    ]
    assert [(item.code, item.ref, item.related["block"]) for item in volatile.diagnostics] == [
        ("I_VOLATILE", "cell:Volatile!B2", "fblock:Volatile:0"),
        ("I_VOLATILE", "cell:Volatile!B3", "fblock:Volatile:1"),
    ]


def test_get_diagnostics_filters_counts_caps_and_immutability(
    p5_indexes: Mapping[str, Path],
) -> None:
    with IndexStore(p5_indexes["F11"]) as store:
        report = store.get_diagnostics(severity="info", max_results=1)
        dynamic = store.get_diagnostics(sheet="DynamicRefs", code="I_DYNAMIC_REF")

    assert report.total == 4
    assert len(report.diagnostics) == 1
    assert report.truncated is True
    assert report.counts_by_severity == {"error": 0, "warn": 0, "info": 4}
    assert report.counts_by_code == {"I_DYNAMIC_REF": 2, "I_VOLATILE": 2}
    assert dynamic.total == 2
    assert dynamic.truncated is False
    with pytest.raises(TypeError):
        dynamic.counts_by_code["I_DYNAMIC_REF"] = 0  # type: ignore[index]
    with pytest.raises(TypeError):
        dynamic.diagnostics[0].related["block"] = "changed"  # type: ignore[index]


@pytest.mark.parametrize(
    ("kwargs", "message"),
    (
        ({"sheet": ""}, "sheet"),
        ({"severity": "fatal"}, "severity"),
        ({"code": "E_NOT_REAL"}, "code"),
        ({"max_results": 0}, "max_results"),
        ({"max_results": 101}, "max_results"),
        ({"max_results": True}, "max_results"),
    ),
)
def test_get_diagnostics_rejects_invalid_filters(
    tmp_path: Path,
    kwargs: Mapping[str, object],
    message: str,
) -> None:
    with (
        IndexStore(tmp_path / "filters.xlsp.db") as store,
        pytest.raises(ValueError, match=message),
    ):
        store.get_diagnostics(**kwargs)  # type: ignore[arg-type]


def test_external_link_health_distinguishes_existing_missing_remote_and_unsafe(
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "model.xlsx"
    workbook.touch()
    linked = tmp_path / "linked.xlsx"
    linked.touch()
    encoded_linked = tmp_path / "linked file.xlsx"
    encoded_linked.touch()

    assert external_link_diagnostic(workbook, "Data", 1, "linked.xlsx") is None
    assert external_link_diagnostic(workbook, "Data", 1, linked.as_uri()) is None
    assert external_link_diagnostic(workbook, "Data", 1, "linked%20file.xlsx") is None
    missing = external_link_diagnostic(workbook, "Data", 2, "missing.xlsx")
    remote = external_link_diagnostic(
        workbook,
        "Data",
        3,
        "https://user:secret@example.test/private/budget.xlsx?token=SECRET",
    )
    unsafe = external_link_diagnostic(workbook, "Data", 4, "file://user@host/share.xlsx")

    assert missing is not None and (missing.severity, missing.related["status"]) == (
        "error",
        "missing",
    )
    assert remote is not None and (remote.severity, remote.related["status"]) == (
        "warn",
        "remote",
    )
    assert unsafe is not None and (unsafe.severity, unsafe.related["status"]) == (
        "warn",
        "uncheckable",
    )
    public = json.dumps([_snapshot(remote), _snapshot(unsafe)], ensure_ascii=False)
    for secret in ("user", "secret", "private", "token", "SECRET", "host"):
        assert secret not in public


def test_network_external_targets_warn_without_filesystem_io(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def reject_filesystem_probe(_path: Path) -> bool:
        raise AssertionError("network target reached Path.is_file")

    monkeypatch.setattr(Path, "is_file", reject_filesystem_probe)
    targets = (
        r"\\server\share\budget.xlsx",
        "file://server/share/budget.xlsx",
        "file:////server/share/budget.xlsx",
    )

    diagnostics = [
        external_link_diagnostic(tmp_path / "model.xlsx", "Data", index, target)
        for index, target in enumerate(targets, start=1)
    ]

    assert all(diagnostic is not None for diagnostic in diagnostics)
    assert [diagnostic.related["status"] for diagnostic in diagnostics if diagnostic] == [
        "remote",
        "remote",
        "remote",
    ]
    assert "server" not in json.dumps(
        [_snapshot(diagnostic) for diagnostic in diagnostics if diagnostic]
    )


def test_unchanged_workbook_refresh_tracks_external_target_create_and_delete(
    tmp_path: Path,
) -> None:
    workbook = generate_all(tmp_path / "fixtures")["F10"]
    index_dir = tmp_path / "indexes"
    first = index_workbook(workbook, index_dir=index_dir)
    target = workbook.parent / "missing" / "linked-budget.xlsx"
    target.parent.mkdir()
    target.touch()

    appeared = index_workbook(workbook, index_dir=index_dir)
    with IndexStore(appeared.index_path) as store:
        assert store.get_diagnostics(code="E_BROKEN_XLINK").total == 0

    target.unlink()
    disappeared = index_workbook(workbook, index_dir=index_dir)
    with IndexStore(disappeared.index_path) as store:
        assert store.get_diagnostics(code="E_BROKEN_XLINK").total == 1
    stable = index_workbook(workbook, index_dir=index_dir)

    assert appeared.changed is True
    assert appeared.reindexed_sheets == ()
    assert appeared.generation == first.generation + 1
    assert disappeared.changed is True
    assert disappeared.reindexed_sheets == ()
    assert disappeared.generation == appeared.generation + 1
    assert stable.changed is False
    assert stable.generation == disappeared.generation


def test_numeric_external_link_identity_survives_same_basename_redaction(
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "model.xlsx"
    workbook.touch()
    existing = tmp_path / "dir-a" / "budget.xlsx"
    existing.parent.mkdir()
    existing.touch()
    left = _descriptor("Left", 0)
    right = _descriptor("Right", 1)
    metadata = _metadata(
        workbook,
        left,
        right,
        external_links=MappingProxyType({1: "dir-a/budget.xlsx", 2: "dir-b/budget.xlsx"}),
    )
    with IndexStore(tmp_path / "same-basename.xlsp.db") as store:
        store.replace_sheet_catalog(metadata.sheets)
        _replace_sheet(store, left, (_formula_cell("A1", "=[1]Data!A1"),))
        _replace_sheet(store, right, (_formula_cell("A1", "=[2]Data!A1"),))
        store.replace_formula_analysis(metadata)
        report = store.get_diagnostics(code="E_BROKEN_XLINK")

    assert report.total == 1
    assert report.diagnostics[0].sheet == "Right"
    assert report.diagnostics[0].related == {
        "linkIndex": 2,
        "reason": "local-file",
        "status": "missing",
        "target": "[budget.xlsx]",
    }


def test_selected_sheet_refresh_replaces_only_its_p5_diagnostics(tmp_path: Path) -> None:
    left = _descriptor("Left", 0)
    right = _descriptor("Right", 1)
    metadata = _metadata(tmp_path / "model.xlsx", left, right)
    with IndexStore(tmp_path / "incremental.xlsp.db") as store:
        store.replace_sheet_catalog(metadata.sheets)
        _replace_sheet(store, left, (_error_cell("A1", "#REF!"),))
        _replace_sheet(store, right, (_error_cell("B2", "#DIV/0!"),))
        store.replace_formula_analysis(metadata)
        assert [item.ref for item in store.get_diagnostics(code="E_ERRVAL").diagnostics] == [
            "cell:Left!A1",
            "cell:Right!B2",
        ]
        assert [item.ref for item in store.get_diagnostics(sheet="Right").diagnostics] == [
            "cell:Right!B2"
        ]
        assert store.get_diagnostics(sheet="Missing").total == 0

        _replace_sheet(store, left, (_number_cell("A1", 1),))
        store.replace_formula_analysis(metadata, (left,))
        assert [item.ref for item in store.get_diagnostics(code="E_ERRVAL").diagnostics] == [
            "cell:Right!B2"
        ]


def test_contiguous_volatile_fill_is_diagnosed_once_at_block_granularity(
    tmp_path: Path,
) -> None:
    descriptor = _descriptor("VolatileFill", 0)
    metadata = _metadata(tmp_path / "volatile-fill.xlsx", descriptor)
    cells = tuple(
        CellRecord(
            ref=f"A{row}",
            row=row,
            col=1,
            value=0,
            value_type="number",
            formula="=NOW()",
        )
        for row in range(1, 51)
    )
    with IndexStore(tmp_path / "volatile-fill.xlsp.db") as store:
        store.replace_sheet_catalog(metadata.sheets)
        _replace_sheet(store, descriptor, cells)
        store.replace_formula_analysis(metadata)
        report = store.get_diagnostics(code="I_VOLATILE")

    assert report.total == 1
    assert report.diagnostics[0].ref == "cell:VolatileFill!A1"
    assert report.diagnostics[0].related == {"block": "fblock:VolatileFill:0"}


@pytest.mark.parametrize(
    ("column", "value"),
    (
        ("code", "E_NOT_REAL"),
        ("severity", "fatal"),
        ("related", "[]"),
        ("row", 0),
    ),
)
def test_get_diagnostics_shapes_corrupt_rows(
    tmp_path: Path,
    column: str,
    value: object,
) -> None:
    descriptor = _descriptor("Data", 0)
    with IndexStore(tmp_path / f"corrupt-{column}.xlsp.db") as store:
        store.replace_sheet_catalog((descriptor,))
        store.connection.execute(
            "INSERT INTO diagnostics(severity, code, sheet_id, row, col, ref, message, related) "
            "VALUES ('warn', 'W_PARSE', 1, 1, 1, 'cell:Data!A1', 'message', '{}')"
        )
        store.connection.execute(f"UPDATE diagnostics SET {column} = ?", (value,))
        with pytest.raises(ExcelLSPError) as captured:
            store.get_diagnostics()
    assert captured.value.code is ErrorCode.CORRUPT


def test_get_diagnostics_validates_corruption_beyond_the_materialized_page(
    tmp_path: Path,
) -> None:
    descriptor = _descriptor("Data", 0)
    with IndexStore(tmp_path / "post-cap-corrupt.xlsp.db") as store:
        store.replace_sheet_catalog((descriptor,))
        store.connection.executemany(
            "INSERT INTO diagnostics(severity, code, sheet_id, row, col, ref, message, related) "
            "VALUES ('warn', 'W_PARSE', 1, ?, 1, ?, 'message', '{}')",
            ((row, f"cell:Data!A{row}") for row in range(1, 102)),
        )
        store.connection.execute("UPDATE diagnostics SET related = '[]' WHERE row = 101")
        with pytest.raises(ExcelLSPError) as captured:
            store.get_diagnostics(code="W_PARSE", max_results=100)
    assert captured.value.code is ErrorCode.CORRUPT


def test_get_diagnostics_shapes_excessive_nesting_beyond_the_materialized_page(
    tmp_path: Path,
) -> None:
    descriptor = _descriptor("Data", 0)
    nested_json = '{"child":' * 2_000 + "0" + "}" * 2_000
    with IndexStore(tmp_path / "post-cap-nested.xlsp.db") as store:
        store.replace_sheet_catalog((descriptor,))
        store.connection.executemany(
            "INSERT INTO diagnostics(severity, code, sheet_id, row, col, ref, message, related) "
            "VALUES ('warn', 'W_PARSE', 1, ?, 1, ?, 'message', '{}')",
            ((row, f"cell:Data!A{row}") for row in range(1, 102)),
        )
        store.connection.execute(
            "UPDATE diagnostics SET related = ? WHERE row = 101",
            (nested_json,),
        )
        with pytest.raises(ExcelLSPError) as captured:
            store.get_diagnostics(code="W_PARSE", max_results=100)
    assert captured.value.code is ErrorCode.CORRUPT


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
    path: Path,
    *sheets: SheetDescriptor,
    external_links: Mapping[int, str] | None = None,
) -> WorkbookMetadata:
    return WorkbookMetadata(
        path=str(path),
        date1904=False,
        sheets=tuple(sheets),
        defined_names=(),
        external_links=(MappingProxyType({}) if external_links is None else external_links),
    )


def _replace_sheet(
    store: IndexStore,
    descriptor: SheetDescriptor,
    cells: Sequence[CellRecord],
) -> None:
    def parse(on_cell: Callable[[CellRecord], None]) -> SheetParseSummary:
        for cell in cells:
            on_cell(cell)
        return SheetParseSummary(
            descriptor=descriptor,
            part_hash="diagnostic-fixture",
            max_row=max((cell.row for cell in cells), default=0),
            max_col=max((cell.col for cell in cells), default=0),
            cell_count=len(cells),
        )

    store.replace_sheet(descriptor, parse)


def _error_cell(ref: str, value: str) -> CellRecord:
    row = int("".join(character for character in ref if character.isdigit()))
    col = ord(ref[0].upper()) - ord("A") + 1
    return CellRecord(ref, row, col, value, "error")


def _number_cell(ref: str, value: int) -> CellRecord:
    row = int("".join(character for character in ref if character.isdigit()))
    col = ord(ref[0].upper()) - ord("A") + 1
    return CellRecord(ref, row, col, value, "number")


def _formula_cell(ref: str, formula: str) -> CellRecord:
    row = int("".join(character for character in ref if character.isdigit()))
    col = ord(ref[0].upper()) - ord("A") + 1
    return CellRecord(ref, row, col, 0, "number", formula=formula)


def _parse_diagnostic(
    *,
    severity: DiagnosticSeverity = "warn",
    code: str = "W_PARSE",
    sheet: str = "Data",
    row: int | None = 1,
    col: int | None = 1,
    related: Mapping[str, object] | None = None,
) -> Diagnostic:
    return Diagnostic(
        severity=severity,
        code=code,
        sheet=sheet,
        row=row,
        col=col,
        ref="cell:Data!A1",
        message="Formula could not be parsed.",
        related={} if related is None else related,
    )
