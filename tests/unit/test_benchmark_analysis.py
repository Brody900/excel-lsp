from __future__ import annotations

import csv
import json
import statistics
from pathlib import Path
from zipfile import ZipFile

import pytest

from benchmarks.analyze_results import (
    consolidate,
    run_guard_accounting,
    write_accuracy_table,
    write_flat_csv,
)
from benchmarks.check import check_transcript
from benchmarks.model import TASKS
from benchmarks.plot import OUTPUTS, plot_all
from benchmarks.run_index_timing import _patch_one_sheet
from excel_lsp.core.index.lifecycle import index_workbook
from excel_lsp.core.index.store import IndexStore
from excel_lsp.core.parse import OOXMLParser
from excel_lsp.core.regions import analyze_sheet_regions
from tests.fixtures.generate import _generate_f06, _generate_f13


def _row(task: str, arm: str, repetition: int, *, correct: bool = True) -> dict[str, object]:
    expected = next(spec.expected for spec in TASKS if spec.task_id == task)
    transcript = "ANSWER: " + json.dumps(expected, separators=(",", ":"))
    if not correct:
        transcript = "ANSWER: {}"
    checked, reason = check_transcript(task, transcript)
    return {
        "task": task,
        "arm": arm,
        "repetition": repetition,
        "status": "ok" if checked else "incorrect",
        "correct": checked,
        "checker_reason": reason,
        "transcript": transcript,
        "wall_seconds": 1.0,
        "input_tokens": 10,
        "cached_input_tokens": 5,
        "output_tokens": 2,
        "reasoning_output_tokens": 1,
        "cost_usd": None,
        "tool_calls": [f"{arm}:tool"],
        "protocol_violations": [],
        "return_code": 0,
    }


def _matrix() -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    excel = [_row(task.task_id, "excel-lsp", repetition) for task in TASKS for repetition in (1, 2)]
    baseline = [
        _row(task.task_id, "naive-dump", repetition) for task in TASKS for repetition in (1, 2)
    ]
    return excel, baseline


def test_consolidation_selects_only_valid_arm_sources() -> None:
    excel, baseline = _matrix()
    preflight = [*excel, _row("B1", "naive-dump", 1, correct=False)]

    rows = consolidate(preflight, baseline)

    assert len(rows) == 24
    assert all(row["correct"] for row in rows)
    assert rows == sorted(rows, key=lambda row: (row["task"], row["arm"], row["repetition"]))


def test_consolidation_rejects_incomplete_or_stale_grades() -> None:
    excel, baseline = _matrix()
    with pytest.raises(ValueError, match="exact 6 x 2 x 2"):
        consolidate(excel[:-1], baseline)

    excel[0]["correct"] = False
    with pytest.raises(ValueError, match="stored grade differs"):
        consolidate(excel, baseline)


def test_consolidation_can_auditably_regrade_unchanged_model_transcript() -> None:
    excel, baseline = _matrix()
    row = next(item for item in excel if item["task"] == "B5" and item["repetition"] == 1)
    row["transcript"] = (
        'ANSWER: {"dependent_ranges":["Summary!C9","Calc!B3:B6","Summary!C10",'
        '"Calc!C2:C6","Summary!C2:C6","Calc!D2:D6","Summary!C8"]}'
    )
    row["status"] = "incorrect"
    row["correct"] = False
    row["checker_reason"] = "answer does not exactly match the frozen expected JSON"

    rows = consolidate(excel, baseline, allow_regrade=True)
    updated = next(
        item
        for item in rows
        if item["task"] == "B5" and item["arm"] == "excel-lsp" and item["repetition"] == 1
    )

    assert updated["correct"] is True
    assert updated["checker_reason"] == "exact set"
    assert updated["regraded_without_model_rerun"] is True
    assert updated["source_grade"] == {
        "status": "incorrect",
        "correct": False,
        "checker_reason": "answer does not exactly match the frozen expected JSON",
    }


def test_flat_csv_and_accuracy_table_report_repetitions_and_agreement(tmp_path: Path) -> None:
    excel, baseline = _matrix()
    baseline[-2] = _row("B6", "naive-dump", 1, correct=False)
    rows = consolidate(excel, baseline)
    csv_path = tmp_path / "llm.csv"
    markdown_path = tmp_path / "accuracy.md"

    write_flat_csv(rows, csv_path)
    write_accuracy_table(rows, markdown_path)

    with csv_path.open(encoding="utf-8", newline="") as stream:
        flat = list(csv.DictReader(stream))
    b6 = [row for row in flat if row["task"] == "B6" and row["arm"] == "naive-dump"]
    assert [row["correct"] for row in b6] == ["False", "True"]
    assert {row["agreement"] for row in b6} == {"False"}
    markdown = markdown_path.read_text(encoding="utf-8")
    assert "| B6 | naive-dump | fail | pass | no |" in markdown
    assert "| naive-dump | 11/12 | 91.7% |" in markdown


def test_dnf_without_final_answer_reports_nonagreement_instead_of_crashing(tmp_path: Path) -> None:
    excel, baseline = _matrix()
    dnf = baseline[-2]
    dnf.update(
        {
            "transcript": "Timed out before producing an answer.",
            "status": "error",
            "correct": False,
            "checker_reason": "the final line must begin with 'ANSWER: '",
            "return_code": 124,
        }
    )
    rows = consolidate(excel, baseline)
    csv_path = tmp_path / "dnf.csv"
    markdown_path = tmp_path / "dnf.md"

    write_flat_csv(rows, csv_path)
    write_accuracy_table(rows, markdown_path)

    with csv_path.open(encoding="utf-8", newline="") as stream:
        flat = list(csv.DictReader(stream))
    b6 = [row for row in flat if row["task"] == "B6" and row["arm"] == "naive-dump"]
    assert [row["status"] for row in b6] == ["error", "ok"]
    assert {row["agreement"] for row in b6} == {"False"}
    assert "| B6 | naive-dump | fail | pass | no |" in markdown_path.read_text(encoding="utf-8")


def test_index_timing_mutation_changes_only_one_worksheet_part(tmp_path: Path) -> None:
    _generate_f06(tmp_path)
    workbook = tmp_path / "perf_1k.xlsx"
    with ZipFile(workbook) as archive:
        before = {name: archive.read(name) for name in archive.namelist()}

    _patch_one_sheet(workbook, 900_001)

    with ZipFile(workbook) as archive:
        after = {name: archive.read(name) for name in archive.namelist()}
    changed = {name for name in before if before[name] != after[name]}
    assert changed == {"xl/worksheets/sheet2.xml"}
    assert b'<c r="A2" t="n"><v>900001</v></c>' in after["xl/worksheets/sheet2.xml"]


def test_sql_region_profiles_equal_the_streaming_reference(tmp_path: Path) -> None:
    workbook = _generate_f13(tmp_path)
    update = index_workbook(workbook, index_dir=tmp_path / "index")

    with OOXMLParser(workbook) as parser, IndexStore(update.index_path) as store:
        sheet = parser.metadata.sheets[0]
        summary = parser.parse_sheet(sheet, lambda _cell: None)
        expected = analyze_sheet_regions(
            summary,
            parser.styles,
            lambda: store._iter_region_cells(1),  # pyright: ignore[reportPrivateUsage]
        )
        persisted = store.connection.execute(
            "SELECT id,n,row_min,row_max,col_min,col_max,header_rows,kind,"
            "list_object_name,confidence FROM regions WHERE sheet_id=1 ORDER BY n"
        ).fetchall()
        assert len(persisted) == len(expected.regions)
        for row, region in zip(persisted, expected.regions, strict=True):
            assert (
                int(row["n"]),
                int(row["row_min"]),
                int(row["row_max"]),
                int(row["col_min"]),
                int(row["col_max"]),
                int(row["header_rows"]),
                str(row["kind"]),
                row["list_object_name"],
                float(row["confidence"]),
            ) == (
                region.n,
                region.rect.row_min,
                region.rect.row_max,
                region.rect.col_min,
                region.rect.col_max,
                region.header_rows,
                region.kind,
                region.list_object_name,
                region.confidence,
            )
            columns = store.connection.execute(
                "SELECT idx,header,norm_header,dtype,nonnull,distinct_est FROM columns "
                "WHERE region_id=? ORDER BY idx",
                (int(row["id"]),),
            ).fetchall()
            assert [
                (
                    int(column["idx"]),
                    str(column["header"]),
                    str(column["norm_header"]),
                    str(column["dtype"]),
                    int(column["nonnull"]),
                    int(column["distinct_est"]),
                )
                for column in columns
            ] == [
                (
                    column.idx,
                    column.header,
                    column.norm_header,
                    column.dtype,
                    column.nonnull,
                    column.distinct_est,
                )
                for column in region.columns
            ]


def test_plotter_emits_all_png_and_svg_assets(tmp_path: Path) -> None:
    root = Path(__file__).parents[2]
    first = plot_all(root / "benchmarks" / "results", tmp_path / "first")
    second = plot_all(root / "benchmarks" / "results", tmp_path / "second")

    assert {path.name for path in first} == {
        f"{name}.{suffix}" for name in OUTPUTS for suffix in ("png", "svg")
    }
    assert [path.name for path in second] == [path.name for path in first]
    for output, repeated in zip(first, second, strict=True):
        assert output.stat().st_size > 1_000
        assert output.read_bytes() == repeated.read_bytes()
        if output.suffix == ".png":
            assert output.read_bytes().startswith(b"\x89PNG\r\n\x1a\n")
        else:
            assert "<svg" in output.read_text(encoding="utf-8")


def test_committed_results_reproduce_s1_and_s5_status() -> None:
    results = Path(__file__).parents[2] / "benchmarks" / "results"
    with (results / "scripted.csv").open(encoding="utf-8", newline="") as stream:
        scripted = list(csv.DictReader(stream))
    with (results / "accuracy.csv").open(encoding="utf-8", newline="") as stream:
        llm = list(csv.DictReader(stream))
    with (results / "index-timing.csv").open(encoding="utf-8", newline="") as stream:
        timing = list(csv.DictReader(stream))
    audit = json.loads((results / "audit-cost.json").read_text(encoding="utf-8"))

    assert len(scripted) == 12
    assert len(llm) == 24
    assert len(timing) == 9
    scripted_totals = {
        arm: sum(int(row["payload_tokens"]) for row in scripted if row["arm"] == arm)
        for arm in ("excel-lsp", "naive-dump")
    }
    llm_means = {
        arm: statistics.mean(int(row["total_tokens"]) for row in llm if row["arm"] == arm)
        for arm in ("excel-lsp", "naive-dump")
    }
    llm_correct = {
        arm: sum(row["correct"] == "True" for row in llm if row["arm"] == arm)
        for arm in ("excel-lsp", "naive-dump")
    }
    fifty_thousand = [row for row in timing if int(row["rows"]) == 50_000]

    assert scripted_totals == {"excel-lsp": 3_410, "naive-dump": 222_289}
    assert llm_means == {
        "excel-lsp": pytest.approx(77_310.5),
        "naive-dump": pytest.approx(64_909.833333),
    }
    assert llm_correct == {"excel-lsp": 12, "naive-dump": 8}
    assert statistics.median(float(row["cold_seconds"]) for row in fifty_thousand) == pytest.approx(
        9.439544
    )
    assert statistics.median(
        float(row["incremental_seconds"]) for row in fifty_thousand
    ) == pytest.approx(0.065912)

    assert scripted_totals["excel-lsp"] * 10 <= scripted_totals["naive-dump"]
    assert llm_means["excel-lsp"] > llm_means["naive-dump"]
    assert llm_correct["excel-lsp"] >= llm_correct["naive-dump"]
    assert audit["run_guard"] == run_guard_accounting()
    assert [entry["runs"] for entry in audit["run_guard"]["accounting"]] == [24, 12, 13, 24]
    assert audit["run_guard"]["accounting"][2]["completed_runs"] == 12
    assert audit["run_guard"]["accounting"][2]["interrupted_runs"] == 1
    assert audit["run_guard"]["observed_headless_runs"] == 73
    assert audit["run_guard"]["remaining"] == 7
