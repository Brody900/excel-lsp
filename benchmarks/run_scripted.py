"""Deterministic no-LLM benchmark replays for Excel LSP and the naive dump."""

from __future__ import annotations

import csv
import json
import time
from collections.abc import Callable
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

import tiktoken

from benchmarks.baseline_server import read_sheet, read_workbook_full
from benchmarks.check import check_transcript
from benchmarks.model import TASKS, TaskSpec, fixture_path
from excel_lsp.server.service import ToolService

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / "benchmarks" / "results" / "scripted.csv"
ENCODING = "o200k_base"


@dataclass(frozen=True, slots=True)
class ScriptedResult:
    mode: str
    task: str
    arm: str
    status: str
    payload_tokens: int
    tool_calls: int
    wall_ms: float
    correct: bool
    checker_reason: str
    answer_json: str
    transcript: str


def _excel_calls(task: TaskSpec, workbook: Path) -> list[tuple[str, dict[str, Any]]]:
    service = ToolService()
    path = str(workbook)
    if task.task_id == "B1":
        return [("trace_precedents", service.trace_precedents(path, "Summary!C10", 8, 200))]
    if task.task_id == "B2":
        return [("get_diagnostics", service.get_diagnostics(path))]
    if task.task_id == "B3":
        return [("get_diagnostics", service.get_diagnostics(path))]
    if task.task_id == "B4":
        return [
            ("list_symbols", service.list_symbols(path, kinds=["regions"])),
            ("get_region_schema", service.get_region_schema(path, "region:MixedTypes:0")),
        ]
    if task.task_id == "B5":
        return [("trace_dependents", service.trace_dependents(path, "Inputs!B2", 8, 200))]
    if task.task_id == "B6":
        return [
            ("find", service.find(path, "Ending Revenue", ["values"], "Summary")),
            ("read_range", service.read_range(path, "Summary!B8:C8")),
        ]
    raise AssertionError(task.task_id)


def _baseline_calls(task: TaskSpec, workbook: Path) -> list[tuple[str, dict[str, Any]]]:
    if task.task_id == "B6":
        return [("read_sheet", read_sheet(str(workbook), "Summary"))]
    return [("read_workbook_full", read_workbook_full(str(workbook)))]


def _payload_text(calls: list[tuple[str, dict[str, Any]]]) -> str:
    return "\n".join(
        f"tool={name}\nresult={json.dumps(payload, ensure_ascii=False, indent=2, default=str)}"
        for name, payload in calls
    )


def _transcript(task: TaskSpec) -> str:
    return "ANSWER: " + json.dumps(task.expected, ensure_ascii=False, separators=(",", ":"))


def collect_scripted(root: Path = ROOT) -> list[ScriptedResult]:
    encoding = tiktoken.get_encoding(ENCODING)
    results: list[ScriptedResult] = []
    arms: tuple[tuple[str, Callable[[TaskSpec, Path], list[tuple[str, dict[str, Any]]]]], ...] = (
        ("excel-lsp", _excel_calls),
        ("naive-dump", _baseline_calls),
    )
    for task in TASKS:
        workbook = fixture_path(task, root)
        for arm, replay in arms:
            started = time.perf_counter()
            calls = replay(task, workbook)
            elapsed_ms = (time.perf_counter() - started) * 1_000
            transcript = _transcript(task)
            correct, reason = check_transcript(task.task_id, transcript)
            results.append(
                ScriptedResult(
                    mode="scripted",
                    task=task.task_id,
                    arm=arm,
                    status="ok",
                    payload_tokens=len(encoding.encode(_payload_text(calls))),
                    tool_calls=len(calls),
                    wall_ms=round(elapsed_ms, 3),
                    correct=correct,
                    checker_reason=reason,
                    answer_json=json.dumps(
                        task.expected, ensure_ascii=False, separators=(",", ":")
                    ),
                    transcript=transcript,
                )
            )
    return results


def write_scripted(results: list[ScriptedResult], output: Path = DEFAULT_OUTPUT) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=tuple(asdict(results[0])))
        writer.writeheader()
        writer.writerows(asdict(result) for result in results)


def run(output: Path = DEFAULT_OUTPUT, root: Path = ROOT) -> list[ScriptedResult]:
    results = collect_scripted(root)
    write_scripted(results, output)
    return results


if __name__ == "__main__":
    rows = run()
    print(json.dumps({"rows": len(rows), "output": str(DEFAULT_OUTPUT)}))
