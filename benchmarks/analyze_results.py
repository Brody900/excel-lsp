"""Consolidate raw eval runs and emit auditable benchmark tables."""

from __future__ import annotations

import csv
import json
from collections import defaultdict
from pathlib import Path
from typing import Any

from benchmarks.check import AnswerContractError, check_transcript, parse_final_answer
from benchmarks.model import TASKS

ROOT = Path(__file__).resolve().parents[1]
RESULTS = ROOT / "benchmarks" / "results"
PREFLIGHT = RESULTS / "llm-eval-preflight-mixed.jsonl"
BASELINE_RERUN = RESULTS / "llm-eval-baseline-rerun.jsonl"
CONSOLIDATED = RESULTS / "llm-eval.jsonl"
FLAT_CSV = RESULTS / "accuracy.csv"
ACCURACY_TABLE = RESULTS / "accuracy.md"
AUDIT_COST = RESULTS / "audit-cost.json"


def load_jsonl(path: Path) -> list[dict[str, Any]]:
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line]


def consolidate(
    preflight: list[dict[str, Any]],
    baseline_rerun: list[dict[str, Any]],
    *,
    allow_regrade: bool = False,
) -> list[dict[str, Any]]:
    rows = [row for row in preflight if row.get("arm") == "excel-lsp"]
    rows.extend(row for row in baseline_rerun if row.get("arm") == "naive-dump")
    rows.sort(key=lambda row: (str(row["task"]), str(row["arm"]), int(row["repetition"])))
    expected_keys = {
        (task.task_id, arm, repetition)
        for task in TASKS
        for arm in ("excel-lsp", "naive-dump")
        for repetition in (1, 2)
    }
    actual_keys = {(str(row["task"]), str(row["arm"]), int(row["repetition"])) for row in rows}
    if actual_keys != expected_keys or len(rows) != len(expected_keys):
        raise ValueError("consolidated eval rows are not the exact 6 x 2 x 2 matrix")
    canonical_rows: list[dict[str, Any]] = []
    for source_row in rows:
        row = dict(source_row)
        correct, reason = check_transcript(str(row["task"]), str(row["transcript"]))
        expected_status = "ok" if correct else "incorrect"
        if row.get("protocol_violations"):
            expected_status = "protocol_violation"
            correct = False
        if row.get("return_code") != 0:
            expected_status = "error"
            correct = False
        grade_changed = (
            bool(row.get("correct")) != correct
            or row.get("status") != expected_status
            or str(row.get("checker_reason")) != reason
        )
        if grade_changed and not allow_regrade:
            raise ValueError(
                f"stored grade differs from current checker: {row['task']} {row['arm']}"
            )
        if grade_changed:
            row["source_grade"] = {
                "status": row.get("status"),
                "correct": bool(row.get("correct")),
                "checker_reason": row.get("checker_reason"),
            }
            row["status"] = expected_status
            row["correct"] = correct
            row["checker_reason"] = reason
            row["regraded_without_model_rerun"] = True
        canonical_rows.append(row)
    return canonical_rows


def _agreements(rows: list[dict[str, Any]]) -> dict[tuple[str, str], bool]:
    answers: dict[tuple[str, str], list[tuple[bool, Any]]] = defaultdict(list)
    for row in rows:
        try:
            answer = parse_final_answer(str(row["transcript"]))
        except AnswerContractError:
            parsed = (False, None)
        else:
            parsed = (True, answer)
        answers[(str(row["task"]), str(row["arm"]))].append(parsed)
    return {
        key: len(values) == 2 and all(valid for valid, _answer in values) and values[0] == values[1]
        for key, values in answers.items()
    }


def write_consolidated(rows: list[dict[str, Any]], path: Path = CONSOLIDATED) -> None:
    encoded = "\n".join(json.dumps(row, ensure_ascii=False, separators=(",", ":")) for row in rows)
    path.write_text(encoded + "\n", encoding="utf-8", newline="\n")


def write_flat_csv(rows: list[dict[str, Any]], path: Path = FLAT_CSV) -> None:
    agreements = _agreements(rows)
    fields = (
        "task",
        "arm",
        "repetition",
        "status",
        "correct",
        "agreement",
        "wall_seconds",
        "input_tokens",
        "cached_input_tokens",
        "output_tokens",
        "reasoning_output_tokens",
        "total_tokens",
        "tool_calls",
        "cost_usd_reported",
    )
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fields)
        writer.writeheader()
        for row in rows:
            input_tokens = int(row.get("input_tokens") or 0)
            output_tokens = int(row.get("output_tokens") or 0)
            writer.writerow(
                {
                    "task": row["task"],
                    "arm": row["arm"],
                    "repetition": row["repetition"],
                    "status": row["status"],
                    "correct": row["correct"],
                    "agreement": agreements[(str(row["task"]), str(row["arm"]))],
                    "wall_seconds": row["wall_seconds"],
                    "input_tokens": row.get("input_tokens"),
                    "cached_input_tokens": row.get("cached_input_tokens"),
                    "output_tokens": row.get("output_tokens"),
                    "reasoning_output_tokens": row.get("reasoning_output_tokens"),
                    "total_tokens": input_tokens + output_tokens,
                    "tool_calls": len(row.get("tool_calls", [])),
                    "cost_usd_reported": row.get("cost_usd"),
                }
            )


def write_accuracy_table(rows: list[dict[str, Any]], path: Path = ACCURACY_TABLE) -> None:
    agreements = _agreements(rows)
    by_key = {(str(row["task"]), str(row["arm"]), int(row["repetition"])): row for row in rows}
    lines = [
        "# Headless Codex exact-answer accuracy",
        "",
        "Each repetition is shown separately. Agreement compares parsed final JSON, not prose.",
        "",
        "| Task | Arm | Rep 1 | Rep 2 | Agreement |",
        "|---|---|---:|---:|---:|",
    ]
    for task in TASKS:
        for arm in ("excel-lsp", "naive-dump"):
            first = by_key[(task.task_id, arm, 1)]
            second = by_key[(task.task_id, arm, 2)]
            agreement = agreements[(task.task_id, arm)]
            lines.append(
                f"| {task.task_id} | {arm} | {_mark(first)} | {_mark(second)} | "
                f"{'yes' if agreement else 'no'} |"
            )
    lines.extend(("", "| Arm | Exact answers | Accuracy |", "|---|---:|---:|"))
    for arm in ("excel-lsp", "naive-dump"):
        arm_rows = [row for row in rows if row["arm"] == arm]
        correct = sum(bool(row["correct"]) for row in arm_rows)
        lines.append(f"| {arm} | {correct}/{len(arm_rows)} | {correct / len(arm_rows):.1%} |")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")


def _mark(row: dict[str, Any]) -> str:
    return "pass" if row["correct"] else "fail"


def write_audit_cost(rows: list[dict[str, Any]], path: Path = AUDIT_COST) -> None:
    audit = [row for row in rows if row["task"] == "B2" and row["arm"] == "excel-lsp"]
    payload = {
        "task": "B2 formula audit",
        "model": "gpt-5.6-sol",
        "reasoning_effort": "high",
        "repetitions": len(audit),
        "average_input_tokens": round(sum(int(row["input_tokens"]) for row in audit) / len(audit)),
        "average_output_tokens": round(
            sum(int(row["output_tokens"]) for row in audit) / len(audit)
        ),
        "average_wall_seconds": round(
            sum(float(row["wall_seconds"]) for row in audit) / len(audit), 3
        ),
        "cost_usd_reported": None,
        "cost_note": (
            "Codex CLI 0.144.5 with ChatGPT authentication emitted token usage but no cost field; "
            "no unsupported dollar conversion is claimed."
        ),
        "run_guard": {"observed_headless_runs": 36, "maximum": 80},
    }
    path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8", newline="\n")


def run() -> list[dict[str, Any]]:
    rows = consolidate(
        load_jsonl(PREFLIGHT),
        load_jsonl(BASELINE_RERUN),
        allow_regrade=True,
    )
    write_consolidated(rows)
    write_flat_csv(rows)
    write_accuracy_table(rows)
    write_audit_cost(rows)
    return rows


if __name__ == "__main__":
    print(json.dumps({"rows": len(run()), "output": str(CONSOLIDATED)}, separators=(",", ":")))
