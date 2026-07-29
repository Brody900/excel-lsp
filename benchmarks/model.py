"""Frozen task definitions shared by benchmark runners and exact checkers."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True, slots=True)
class TaskSpec:
    task_id: str
    title: str
    fixture: str
    question: str
    answer_shape: str
    expected: dict[str, Any]
    unordered_array_key: str | None = None

    def prompt(self, workbook: Path) -> str:
        return (
            f"Workbook: {workbook.resolve()}\n"
            f"Task {self.task_id} — {self.title}: {self.question}\n"
            "Use only the configured workbook MCP tools. Do not use Python, shell commands, "
            "or another spreadsheet library. You may explain your reasoning briefly.\n"
            f"Expected JSON shape: {self.answer_shape}\n"
            "The last line of your reply must be exactly: `ANSWER: <json>`"
        )

    def markdown(self) -> str:
        """Render the public task contract from the evaluated task definition."""
        grading = (
            "Array order is not significant to grading; duplicates are not accepted."
            if self.unordered_array_key is not None
            else "Array order and all scalar values are significant to grading."
        )
        return (
            f"# {self.task_id} — {self.title}\n\n"
            f"Workbook: `tests/fixtures/generated/benchmarks/{self.fixture}`. This is the "
            "canonical task fixture plus the deterministic, disclosed archive workload from "
            "`benchmarks/workloads.py`.\n\n"
            f"{self.question}\n\n"
            "Use only the configured workbook MCP tools. Do not use Python, shell commands, "
            "or another spreadsheet library. You may explain your reasoning briefly.\n\n"
            f"Expected JSON shape: `{self.answer_shape}`\n\n"
            f"Grading semantics: {grading}\n\n"
            "The last line of your reply must be exactly: `ANSWER: <json>`\n"
        )


TASKS: tuple[TaskSpec, ...] = (
    TaskSpec(
        "B1",
        "Lineage",
        "cross_sheet_model.xlsx",
        "Enumerate the distinct input cells on the Inputs sheet that feed Summary!C10, "
        "directly or transitively.",
        '{"input_cells":["Sheet!A1",...]}',
        {"input_cells": ["Inputs!B2", "Inputs!B3", "Inputs!B4", "Inputs!B5"]},
        unordered_array_key="input_cells",
    ),
    TaskSpec(
        "B2",
        "Formula audit",
        "formula_blocks.xlsx",
        "Find every hand-edited formula cell that breaks the dominant copied-formula pattern.",
        '{"tampered_cells":["Sheet!A1",...]}',
        {"tampered_cells": ["FormulaBlocks!C12"]},
        unordered_array_key="tampered_cells",
    ),
    TaskSpec(
        "B3",
        "Error census",
        "errors.xlsx",
        "List every cell containing a cached Excel error and its exact Excel error code.",
        '{"errors":[{"ref":"Sheet!A1","code":"#REF!"},...]}',
        {
            "errors": [
                {"ref": f"Errors!B{row}", "code": code}
                for row, code in enumerate(
                    (
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
                    ),
                    start=2,
                )
            ]
        },
        unordered_array_key="errors",
    ),
    TaskSpec(
        "B4",
        "Schema",
        "mixed_types.xlsx",
        "Report the ordered column names and inferred dtypes for the MixedTypes table.",
        '{"columns":[{"name":"Column","dtype":"str"},...]}',
        {
            "columns": [
                {"name": "RecordID", "dtype": "int"},
                {"name": "PostingDate", "dtype": "date"},
                {"name": "Amount", "dtype": "float"},
                {"name": "MarginPct", "dtype": "float"},
                {"name": "AccountCode", "dtype": "str"},
                {"name": "Approved", "dtype": "bool"},
                {"name": "MixedSample", "dtype": "mixed"},
            ]
        },
    ),
    TaskSpec(
        "B5",
        "Impact",
        "cross_sheet_model.xlsx",
        "If Inputs!B2 changes, report the conservative dependent formula ranges returned by "
        "semantic block-level impact analysis. De-duplicate ranges.",
        '{"dependent_ranges":["Sheet!A1:A2",...]}',
        {
            "dependent_ranges": [
                "Calc!B3:B6",
                "Calc!C2:C6",
                "Calc!D2:D6",
                "Summary!C2:C6",
                "Summary!C8",
                "Summary!C9",
                "Summary!C10",
            ]
        },
        unordered_array_key="dependent_ranges",
    ),
    TaskSpec(
        "B6",
        "QA lookup",
        "cross_sheet_model.xlsx",
        "What numeric value is listed for Ending Revenue on the Summary sheet?",
        '{"value":1464.1}',
        {"value": 1464.1},
    ),
)

TASK_BY_ID = {task.task_id: task for task in TASKS}


def fixture_path(task: TaskSpec, root: Path) -> Path:
    return root / "tests" / "fixtures" / "generated" / "benchmarks" / task.fixture


__all__ = ["TASKS", "TASK_BY_ID", "TaskSpec", "fixture_path"]
