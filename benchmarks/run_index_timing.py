"""Measure cold and one-sheet incremental indexing on the F06 family."""

from __future__ import annotations

import argparse
import csv
import json
import os
import platform
import shutil
import subprocess
import tempfile
import time
from dataclasses import asdict, dataclass
from importlib.metadata import version
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from benchmarks.run_llm_eval import resolve_codex_launcher
from excel_lsp.core.index.lifecycle import index_workbook

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / "benchmarks" / "results" / "index-timing.csv"
DEFAULT_ENVIRONMENT = ROOT / "benchmarks" / "results" / "environment.json"
SIZES = (1_000, 10_000, 50_000)


@dataclass(frozen=True, slots=True)
class TimingRow:
    rows: int
    columns: int
    repetition: int
    cold_seconds: float
    incremental_seconds: float
    cold_reindexed_sheets: int
    incremental_reindexed_sheets: int


def _patch_one_sheet(workbook: Path, replacement: int) -> None:
    """Change only Control!A2 in a benchmark copy without invoking the product editor."""
    temporary = workbook.with_suffix(workbook.suffix + ".tmp")
    needle = b'<c r="A2" t="n"><v>1</v></c>'
    encoded = str(replacement).encode("ascii")
    replacement_cell = b'<c r="A2" t="n"><v>' + encoded + b"</v></c>"
    with ZipFile(workbook, "r") as source, ZipFile(temporary, "w", ZIP_DEFLATED) as target:
        for item in source.infolist():
            payload = source.read(item.filename)
            if item.filename == "xl/worksheets/sheet2.xml":
                if payload.count(needle) != 1:
                    raise ValueError("F06 Control!A2 timing sentinel is missing or ambiguous")
                payload = payload.replace(needle, replacement_cell, 1)
            target.writestr(item, payload)
    os.replace(temporary, workbook)


def collect(repetitions: int = 3, root: Path = ROOT) -> list[TimingRow]:
    if repetitions < 1:
        raise ValueError("repetitions must be positive")
    rows: list[TimingRow] = []
    with tempfile.TemporaryDirectory(prefix="excel-lsp-index-bench-") as directory:
        workspace = Path(directory)
        for size in SIZES:
            source = root / "tests" / "fixtures" / "generated" / f"perf_{size // 1000}k.xlsx"
            if not source.is_file():
                raise FileNotFoundError(f"generate F06 before benchmarking: {source}")
            for repetition in range(1, repetitions + 1):
                workbook = workspace / f"perf-{size}-r{repetition}.xlsx"
                index_dir = workspace / f"indexes-{size}-r{repetition}"
                shutil.copy2(source, workbook)

                started = time.perf_counter()
                cold = index_workbook(workbook, index_dir=index_dir)
                cold_seconds = time.perf_counter() - started

                _patch_one_sheet(workbook, replacement=900_000 + repetition)
                started = time.perf_counter()
                incremental = index_workbook(workbook, index_dir=index_dir)
                incremental_seconds = time.perf_counter() - started

                if not cold.changed or not incremental.changed:
                    raise RuntimeError("timing run did not perform both required index updates")
                if len(incremental.reindexed_sheets) != 1:
                    raise RuntimeError("F06 incremental timing did not reindex exactly one sheet")
                rows.append(
                    TimingRow(
                        rows=size,
                        columns=10,
                        repetition=repetition,
                        cold_seconds=round(cold_seconds, 6),
                        incremental_seconds=round(incremental_seconds, 6),
                        cold_reindexed_sheets=len(cold.reindexed_sheets),
                        incremental_reindexed_sheets=len(incremental.reindexed_sheets),
                    )
                )
    return rows


def write_rows(rows: list[TimingRow], output: Path = DEFAULT_OUTPUT) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, TimingRow.__dataclass_fields__)
        writer.writeheader()
        writer.writerows(asdict(row) for row in rows)


def write_environment(output: Path = DEFAULT_ENVIRONMENT) -> None:
    codex_version = subprocess.run(
        [*resolve_codex_launcher(), "--version"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=15,
        check=False,
    ).stdout.strip()
    revision = subprocess.run(
        ["git", "rev-parse", "HEAD"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=15,
        check=False,
    ).stdout.strip()
    status = subprocess.run(
        ["git", "status", "--porcelain"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=15,
        check=False,
    )
    captured_at = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    compact_timestamp = captured_at.translate(str.maketrans("", "", "-:"))
    payload = {
        "environment_id": (
            f"{platform.system().casefold()}-{platform.machine().casefold()}-"
            f"py{platform.python_version().replace('.', '')}-{compact_timestamp}"
        ),
        "captured_at_utc": captured_at,
        "git_revision": revision or None,
        "git_worktree_dirty": bool(status.stdout.strip()) if status.returncode == 0 else None,
        "platform": platform.platform(),
        "processor": platform.processor(),
        "python": platform.python_version(),
        "codex_cli": codex_version,
        "codex_model": "gpt-5.6-sol",
        "codex_reasoning_effort": "high",
        "scripted_tokenizer": "tiktoken o200k_base",
        "dependencies": {
            package: version(package)
            for package in (
                "lxml",
                "matplotlib",
                "mcp",
                "openpyxl",
                "pydantic",
                "regex",
                "tiktoken",
                "typer",
            )
        },
        "notes": [
            "Each cold sample uses a new workbook path and empty index directory.",
            "Incremental timing begins after a one-part OOXML mutation and excludes mutation time.",
            "Filesystem cache state is uncontrolled and disclosed rather than cleared "
            "destructively.",
            "All committed timing rows in index-timing.csv share this singleton "
            "environment record.",
        ],
    }
    output.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8", newline="\n")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repetitions", type=int, default=3)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    arguments = parser.parse_args()
    rows = collect(arguments.repetitions)
    write_rows(rows, arguments.output)
    write_environment()
    print(json.dumps({"rows": len(rows), "output": str(arguments.output)}, separators=(",", ":")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
