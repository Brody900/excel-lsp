"""Generate benchmark charts from committed CSV evidence."""

from __future__ import annotations

import csv
import json
import statistics
from collections import defaultdict
from pathlib import Path
from typing import Any

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.axes import Axes
from matplotlib.figure import Figure

ROOT = Path(__file__).resolve().parents[1]
RESULTS = ROOT / "benchmarks" / "results"
ASSETS = ROOT / "docs" / "assets"
ARMS = ("excel-lsp", "naive-dump")
COLORS = {"excel-lsp": "#2563EB", "naive-dump": "#F97316"}
OUTPUTS = (
    "benchmark-token-hero",
    "benchmark-token-modes",
    "benchmark-tool-calls",
    "benchmark-index-time",
    "benchmark-audit-cost",
)


def _read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8", newline="") as stream:
        return list(csv.DictReader(stream))


def _mean_by(rows: list[dict[str, str]], field: str) -> dict[tuple[str, str], float]:
    grouped: dict[tuple[str, str], list[float]] = defaultdict(list)
    for row in rows:
        grouped[(row["task"], row["arm"])].append(float(row[field]))
    return {key: statistics.mean(values) for key, values in grouped.items()}


def _style() -> None:
    plt.rcParams.update(
        {
            "axes.spines.top": False,
            "axes.spines.right": False,
            "axes.grid": True,
            "axes.grid.axis": "y",
            "grid.alpha": 0.22,
            "font.size": 10,
            "figure.dpi": 140,
            "svg.hashsalt": "excel-lsp-benchmark-v1",
        }
    )


def _save(figure: Figure, name: str, assets: Path) -> None:
    assets.mkdir(parents=True, exist_ok=True)
    svg_path = assets / f"{name}.svg"
    figure.savefig(
        assets / f"{name}.png",
        bbox_inches="tight",
        metadata={"Software": "excel-lsp benchmark plotter"},
    )
    figure.savefig(
        svg_path,
        bbox_inches="tight",
        metadata={"Creator": "excel-lsp benchmark plotter", "Date": None},
    )
    # Matplotlib emits path data with a space before each newline. Normalize
    # generated text so the committed assets pass Git whitespace checks.
    svg = svg_path.read_text(encoding="utf-8")
    svg_path.write_text(
        "\n".join(line.rstrip() for line in svg.splitlines()) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    plt.close(figure)


def _grouped_bars(
    axis: Axes,
    values: dict[tuple[str, str], float],
    *,
    ylabel: str,
    log: bool = False,
) -> None:
    tasks = tuple(f"B{number}" for number in range(1, 7))
    x_values = list(range(len(tasks)))
    width = 0.36
    for arm_index, arm in enumerate(ARMS):
        offsets = [value + (arm_index - 0.5) * width for value in x_values]
        heights = [values[(task, arm)] for task in tasks]
        bars = axis.bar(offsets, heights, width, label=arm, color=COLORS[arm])
        axis.bar_label(bars, fmt="%.0f", fontsize=7, padding=2)
    axis.set_xticks(x_values, tasks)
    axis.set_ylabel(ylabel)
    if log:
        axis.set_yscale("log")
    axis.legend(frameon=False, ncols=2)


def plot_hero(llm_rows: list[dict[str, str]], assets: Path) -> None:
    values = _mean_by(llm_rows, "total_tokens")
    figure, axis = plt.subplots(figsize=(9.2, 4.8))
    _grouped_bars(axis, values, ylabel="Mean CLI tokens (input + output, log scale)", log=True)
    axis.set_title("Headless Codex token usage by benchmark task")
    axis.text(
        0,
        -0.2,
        "Two repetitions per task/arm; includes fixed agent context and MCP schemas.",
        transform=axis.transAxes,
        fontsize=8,
        color="#475569",
    )
    _save(figure, "benchmark-token-hero", assets)


def plot_modes(
    scripted_rows: list[dict[str, str]], llm_rows: list[dict[str, str]], assets: Path
) -> None:
    scripted = {
        arm: statistics.mean(
            float(row["payload_tokens"]) for row in scripted_rows if row["arm"] == arm
        )
        for arm in ARMS
    }
    llm = {
        arm: statistics.mean(float(row["total_tokens"]) for row in llm_rows if row["arm"] == arm)
        for arm in ARMS
    }
    figure, axis = plt.subplots(figsize=(7.8, 4.8))
    labels = ("Scripted payload\n(o200k proxy)", "Headless Codex\n(CLI usage)")
    width = 0.34
    for arm_index, arm in enumerate(ARMS):
        offsets = [index + (arm_index - 0.5) * width for index in range(2)]
        bars = axis.bar(offsets, (scripted[arm], llm[arm]), width, color=COLORS[arm], label=arm)
        axis.bar_label(bars, fmt="%.0f", fontsize=8, padding=2)
    axis.set_xticks(range(2), labels)
    axis.set_yscale("log")
    axis.set_ylabel("Mean tokens per task (log scale)")
    axis.set_title("Workbook payload and full agent usage are different measurements")
    axis.legend(frameon=False, ncols=2)
    _save(figure, "benchmark-token-modes", assets)


def plot_tool_calls(llm_rows: list[dict[str, str]], assets: Path) -> None:
    values = _mean_by(llm_rows, "tool_calls")
    figure, axis = plt.subplots(figsize=(9.2, 4.6))
    _grouped_bars(axis, values, ylabel="Mean completed tool calls")
    axis.set_ylim(0, max(values.values()) + 1.0)
    axis.set_title("Headless Codex tool calls by task")
    _save(figure, "benchmark-tool-calls", assets)


def plot_audit_cost(audit: dict[str, Any], assets: Path) -> None:
    figure, axis = plt.subplots(figsize=(8.4, 4.5))
    axis.axis("off")
    axis.set_title("Cost of one B2 formula audit", fontsize=16, pad=16)
    input_tokens = int(audit["average_input_tokens"])
    output_tokens = int(audit["average_output_tokens"])
    wall_seconds = float(audit["average_wall_seconds"])
    axis.text(
        0.5,
        0.64,
        f"{input_tokens + output_tokens:,} mean CLI tokens",
        ha="center",
        va="center",
        fontsize=26,
        color=COLORS["excel-lsp"],
        fontweight="bold",
    )
    axis.text(
        0.5,
        0.43,
        f"{input_tokens:,} input + {output_tokens:,} output  •  {wall_seconds:.1f} seconds",
        ha="center",
        va="center",
        fontsize=12,
        color="#334155",
    )
    axis.text(
        0.5,
        0.23,
        "Dollar cost not reported by Codex CLI; no conversion is fabricated.",
        ha="center",
        va="center",
        fontsize=10,
        color="#64748B",
    )
    _save(figure, "benchmark-audit-cost", assets)


def plot_index(index_rows: list[dict[str, str]], assets: Path) -> None:
    grouped: dict[int, dict[str, list[float]]] = defaultdict(
        lambda: {"cold": [], "incremental": []}
    )
    for row in index_rows:
        size = int(row["rows"])
        grouped[size]["cold"].append(float(row["cold_seconds"]))
        grouped[size]["incremental"].append(float(row["incremental_seconds"]))
    sizes = sorted(grouped)
    cold = [statistics.median(grouped[size]["cold"]) for size in sizes]
    incremental = [statistics.median(grouped[size]["incremental"]) for size in sizes]

    figure, axis = plt.subplots(figsize=(8.4, 4.8))
    axis.plot(sizes, cold, marker="o", linewidth=2.4, color=COLORS["excel-lsp"], label="Cold")
    axis.plot(
        sizes,
        incremental,
        marker="o",
        linewidth=2.4,
        color="#16A34A",
        label="Incremental (Control sheet)",
    )
    axis.axhline(10, color="#DC2626", linestyle="--", linewidth=1.2, label="Cold gate: 10s")
    axis.axhline(1, color="#A855F7", linestyle=":", linewidth=1.4, label="Incremental gate: 1s")
    axis.set_xscale("log")
    axis.set_yscale("log")
    axis.set_xticks(sizes, [f"{size // 1000}k" for size in sizes])
    axis.set_xlabel("Rows in 10-column Perf sheet")
    axis.set_ylabel("Median seconds (log scale)")
    axis.set_title("Cold and one-sheet incremental indexing (three repetitions)")
    axis.legend(frameon=False, ncols=2, fontsize=8)
    _save(figure, "benchmark-index-time", assets)


def plot_all(results: Path = RESULTS, assets: Path = ASSETS) -> tuple[Path, ...]:
    _style()
    llm_rows = _read_csv(results / "accuracy.csv")
    scripted_rows = _read_csv(results / "scripted.csv")
    index_rows = _read_csv(results / "index-timing.csv")
    audit = json.loads((results / "audit-cost.json").read_text(encoding="utf-8"))
    plot_hero(llm_rows, assets)
    plot_modes(scripted_rows, llm_rows, assets)
    plot_tool_calls(llm_rows, assets)
    plot_index(index_rows, assets)
    plot_audit_cost(audit, assets)
    return tuple(assets / f"{name}.{suffix}" for name in OUTPUTS for suffix in ("png", "svg"))


if __name__ == "__main__":
    for output in plot_all():
        print(output)
