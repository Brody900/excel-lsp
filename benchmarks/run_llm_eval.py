"""Cost-guarded isolated headless-Codex benchmark runner."""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
from collections.abc import Sequence
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from benchmarks.check import check_transcript
from benchmarks.model import TASKS, TaskSpec, fixture_path
from benchmarks.workloads import build_workloads

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / "benchmarks" / "results" / "llm-eval.jsonl"
CODEX_MODEL = "gpt-5.6-sol"
CODEX_REASONING = "high"
MAX_RUNS = 80
MAX_COST_USD = 15.0
RUN_TIMEOUT_SECONDS = 300
REDACTED_WORKSPACE = "<WORKSPACE>"

_DISABLED_FEATURES = (
    "shell_tool",
    "plugins",
    "apps",
    "browser_use",
    "in_app_browser",
    "computer_use",
    "image_generation",
    "multi_agent",
    "goals",
)
_USER_MCP_DISABLES = (
    "mcp_servers.node_repl.enabled=false",
    "mcp_servers.openaiDeveloperDocs.enabled=false",
    "mcp_servers.serena.enabled=false",
)
_EXCEL_TOOLS: dict[str, tuple[str, ...]] = {
    "B1": ("trace_precedents",),
    "B2": ("get_diagnostics",),
    "B3": ("get_diagnostics",),
    "B4": ("list_symbols", "get_region_schema"),
    "B5": ("trace_dependents",),
    "B6": ("find", "read_range"),
}


@dataclass(frozen=True, slots=True)
class EvalResult:
    task: str
    arm: str
    repetition: int
    status: str
    correct: bool
    checker_reason: str
    transcript: str
    wall_seconds: float
    input_tokens: int | None
    cached_input_tokens: int | None
    output_tokens: int | None
    reasoning_output_tokens: int | None
    cost_usd: float | None
    tool_calls: tuple[str, ...]
    protocol_violations: tuple[str, ...]
    return_code: int
    stderr_tail: str
    raw_events: tuple[dict[str, Any], ...]


def _toml_string(value: str) -> str:
    return json.dumps(value.replace("\\", "/"))


def _windows_npm_launcher(shim: Path) -> list[str] | None:
    package = shim.parent / "node_modules" / "@openai" / "codex"
    native = sorted(package.glob("node_modules/@openai/codex-win32-*/vendor/*/bin/codex.exe"))
    if len(native) == 1:
        return [str(native[0])]
    javascript = package / "bin" / "codex.js"
    node = shutil.which("node.exe")
    if javascript.is_file() and node:
        return [node, str(javascript)]
    return None


def resolve_codex_launcher() -> list[str]:
    if os.name == "nt":
        shim = shutil.which("codex.cmd")
        if shim and (launcher := _windows_npm_launcher(Path(shim))):
            return launcher
        executable = shutil.which("codex.exe")
        if executable:
            return [executable]
    executable = shutil.which("codex")
    if executable:
        return [executable]
    raise FileNotFoundError("Codex CLI launcher was not found on PATH")


def build_command(
    task: TaskSpec,
    arm: str,
    root: Path = ROOT,
    launcher: Sequence[str] | None = None,
) -> list[str]:
    if arm not in {"excel-lsp", "naive-dump"}:
        raise ValueError(f"unknown arm: {arm}")
    command = [
        *(launcher or resolve_codex_launcher()),
        "-a",
        "never",
        "-s",
        "read-only",
        "-m",
        CODEX_MODEL,
        "-c",
        f'model_reasoning_effort="{CODEX_REASONING}"',
        "-c",
        "project_doc_max_bytes=0",
    ]
    for feature in _DISABLED_FEATURES:
        command.extend(("--disable", feature))
    for override in _USER_MCP_DISABLES:
        command.extend(("-c", override))

    server = "excel_lsp" if arm == "excel-lsp" else "naive_dump"
    if arm == "excel-lsp":
        executable = "uv"
        arguments = ["run", "excel-lsp", "serve"]
        tools = _EXCEL_TOOLS[task.task_id]
    else:
        executable = sys.executable
        arguments = ["benchmarks/baseline_server.py"]
        tools = ("read_sheet",) if task.task_id == "B6" else ("read_workbook_full",)
    command.extend(
        (
            "-c",
            f"mcp_servers.{server}.command={_toml_string(executable)}",
            "-c",
            f"mcp_servers.{server}.args={json.dumps(arguments, separators=(',', ':'))}",
            "-c",
            f"mcp_servers.{server}.cwd={_toml_string(str(root))}",
            "-c",
            f"mcp_servers.{server}.required=true",
            "-c",
            f"mcp_servers.{server}.enabled_tools={json.dumps(tools, separators=(',', ':'))}",
        )
    )
    if arm == "excel-lsp":
        allowed_root = root / "tests" / "fixtures" / "generated"
        command.extend(
            (
                "-c",
                f"mcp_servers.{server}.env={{EXCEL_LSP_ROOT={_toml_string(str(allowed_root))}}}",
            )
        )
    command.extend(
        (
            "exec",
            "--json",
            "--ephemeral",
            "--ignore-rules",
            "--skip-git-repo-check",
            "-C",
            str(root / "benchmarks"),
            "-",
        )
    )
    return command


def parse_jsonl(output: str) -> tuple[dict[str, Any], ...]:
    events: list[dict[str, Any]] = []
    for line in output.splitlines():
        try:
            value = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict) and "type" in value:
            events.append(value)
    return tuple(events)


def _redact_workspace(value: Any, root: Path = ROOT) -> Any:
    """Remove the local checkout path from public benchmark artifacts."""
    if isinstance(value, str):
        redacted = value
        variants = {str(root), str(root).replace("\\", "/")}
        for variant in variants:
            redacted = re.sub(
                re.escape(variant),
                REDACTED_WORKSPACE,
                redacted,
                flags=re.IGNORECASE,
            )
        return redacted
    if isinstance(value, dict):
        return {key: _redact_workspace(item, root) for key, item in value.items()}
    if isinstance(value, list):
        return [_redact_workspace(item, root) for item in value]
    if isinstance(value, tuple):
        return tuple(_redact_workspace(item, root) for item in value)
    return value


def _summarize_run(
    task: TaskSpec,
    arm: str,
    repetition: int,
    completed: subprocess.CompletedProcess[str],
    elapsed: float,
    root: Path = ROOT,
) -> EvalResult:
    events = _redact_workspace(parse_jsonl(completed.stdout), root)
    messages = [
        str(event["item"].get("text", ""))
        for event in events
        if event.get("type") == "item.completed"
        and isinstance(event.get("item"), dict)
        and event["item"].get("type") == "agent_message"
    ]
    transcript = messages[-1] if messages else ""
    correct, reason = check_transcript(task.task_id, transcript)
    usage: dict[str, Any] = {}
    for event in events:
        candidate = event.get("usage")
        if event.get("type") == "turn.completed" and isinstance(candidate, dict):
            usage = {str(key): value for key, value in candidate.items()}
    tool_calls: list[str] = []
    violations: list[str] = []
    for event in events:
        item = event.get("item")
        if not isinstance(item, dict) or not str(event.get("type", "")).startswith("item."):
            continue
        item_type = str(item.get("type", ""))
        if item_type == "mcp_tool_call" and event.get("type") == "item.completed":
            tool_calls.append(f"{item.get('server')}:{item.get('tool')}")
        elif item_type in {"command_execution", "file_change", "web_search"}:
            violations.append(item_type)
    failed_turn = any(event.get("type") == "turn.failed" for event in events)
    status = "ok"
    if completed.returncode != 0 or failed_turn:
        status = "error"
    elif violations:
        status = "protocol_violation"
    elif not correct:
        status = "incorrect"
    return EvalResult(
        task=task.task_id,
        arm=arm,
        repetition=repetition,
        status=status,
        correct=correct and status == "ok",
        checker_reason=reason,
        transcript=transcript,
        wall_seconds=round(elapsed, 3),
        input_tokens=_optional_int(usage.get("input_tokens")),
        cached_input_tokens=_optional_int(usage.get("cached_input_tokens")),
        output_tokens=_optional_int(usage.get("output_tokens")),
        reasoning_output_tokens=_optional_int(usage.get("reasoning_output_tokens")),
        cost_usd=_optional_float(usage.get("cost_usd")),
        tool_calls=tuple(tool_calls),
        protocol_violations=tuple(sorted(set(violations))),
        return_code=completed.returncode,
        stderr_tail=_redact_workspace(completed.stderr[-4_000:], root),
        raw_events=events,
    )


def _optional_int(value: Any) -> int | None:
    return value if type(value) is int else None


def _optional_float(value: Any) -> float | None:
    return float(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else None


def run_one(task: TaskSpec, arm: str, repetition: int, root: Path = ROOT) -> EvalResult:
    build_workloads(root)
    command = build_command(task, arm, root)
    started = time.perf_counter()
    try:
        completed = subprocess.run(
            command,
            cwd=root,
            input=task.prompt(fixture_path(task, root)),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=RUN_TIMEOUT_SECONDS,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        completed = subprocess.CompletedProcess(
            command,
            124,
            _decoded(exc.stdout),
            _decoded(exc.stderr) + f"\nTimed out after {RUN_TIMEOUT_SECONDS} seconds.",
        )
    return _summarize_run(
        task,
        arm,
        repetition,
        completed,
        time.perf_counter() - started,
        root,
    )


def _decoded(value: str | bytes | None) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return value


def _load_existing(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines() if line]


def sanitize_jsonl(path: Path, root: Path = ROOT) -> int:
    """Redact checkout paths in an existing raw JSONL artifact in place."""
    rows = _load_existing(path)
    sanitized = [_redact_workspace(row, root) for row in rows]
    payload = "".join(
        json.dumps(row, ensure_ascii=False, separators=(",", ":")) + "\n" for row in sanitized
    )
    path.write_text(payload, encoding="utf-8", newline="\n")
    return len(sanitized)


def run_matrix(
    output: Path = DEFAULT_OUTPUT,
    tasks: Sequence[TaskSpec] = TASKS,
    arms: Sequence[str] = ("excel-lsp", "naive-dump"),
    repetitions: int = 2,
    resume: bool = False,
) -> list[dict[str, Any]]:
    build_workloads(ROOT, force=True)
    existing = _load_existing(output) if resume else []
    if output.exists() and not resume:
        raise FileExistsError(f"refusing to overwrite existing eval data: {output}")
    completed_keys = {
        (str(row.get("task")), str(row.get("arm")), int(row.get("repetition", 0)))
        for row in existing
    }
    planned_keys = {
        (task.task_id, arm, repetition)
        for task in tasks
        for arm in arms
        for repetition in range(1, repetitions + 1)
    }
    pending_keys = planned_keys - completed_keys
    if len(existing) + len(pending_keys) > MAX_RUNS:
        raise ValueError(f"cost guard allows at most {MAX_RUNS} headless runs")
    known_cost = sum(float(row.get("cost_usd") or 0) for row in existing)
    output.parent.mkdir(parents=True, exist_ok=True)
    mode = "a" if resume else "w"
    with output.open(mode, encoding="utf-8", newline="\n") as stream:
        for task in tasks:
            for arm in arms:
                for repetition in range(1, repetitions + 1):
                    key = (task.task_id, arm, repetition)
                    if key in completed_keys:
                        continue
                    if known_cost >= MAX_COST_USD:
                        raise RuntimeError(f"cost guard reached ${MAX_COST_USD:.2f}")
                    print(
                        f"running {task.task_id} {arm} repetition {repetition}",
                        file=sys.stderr,
                        flush=True,
                    )
                    result = run_one(task, arm, repetition)
                    row = asdict(result)
                    stream.write(json.dumps(row, ensure_ascii=False, separators=(",", ":")) + "\n")
                    stream.flush()
                    existing.append(row)
                    known_cost += result.cost_usd or 0
                    print(
                        f"finished {task.task_id} {arm} repetition {repetition}: "
                        f"{result.status}, correct={result.correct}, {result.wall_seconds:.3f}s",
                        file=sys.stderr,
                        flush=True,
                    )
    return existing


def _main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--resume", action="store_true")
    parser.add_argument("--task", action="append", choices=[task.task_id for task in TASKS])
    parser.add_argument("--arm", action="append", choices=("excel-lsp", "naive-dump"))
    parser.add_argument("--repetitions", type=int, default=2, choices=(1, 2))
    arguments = parser.parse_args()
    selected_tasks = [
        task for task in TASKS if not arguments.task or task.task_id in arguments.task
    ]
    rows = run_matrix(
        arguments.output,
        selected_tasks,
        arguments.arm or ("excel-lsp", "naive-dump"),
        arguments.repetitions,
        arguments.resume,
    )
    print(json.dumps({"rows": len(rows), "output": str(arguments.output)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
