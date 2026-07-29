from __future__ import annotations

import json
import subprocess
from pathlib import Path

import pytest

from benchmarks.model import TASK_BY_ID
from benchmarks.run_llm_eval import (
    _decoded,
    _redact_workspace,
    _summarize_run,
    _windows_npm_launcher,
    build_command,
    parse_jsonl,
    sanitize_jsonl,
)


def _event(event_type: str, **payload: object) -> str:
    return json.dumps({"type": event_type, **payload}, separators=(",", ":"))


def test_build_command_isolates_excel_lsp_to_task_specific_tools(tmp_path: Path) -> None:
    command = build_command(TASK_BY_ID["B1"], "excel-lsp", tmp_path, ("codex",))
    encoded = "\n".join(command)

    assert command[:5] == ["codex", "-a", "never", "-s", "read-only"]
    assert "--ignore-user-config" not in command
    assert "--ignore-rules" in command
    assert "project_doc_max_bytes=0" in command
    assert "mcp_servers.excel_lsp.required=true" in command
    assert 'mcp_servers.excel_lsp.enabled_tools=["trace_precedents"]' in command
    assert "EXCEL_LSP_ROOT=" in encoded
    assert "shell_tool" in command
    assert "mcp_servers.serena.enabled=false" in command
    assert command[-1] == "-"


def test_build_command_configures_only_the_needed_naive_tool(tmp_path: Path) -> None:
    full = build_command(TASK_BY_ID["B3"], "naive-dump", tmp_path, ("codex",))
    sheet = build_command(TASK_BY_ID["B6"], "naive-dump", tmp_path, ("codex",))

    assert 'mcp_servers.naive_dump.enabled_tools=["read_workbook_full"]' in full
    assert 'mcp_servers.naive_dump.enabled_tools=["read_sheet"]' in sheet
    assert not any("EXCEL_LSP_ROOT=" in argument for argument in full)


def test_build_command_rejects_unknown_arm(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="unknown arm"):
        build_command(TASK_BY_ID["B1"], "other", tmp_path, ("codex",))


def test_windows_npm_launcher_prefers_packaged_native_binary(tmp_path: Path) -> None:
    shim = tmp_path / "codex.cmd"
    native = (
        tmp_path
        / "node_modules"
        / "@openai"
        / "codex"
        / "node_modules"
        / "@openai"
        / "codex-win32-x64"
        / "vendor"
        / "x86_64-pc-windows-msvc"
        / "bin"
        / "codex.exe"
    )
    native.parent.mkdir(parents=True)
    native.touch()

    assert _windows_npm_launcher(shim) == [str(native)]


def test_parse_jsonl_ignores_non_events_and_invalid_lines() -> None:
    output = "\n".join(("progress", "{bad", _event("thread.started", thread_id="t")))

    assert parse_jsonl(output) == ({"type": "thread.started", "thread_id": "t"},)


def test_workspace_redaction_covers_nested_and_slash_normalized_paths(tmp_path: Path) -> None:
    root = tmp_path / "Private Checkout"
    normalized_root = str(root).replace("\\", "/")
    value = {
        "command": [
            str(root / "benchmarks"),
            f"open {normalized_root}/fixture.xlsx",
            str(root).upper(),
        ],
        "nested": ({"cwd": str(root)},),
    }

    redacted = _redact_workspace(value, root)

    encoded = json.dumps(redacted)
    assert str(root) not in encoded
    assert normalized_root not in encoded
    assert encoded.count("<WORKSPACE>") == 4


def test_sanitize_jsonl_redacts_existing_public_rows(tmp_path: Path) -> None:
    root = tmp_path / "Private Checkout"
    path = tmp_path / "raw.jsonl"
    path.write_text(
        json.dumps({"transcript": str(root), "raw_events": [{"cwd": str(root)}]}) + "\n",
        encoding="utf-8",
    )

    assert sanitize_jsonl(path, root) == 1
    assert str(root) not in path.read_text(encoding="utf-8")
    assert json.loads(path.read_text(encoding="utf-8"))["transcript"] == "<WORKSPACE>"


def test_timeout_output_decoder_handles_subprocess_variants() -> None:
    assert _decoded(None) == ""
    assert _decoded("already text") == "already text"
    assert _decoded("café".encode()) == "café"


def test_summarize_accepts_transient_error_when_turn_completes() -> None:
    transcript = 'ANSWER: {"value":1464.1}'
    stdout = "\n".join(
        (
            _event("error", message="transient retry"),
            _event(
                "item.completed",
                item={"type": "mcp_tool_call", "server": "excel_lsp", "tool": "find"},
            ),
            _event("item.completed", item={"type": "agent_message", "text": transcript}),
            _event(
                "turn.completed",
                usage={
                    "input_tokens": 123,
                    "cached_input_tokens": 100,
                    "output_tokens": 12,
                    "reasoning_output_tokens": 3,
                },
            ),
        )
    )
    completed = subprocess.CompletedProcess(["codex"], 0, stdout, "warning")

    result = _summarize_run(TASK_BY_ID["B6"], "excel-lsp", 1, completed, 1.23456)

    assert result.status == "ok"
    assert result.correct is True
    assert result.wall_seconds == 1.235
    assert result.tool_calls == ("excel_lsp:find",)
    assert result.input_tokens == 123
    assert result.cost_usd is None


def test_summarize_redacts_workspace_from_transcript_events_and_stderr(tmp_path: Path) -> None:
    root = tmp_path / "Private Checkout"
    transcript = f'Inspected {root}.\nANSWER: {{"value":1464.1}}'
    stdout = "\n".join(
        (
            _event(
                "item.completed",
                item={
                    "type": "mcp_tool_call",
                    "server": "excel_lsp",
                    "tool": "find",
                    "arguments": {"path": str(root / "fixture.xlsx")},
                },
            ),
            _event("item.completed", item={"type": "agent_message", "text": transcript}),
            _event("turn.completed", usage={}),
        )
    )
    completed = subprocess.CompletedProcess(["codex"], 0, stdout, f"cwd={root}")

    result = _summarize_run(TASK_BY_ID["B6"], "excel-lsp", 1, completed, 0.1, root)

    encoded = json.dumps(result.raw_events)
    assert str(root) not in encoded
    assert str(root) not in result.transcript
    assert str(root) not in result.stderr_tail
    assert result.transcript.startswith("Inspected <WORKSPACE>.")
    assert result.stderr_tail == "cwd=<WORKSPACE>"


@pytest.mark.parametrize("item_type", ["command_execution", "file_change", "web_search"])
def test_summarize_rejects_disallowed_capability_use(item_type: str) -> None:
    stdout = "\n".join(
        (
            _event("item.completed", item={"type": item_type}),
            _event(
                "item.completed",
                item={"type": "agent_message", "text": 'ANSWER: {"value":1464.1}'},
            ),
            _event("turn.completed", usage={}),
        )
    )
    completed = subprocess.CompletedProcess(["codex"], 0, stdout, "")

    result = _summarize_run(TASK_BY_ID["B6"], "excel-lsp", 1, completed, 0.1)

    assert result.status == "protocol_violation"
    assert result.correct is False
    assert result.protocol_violations == (item_type,)


def test_summarize_failed_turn_overrides_exact_answer() -> None:
    stdout = "\n".join(
        (
            _event(
                "item.completed",
                item={"type": "agent_message", "text": 'ANSWER: {"value":1464.1}'},
            ),
            _event("turn.failed", error={"message": "terminal failure"}),
        )
    )
    completed = subprocess.CompletedProcess(["codex"], 1, stdout, "failed")

    result = _summarize_run(TASK_BY_ID["B6"], "excel-lsp", 1, completed, 0.1)

    assert result.status == "error"
    assert result.correct is False
