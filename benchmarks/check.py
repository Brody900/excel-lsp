"""Strict final-line benchmark answer checker."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from benchmarks.model import TASK_BY_ID

ANSWER_PREFIX = "ANSWER: "


class AnswerContractError(ValueError):
    """The transcript does not satisfy the frozen final-line contract."""


def parse_final_answer(transcript: str) -> Any:
    lines = transcript.rstrip("\r\n").splitlines()
    if not lines or not lines[-1].startswith(ANSWER_PREFIX):
        raise AnswerContractError("the final line must begin with 'ANSWER: '")
    encoded = lines[-1][len(ANSWER_PREFIX) :]
    if not encoded or encoded != encoded.strip():
        raise AnswerContractError("the final ANSWER payload must be non-empty JSON without padding")
    try:
        return json.loads(encoded)
    except json.JSONDecodeError as exc:
        raise AnswerContractError(f"the final ANSWER payload is invalid JSON: {exc.msg}") from exc


def check_transcript(task_id: str, transcript: str) -> tuple[bool, str]:
    task = TASK_BY_ID.get(task_id.upper())
    if task is None:
        return False, f"unknown task: {task_id}"
    try:
        answer = parse_final_answer(transcript)
    except AnswerContractError as exc:
        return False, str(exc)
    if task.unordered_array_key is None:
        if answer != task.expected:
            return False, "answer does not exactly match the frozen expected JSON"
        return True, "exact"
    key = task.unordered_array_key
    if not isinstance(answer, dict) or set(answer) != set(task.expected):
        return False, "answer does not match the frozen expected JSON shape"
    actual_values = answer.get(key)
    expected_values = task.expected[key]
    if not isinstance(actual_values, list) or not isinstance(expected_values, list):
        return False, "answer does not match the frozen expected JSON shape"
    canonical_actual = [
        json.dumps(value, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
        for value in actual_values
    ]
    canonical_expected = [
        json.dumps(value, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
        for value in expected_values
    ]
    if len(canonical_actual) != len(set(canonical_actual)):
        return False, "set-valued answer contains duplicate entries"
    if sorted(canonical_actual) != sorted(canonical_expected):
        return False, "answer does not match the frozen expected JSON set"
    return True, "exact set"


def _main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("task", choices=sorted(TASK_BY_ID))
    parser.add_argument("transcript", type=Path)
    arguments = parser.parse_args()
    passed, reason = check_transcript(
        arguments.task, arguments.transcript.read_text(encoding="utf-8")
    )
    print(json.dumps({"task": arguments.task, "passed": passed, "reason": reason}))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(_main())
