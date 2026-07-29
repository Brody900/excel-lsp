"""Stable entry point for ``excel-lsp bench``."""

from __future__ import annotations

import json

from benchmarks.run_scripted import DEFAULT_OUTPUT, run


def run_cli() -> int:
    rows = run(DEFAULT_OUTPUT)
    failed = [row for row in rows if row.status != "ok" or not row.correct]
    print(
        json.dumps(
            {
                "mode": "scripted",
                "rows": len(rows),
                "failed": len(failed),
                "output": str(DEFAULT_OUTPUT),
                "next": "Run benchmarks/run_llm_eval.py for the cost-guarded headless mode.",
            },
            separators=(",", ":"),
        )
    )
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(run_cli())
