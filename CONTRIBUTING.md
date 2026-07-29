# Contributing to Excel LSP

Excel LSP was built through ordered, evidence-backed phases. Contributions are
welcome when they preserve the public v0.1 contracts, fit the architecture,
and include focused verification.

## Before starting

1. Read `AGENTS.md`, `HANDOFF.md`, and the current checklist in `PLAN.md`.
2. Search for an existing implementation, test, fixture, or decision before
   adding a new abstraction.
3. Keep a change scoped to one behavior. Avoid opportunistic refactors and
   unrelated formatting.
4. For security-sensitive reports, follow `SECURITY.md` instead of opening a
   public issue containing exploit details or private workbook data.

`CLAUDE.md` is retained only because the frozen delivery contract names that
compatibility file. Codex and `AGENTS.md` are the active orchestration and
repository instructions.

## Development setup

Install `uv`, clone the repository, and create the locked development
environment:

```console
uv sync --all-extras --dev
```

The project supports Python 3.11 and later. Do not hand-edit generated lock
content or add a production dependency without demonstrating why the existing
toolchain cannot satisfy the requirement.

## Standard checks

Run focused tests first, followed by the repository checks appropriate to the
change:

```console
uv run pytest
uv run ruff check .
uv run ruff format --check .
uv run pyright
uv run python tests/fixtures/generate.py
```

Production core coverage must remain at least 85% branch coverage. Bug fixes
should include a regression test that fails for the original behavior. Claims
about performance, Excel compatibility, or package installation require a
committed evidence artifact, not only an assertion in prose.

## Workbook fixtures

Generated workbooks come from `tests/fixtures/generate.py` using the pinned
openpyxl version plus deterministic OOXML post-processing. Do not commit files
from `tests/fixtures/generated/`; CI regenerates them. The sanctioned
`tests/fixtures/assets/vbaProject.bin` is the only committed binary fixture
source and has documented provenance in `tests/fixtures/README.md`.

Never load and save an existing workbook through openpyxl in production code.
The v0.1.0 writer is a surgical OOXML editor, and tests must prove that every
untouched ZIP part remains byte-identical.

## Code and tests

- Use the shared public value-normalization boundary for values exposed by core,
  CLI, or MCP APIs.
- Keep response limits, stable symbol IDs, structured errors, and generation
  semantics unchanged unless the authoritative specification changes.
- Prefer sparse operations; never build a dense in-memory grid for a large
  worksheet.
- Add unit tests for ordinary and error behavior. Use property, golden, oracle,
  MCP, or live tests where the phase contract requires them.
- Keep live Microsoft Excel tests under the `live` marker so automated CI does
  not pretend desktop Excel is present.

## Documentation and public claims

The README is a persuasion document, but it must remain auditable. Add or change
a product claim only when `docs/evidence/readme-claims-to-artifacts.md` names the
producing phase and committed proof. Do not invent benchmark numbers, collapse
two LLM repetitions into an unexplained average, or describe a planned feature
as released.

Public quickstarts are Codex-first. Generic MCP compatibility examples are
welcome when clearly labeled, but do not replace current Codex CLI syntax with
commands from a different agent client.

## Commits and pull requests

Use a conventional subject such as `feat:`, `fix:`, `test:`, `docs:`, or
`chore:`. A pull request should state the behavior changed, the authoritative
contract, verification commands and results, generated evidence, and remaining
risks. Never commit credentials, private workbooks, `.xlsp.db*` sidecars,
environment files, caches, or local Excel lockfiles.

The project does not accept force-pushed history on `main`. Release tags and
benchmark evidence must point to the exact reviewed commit.
