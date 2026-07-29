# P9 documentation and release evidence

Date: 2026-07-29. Source base: P8 milestone
`4e8e0852a69e3786020356ef66cb90bcff0916a8`; the release evidence records the
exact P9 staged-tree fingerprint and review verdicts when the candidate is
frozen.

## Release candidate scope

- README follows the frozen persuasion order, leads with the measured hero and
  live demo, documents all 14 tools, and distinguishes the passing S5 payload
  metric from the separately reported full-agent usage measurement.
- Comparison cells cite dated, immutable upstream revisions and use “not
  documented” rather than inferring nonexistence.
- Package metadata includes README, license, classifiers, project URLs, PEP 561
  marker, console entry point, and official MCP Registry ownership token.
- CI covers Linux, macOS, and Windows on Python 3.11, 3.12, and 3.13. The tag
  workflow builds and checks both distributions; trusted PyPI publishing is
  explicitly gated by repository configuration.
- Security policy links private vulnerability reporting, states the local-file
  and network boundaries, and identifies the supported 0.1.x line.
- Official MCP Registry metadata and exact account-scoped copy for Smithery,
  mcp.so, and PulseMCP are prepared.

## Reproduced release defect

The first isolated wheel probe found that MCP initialization advertised the
installed MCP SDK version (`1.29.0`) instead of Excel LSP `0.1.0`. FastMCP left
its low-level implementation version unset, so the framework supplied its own
version. The candidate sets that field from `excel_lsp.__version__` and adds
assertions to both real-stdio conformance and the clean-install probe. The final
wheel now reports `excel-lsp` / `0.1.0`.

## First formal-review remediation

The first combined reviewer returned `R-mech: APPROVE`, `R-test: REVISE`, and
`R-repo: REVISE`. The release blockers were mandatory S5, missing F17,
unpublished-PyPI commands in the committed examples, and two whitespace
defects. The main agent retained the P8 failure evidence, added the disclosed
byte-preserving benchmark workload and fresh 24-run matrix, completed F17 with
production-parser and independent-oracle coverage, changed current examples to
the public-git fallback, and removed the whitespace. Full rationale and exact
benchmark rows are in [`p9-s5-remediation.md`](p9-s5-remediation.md).

## Fresh local verification

| Check | Result |
|---|---|
| `uv run python tests/fixtures/generate.py` | 21 numbered fixture cases completed as 22 files because F09a/F09b are separate workbooks |
| `uv run pytest` | 2,109 passed, 3 live deselected, 506.95 s |
| Coverage run with `--cov-fail-under=85` | 2,109 passed, 3 live deselected; 89.65% branch coverage; 714.26 s |
| Focused benchmark/fixture/oracle tests | 42 passed, 102.26 s |
| Focused benchmark/runner/README/release tests | 48 passed, 25.36 s |
| `uv run ruff check .` | Pass |
| `uv run ruff format --check .` | Pass |
| `uv run pyright` | 0 errors, 0 warnings |
| `uv build` | Wheel and sdist built |
| `twine check dist/*` | Both distributions pass |
| Workflow YAML parse | Pass |
| Official `server.json` schema validation | Pass |
| Clean Python 3.12 wheel install/map/MCP | Version 0.1.0, 3 sheets, 14 tools, open pass |
| Local-wheel `uvx` MCP probe | Version 0.1.0, 14 tools, open pass |
| Public-git fallback | Clone/build/version/map/MCP pass |
| Isolated Codex MCP add/get/remove | Pass with Codex CLI 0.144.5 |
| Headless run guard | 73/80, including 12 deleted completed runs and 1 interrupted invocation |
| Runtime network-denial regression | 1 passed |

The final probed wheel SHA-256 is
`16400354213378e339170a2e0f5c7da214d2c3c5caae0afc9fdb512c0a30665d`.
Full sanitized commands and environment details are in
[`fresh-install.md`](fresh-install.md).

## Public repository state

<https://github.com/Brody900/excel-lsp> is public on `main`. GitHub secret
scanning and push protection are enabled. The release audit enabled private
vulnerability reporting and set the public description, homepage, and topics.
GitHub Actions run
[`30464274476`](https://github.com/Brody900/excel-lsp/actions/runs/30464274476)
passed all nine Linux/macOS/Windows × Python 3.11/3.12/3.13 jobs on temporary
commit `1d0957bc9882aed2099ea4b99563b4458383c2ea`, whose tree is the exact frozen
candidate `8e823d81d335c518d61cddf63a69c15adaae62ba` reviewed below.

## Registry and package state

The official PyPI JSON endpoint for `excel-lsp` returned 404 on 2026-07-29 and
no recognized PyPI publishing token was present. Therefore the working
public-git fallback is the release-safe path if trusted publishing is not
configured. [`server.json`](../../server.json) is ready for the official MCP
Registry after the PyPI artifact exists. The complete authenticated/paid
follow-up packet is in [`docs/registry-submissions.md`](../registry-submissions.md).

## Gate state

Local candidate construction, broad verification, and the required 3×3 CI
matrix are complete. The single combined formal reviewer returned
`R-mech: APPROVE`, `R-test: APPROVE`, and `R-repo: APPROVE` with no findings on
the same frozen tree `8e823d81d335c518d61cddf63a69c15adaae62ba` and rechecked the fingerprint at
exit. The review gate is closed.

Milestone commit `709a9779e428578245a514b682be9bd988361ec6` has tree
`f4c5c8a52c87e6f6079741d592bc8b8e78866552` and contains only the approved
candidate plus final-verdict bookkeeping. Its [final CI run](https://github.com/Brody900/excel-lsp/actions/runs/30466457991)
ended green across all nine cells after a failed-cell-only retry: the first
macOS/Python 3.12 runner passed 1,046 tests before a runner-wide SQLite `disk I/O
error` cascade; all eight other cells passed, and the fresh replacement runner
passed the same macOS/Python cell without a code change.

The annotated [`v0.1.0` tag](https://github.com/Brody900/excel-lsp/releases/tag/v0.1.0)
points to the milestone commit. [Release workflow 30469971447](https://github.com/Brody900/excel-lsp/actions/runs/30469971447)
built and checked both distributions; its PyPI job skipped because trusted
publishing was not configured. The public GitHub release attaches the wheel,
source distribution, and benchmark hero in PNG and SVG form. A fresh-cache
public-tag probe resolved the tag to the milestone commit, reported version
`0.1.0`, mapped all three F03 sheets, initialized through real MCP stdio, listed
14 tools, and opened F03 successfully. The documented Git-tag fallback is
therefore the completed distribution path for this release.
