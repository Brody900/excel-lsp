# Excel LSP — Build Handoff Document

> **Codex execution note (2026-07-15):** Codex is the orchestrator for this
> build. Every reference below to Claude or headless Claude means Codex or
> headless Codex, respectively. Anthropic-specific commands and flags are
> behavioral examples only; the implementation must use verified equivalents
> from the installed Codex CLI. `CLAUDE.md` remains only because this handoff's
> frozen Definition of Done names it; repository instructions for Codex live in
> `AGENTS.md`. No Claude/Anthropic service is used by this project.

**Version:** 1.1 (2026-07-15). Supersedes v1.0. All changes are integrated into the body, which is authoritative; the delta summary below is for orientation only.

**Audience:** You are a Claude Code agent acting as the **orchestrator** for this project. You will build, test, benchmark, document, and release this project end-to-end on this machine, spawning subagents for implementation, review, and benchmarking as specified in §11.

**Read this entire document before writing any code.**

---

## What changed in v1.1 (orientation only — the body is authoritative)

1. **Write path redesigned (§5.7).** openpyxl's load/save silently drops charts, images, and drawings from existing files — with no repair dialog, so the v1.0 live pass (S4) would have passed while destroying user content. All edits now go through a **surgical lxml XML patcher** that copies every untouched zip part byte-for-byte (new invariant I18; fixtures F16/F21). This also eliminates the slow, memory-heavy normal-mode load on large workbooks. `keep_vba` is obsolete: VBA survives by construction.
2. **Read path promoted (§2, §5.1).** The single-pass lxml parser (v1.0's "upgrade path") is now the **primary indexer**. Four findings forced this: shared-formula expansion in openpyxl read-only mode was unverified; `ws.tables` and `merged_cells` are likely unavailable in read-only mode; the dual load was the S1 bottleneck; and date/dtype classification needs `styles.xml` parsing we control anyway. openpyxl remains for fixture generation, the formula tokenizer/Translator, and as a **test oracle** (§8.1a).
3. **Generated fixtures now carry cached values (§8.3).** openpyxl never computes formulas, so v1.0's generated fixtures would have had no cached value for any formula cell — silently gutting `E_ERRVAL` detection, `profile_column`, dtype inference on formula columns, and benchmark B3. `generate.py` now injects `<v>` (and `t="e"` for errors) via an lxml post-pass.
4. **`W_INCONSISTENT_FORMULA` re-specified (§5.4).** The v1.0 definition was vacuous: fblocks partition formula cells by identical R1C1, so no differing cell can lie inside a block's rectangle. It is now a minority-pattern detector over per-column formula runs. Benchmark B2 depends on it.
5. **Modern formula constructs (§5.3, fixture F19).** `_xlfn.`/`_xlws.` prefix normalization; LET-binding and LAMBDA-parameter suppression (v1.0 would emit a spurious `W_UNKNOWN_NAME` on every LET); LAMBDA-in-names resolved for function-call identifiers; spill refs `A1#`; implicit intersection `@`.
6. **Defined names generalized (§4.1).** `kind` column + `name_areas` child table cover multi-area, constant, formula, and LAMBDA names; the single-rect schema could not represent them.
7. **Tool surface (§6).** New tool `trace_path` (**14 tools total**); `get_region_schema` sample rows capped so wide tables cannot breach the 200-value rule, and data-validation constraints included; MCP tool annotations (`readOnlyHint`/`destructiveHint`); server `instructions` at initialize; progress notifications during indexing; canonical tool-error table.
8. **Operational hardening (§5.8, §6.1).** Per-call freshness check with automatic incremental reindex; cursors carry an index `generation` (→ `E_STALE_CURSOR`); `~$` lockfile detection before writes (→ `E_OPEN_IN_EXCEL`); regex guard on `find`; optional `EXCEL_LSP_ROOT` path confinement; SQLite WAL + busy_timeout; path-hash DB filenames under `EXCEL_LSP_INDEX_DIR`; `*.xlsp.db*` gitignore.
9. **Fixtures.** F19 (modern functions), F20 (map stress), F21 (chart+image preservation) added: **21 generated + 3 live**. F16 (.xlsm) gets a committed `vbaProject.bin` blob as the one sanctioned binary exception (v1.0's "scripts, not binaries" rule made F16 impossible to generate); provenance documented, blob authored at P0.
10. **Map degradation rules (§4.3, F20)** so many-sheet / many-name workbooks cannot blow the 8k response cap; `hidden`/`veryHidden` visibility and `hasVBA` surfaced in the map (hidden sheets are gold for audits).
11. **Circular re-verification made O(cells) (§5.5)** via a per-cell self-inclusion test, with a bounded fallback that downgrades to `W_POSSIBLE_CIRCULAR` instead of spending O(n²) on 50k-row running totals.
12. **Benchmarks (§9).** Strict `ANSWER:` final-line contract so `check.py` can grade transcripts; both repetitions reported individually (n=2 disagreement is signal); incremental-reindex series added to chart 5; openpyxl/tiktoken pinned so regenerated fixtures and goldens stay stable across the CI matrix.
13. **Orchestration (§0, §11).** Repo-root `CLAUDE.md` required (subagents inherit it; survives context compaction); subagent failure protocol (§11.6); one R-repo invocation spent early at P2 on the README-claims-to-artifacts plan; invariant I2 re-specified to be actually testable (canonical exports, not raw rows with surrogate ids); extrusion clamped to sheet bounds; `#CALC!`/`#BLOCKED!` and `RANDARRAY`/`CELL`/`INFO` added; external-link `[n]` index mapping and chartsheet tolerance specified; S2 wording aligned to the `open_workbook` tool.

---

## 0. How to operate

1. Create `PLAN.md` (phase checklist, from §11.2), `docs/agent-log.md` (append-only decision log), and **`CLAUDE.md`** (repo root; template in §14.4) as your first files. Claude Code auto-loads `CLAUDE.md`, so every subagent inherits the conventions without brief bloat, and it survives context compaction — keep it current. Every non-trivial decision you make gets a dated entry in the agent log: *decision, alternatives considered, rationale*.
2. Work through Phases 0–9 in order. Each phase has deliverables and review gates. Do not skip gates.
3. Commit per milestone with conventional commits (`feat:`, `fix:`, `test:`, `docs:`, `chore:`). Never force-push. Work on `main`.
4. **Decision authority.** You are trusted to decide: heuristic weights and thresholds, internal module naming, extra diagnostics or tools beyond the required set, chart styling, additional fixtures, type-checker strictness details. You may NOT change: the symbol ID scheme (§4.2), the no-bulk-data rule (§6.1), **the part-preservation rule (§5.7 / I18)**, tool response caps, review budgets (§11.4), the license, the package name, phase gating, or the Definition of Done (§13). Log every delegated decision.
5. **Pause points — stop and ask the user when you hit these:**
   - GitHub remote creation / push credentials, if not already configured.
   - PyPI token at release time (Phase 9). If the user doesn't provide one, fall back to a git tag release and document `uvx --from git+<repo-url> excel-lsp` as the install path.
   - Anthropic auth for LLM evals if headless `claude -p` fails auth (test with `claude -p "ping" --output-format json --max-turns 1` early in Phase 0).
   - OS screen-control / accessibility permission prompts for computer use, or Excel activation dialogs on first launch — this includes authoring the F16 VBA blob at P0 (§8.3).
6. If a required capability is genuinely absent from this machine (e.g., Excel not installed), pause and report; do not silently substitute.
7. **Subagent failure handling:** follow §11.6. Never let a subagent churn past two failed attempts at its done-criteria.

---

## 1. Mission and product spec

**Mission.** Excel LSP is an open-source, LSP-style semantic index for Excel workbooks, exposed to AI agents via MCP. Agents navigate workbooks the way coding agents navigate repos with Serena: by symbols, references, and diagnostics — never by dumping cell ranges into context. It is a persistent, incrementally updated index with go-to-definition, find-precedents/dependents, workspace symbols, and formula diagnostics.

**What it is:** a Python library (core) + MCP server (stdio) + debug CLI, backed by a SQLite index, operating on `.xlsx` / `.xlsm` / `.xltx` / `.xltm` files on disk.

**Positioning note for all public copy:** qualify "LSP" once, prominently — *"LSP-style: the ideas (symbols, references, diagnostics, incremental index), not the LSP wire protocol."* This pre-empts the inevitable pedantry; an actual LSP wire-protocol server for formula editing in editors is a roadmap item (§10.2).

**Non-goals for v1** (state these in the README): no formula recalculation engine (we read cached values and delegate recalc to Excel), no chart/pivot *creation*, no rename refactoring (roadmap), no datetime cell *writes* (§5.7; roadmap), no Google Sheets, no `.xls`/`.xlsb` (legacy binary), no collaborative/live editing, no network access at runtime, no telemetry.

**Success criteria (verifiable, all must hold at DoD):**

| # | Criterion |
|---|---|
| S1 | Indexes a 50,000-row × 10-col workbook in < 10 s cold, < 1 s incremental after a one-sheet change |
| S2 | The map returned by `open_workbook` for the reference model fixture ≤ 1,500 tokens (tiktoken proxy, §9.4) |
| S3 | Cross-sheet precedent/dependent traces are exactly correct on all graph fixtures (F03, F04, F05, F15, F19) |
| S4 | A workbook edited through Excel LSP opens in desktop Excel with **zero** repair dialogs, recalculated values match expectations (live pass, §8.6), **and every zip part not deliberately modified is byte-identical** — charts, images, and VBA preserved (F16/F21 part-diff, I18) |
| S5 | Benchmarks show ≥ 10× token reduction vs. the naive-dump baseline on the defined task suite, with equal-or-better task accuracy in LLM evals |
| S6 | No tool response ever exceeds the response cap or returns bulk data (enforced by tests, §6.1) |
| S7 | `uvx excel-lsp serve` (or the git+ fallback) works from a clean environment |

---

## 2. Stack and foundational decisions

- Python ≥ 3.11. Package manager: `uv` (**commit `uv.lock`** — reproducibility underpins fixture determinism and R-repo's fresh-env check). Build backend: `hatchling`. Layout: `src/excel_lsp/`.
- Package name `excel-lsp` (import `excel_lsp`, CLI `excel-lsp`). Check PyPI name availability in Phase 0; if taken, fall back to `excel-lsp-mcp` and log it. README footer: "Not affiliated with Microsoft. Excel is a trademark of Microsoft Corporation."
- License: MIT.
- Dependencies (runtime): `openpyxl>=3.1` — **pin the exact version chosen at P0** in `pyproject.toml`/`uv.lock` and record it in the agent log; golden snapshots depend on generator determinism, and a minor openpyxl release changing default styles would break goldens across the whole CI matrix. Also: `mcp` (official Python SDK, use FastMCP), `pydantic>=2`, `typer` (CLI), `lxml`. Dev: `pytest`, `hypothesis`, `pytest-cov`, `ruff`, `pyright`, `tiktoken` (**pinned**), `matplotlib`. Platform extras group `[live]`: `pywin32; sys_platform == 'win32'`.
- Storage: stdlib `sqlite3`, one index DB per workbook at `<workbook>.xlsp.db` next to the file. When `EXCEL_LSP_INDEX_DIR` is set, the filename is `<stem>.<first 8 hex of sha256(absolute path)>.xlsp.db` so two workbooks with the same basename cannot collide. Open every connection with WAL mode, `busy_timeout=5000`, `synchronous=NORMAL` — multiple server instances pointed at one workbook must degrade gracefully, not corrupt. Suggested `.gitignore` snippet in docs: `*.xlsp.db*` (covers the `-wal`/`-shm` sidecars). Docs note: next-to-file placement churns cloud-synced folders (OneDrive/Dropbox); `EXCEL_LSP_INDEX_DIR` is the remedy.
- SQLite R\*Tree for range queries. **VERIFY in Phase 0:** `CREATE VIRTUAL TABLE t USING rtree(id, minX, maxX)` succeeds in this Python build. If not, implement the documented fallback: plain table with indexed `(sheet_id, row_min, row_max, col_min, col_max)` and interval-overlap SQL. Keep both behind one `EdgeStore` interface either way.
- Formatting/linting: `ruff` (format + lint). Types: `pyright` in `basic` mode minimum; upgrade to `strict` for `core/` if it doesn't cost more than an hour.

**Parser architecture (v1.1 decision — log the rationale).** The primary index path is a **single-pass lxml parser** over the OOXML package parts, specified in §5.1. openpyxl is retained for exactly three jobs: (1) **fixture generation** — creating new files from scratch is safe, there is nothing to lose; (2) **formula utilities** — `openpyxl.formula.tokenizer.Tokenizer` and `openpyxl.formula.translate.Translator`; (3) the **cross-validation oracle** — a dual-load harness (`load_workbook(read_only=True, data_only=False)` for formulas + `data_only=True` for cached values, zipped per cell) lives under `tests/oracle/`, and every generated fixture asserts the lxml-derived `(ref, value, formula)` stream equals the openpyxl-derived one, modulo a documented skip-list (§8.1a). Rationale: openpyxl read-only mode has unverified shared-formula expansion and most likely does not populate `ws.tables` or `merged_cells` (both required by §5.2); the dual load was the S1 bottleneck; date/dtype classification requires `styles.xml` parsing we must own regardless; and the editor (§5.7) needs part-level zip machinery anyway, so the parser and editor share infrastructure.

**VERIFY (Phase 1, informational — calibrates the oracle skip-list, no longer load-bearing):** record actual openpyxl read-only behavior for (a) shared-formula follower cells (translated text vs. empty), (b) `ws.tables`, (c) `merged_cells.ranges`, in the agent log against fixture F07.

**Stored-name normalization (applies everywhere):** post-2007 functions are stored in file XML as `_xlfn.XLOOKUP(...)`; worksheet-scoped ones as `_xlfn._xlws.FILTER(...)`. Strip these prefixes for display and for function-name matching; preserve the stored formula text verbatim in the index. Excel-authored files (the live fixtures L1–L3) *will* contain them.

---

## 3. Architecture

Three layers, strictly separated:

```
src/excel_lsp/
  core/        # pure library. parse/ (lxml package parser + formula utils),
               # index/, graph/, diagnostics/, edit/ (surgical writer).
               # No MCP; no I/O beyond file + sqlite.
  server/      # MCP stdio server wrapping core. Tool schemas, response shaping,
               # caps, pagination, annotations, instructions, progress.
  cli/         # typer CLI: serve, map, trace, path, diag, find, schema, graph, bench.
```

`core` must be importable and fully usable without the server (this is a selling point: embed the index in any agent). The CLI exists so *you* can debug without MCP plumbing and so the README demo is copy-pasteable.

---

## 4. Data model

### 4.1 SQLite schema (per workbook)

```sql
CREATE TABLE meta        (key TEXT PRIMARY KEY, value TEXT);
-- keys: schema_version, workbook_path, workbook_hash, indexed_at,
--       mtime_ns, size, generation, date1904
CREATE TABLE sheets      (id INTEGER PRIMARY KEY, name TEXT UNIQUE, xml_part TEXT,
                          part_hash TEXT, kind TEXT,                 -- 'worksheet'|'chartsheet'|'macro'|'dialog'
                          visibility TEXT,                           -- 'visible'|'hidden'|'veryHidden'
                          max_row INTEGER, max_col INTEGER);
CREATE TABLE regions     (id INTEGER PRIMARY KEY, sheet_id INT, n INT,     -- n = ordinal by (top,left)
                          row_min INT, row_max INT, col_min INT, col_max INT,
                          header_rows INT, kind TEXT,                       -- 'table'|'region'
                          list_object_name TEXT, confidence REAL);
CREATE TABLE columns     (id INTEGER PRIMARY KEY, region_id INT, idx INT, header TEXT,
                          norm_header TEXT, dtype TEXT,                     -- 'int'|'float'|'date'|'str'|'bool'|'mixed'|'empty'
                          nonnull INT, distinct_est INT, formula_block_id INT NULL);
CREATE TABLE fblocks     (id INTEGER PRIMARY KEY, sheet_id INT, n INT, r1c1 TEXT,
                          row_min INT, row_max INT, col_min INT, col_max INT,
                          volatile INTEGER, opaque INTEGER);                -- opaque: contains INDIRECT/OFFSET-class refs
CREATE TABLE defined_names (id INTEGER PRIMARY KEY, name TEXT, scope_sheet_id INT NULL,
                          refers_to TEXT,
                          kind TEXT,                     -- 'range'|'multi_range'|'constant'|'formula'|'lambda'
                          is_builtin INTEGER);
CREATE TABLE name_areas  (id INTEGER PRIMARY KEY, name_id INT, sheet_id INT,
                          row_min INT, row_max INT, col_min INT, col_max INT);
                          -- one row per area; kind 'range' has one, 'multi_range' several,
                          -- 'constant'/'formula'/'lambda' have zero
CREATE TABLE validations (id INTEGER PRIMARY KEY, sheet_id INT,
                          row_min INT, row_max INT, col_min INT, col_max INT,
                          vtype TEXT, operator TEXT, formula1 TEXT, formula2 TEXT,
                          allow_blank INTEGER);
CREATE TABLE edges       (id INTEGER PRIMARY KEY, src_kind TEXT, src_id INT,  -- src: fblock or single cell (kind 'cell', id row<<16|col packed + sheet in src_sheet)
                          src_sheet_id INT, dst_sheet_id INT NULL,            -- NULL dst_sheet => external/opaque
                          dst_row_min INT, dst_row_max INT, dst_col_min INT, dst_col_max INT,
                          via TEXT);   -- 'ref'|'name:<n>'|'name-relative'|'structured:<table[col]>'|'3d'
                                       -- |'spill'|'external:<target>'|'opaque:<fn>'
CREATE VIRTUAL TABLE edge_rtree USING rtree(edge_id, sheet_min, sheet_max, row_min, row_max, col_min, col_max);
CREATE TABLE diagnostics (id INTEGER PRIMARY KEY, severity TEXT, code TEXT, sheet_id INT,
                          row INT NULL, col INT NULL, ref TEXT, message TEXT, related TEXT);
CREATE TABLE staleness   (sheet_id INT, row_min INT, row_max INT, col_min INT, col_max INT, since TEXT);
```

Notes: `edge_rtree` mirrors the *destination* rectangle of each edge so "who reads cell X / range R" is one spatial query (sheet id packed as the first dimension pair, `sheet_min = sheet_max = sheet_id`). R\*Tree supports up to 5 dimension pairs; we use 3. Whole-column refs store `row_min=1, row_max=1048576`; whole-row refs `col_min=1, col_max=16384` (coordinates ≤ 2^20 are exactly representable in the R\*Tree's float32 storage, so no precision issue). `meta.generation` is a monotonically increasing counter bumped on **every** index mutation — any reindex, incremental patch, or write — and is embedded in pagination cursors (§6.1). `meta.mtime_ns`/`size` back the per-call freshness check (§5.8).

### 4.2 Symbol ID scheme (frozen)

```
sheet:{sheetName}
region:{sheetName}:{n}            # n ordered by (row_min, col_min), 0-based
col:{sheetName}:{n}:{normHeader}[#k]   # #k suffix only for duplicate headers
name:{definedName}                 # global scope
name:{sheetName}!{definedName}     # sheet scope
fblock:{sheetName}:{n}
cell:{sheetName}!{A1}
```

Stable across reindexes when content is unchanged. Document that region ids can shift if regions are added/moved (acceptable; the map is cheap to re-fetch). Content-addressed stable aliases are a roadmap idea, not v1 — the scheme above stays frozen.

### 4.3 The workbook map (the flagship output)

Compact JSON, budget ≤ 1,500 tokens on the reference model (S2). Shape:

```json
{
  "workbook": "model.xlsx", "sheets": 3, "indexed_at": "...", "stale": false, "hasVBA": false,
  "sheetList": [
    {"sheet": "Inputs", "dims": "A1:D40", "regions": [
      {"id": "region:Inputs:0", "range": "A1:D38", "kind": "region", "headerRows": 1,
       "cols": [{"h": "Rate", "t": "float"}, {"h": "Volume", "t": "int"}],
       "rows": 37, "fblocks": 0, "conf": 0.94}
    ]},
    {"sheet": "Calc", "dims": "A1:H50001", "regions": [
      {"id": "region:Calc:0", "...": "...", "fblocks": 6},
      {"more": 12}
    ]},
    {"sheet": "Scratch", "vis": "hidden", "dims": "A1:B9", "regions": []}
  ],
  "names": [{"id": "name:TaxRate", "ref": "Inputs!$B$4"}],
  "namesMore": 0,
  "externalLinks": ["[budget2025.xlsx]"],
  "diagCounts": {"error": 3, "warn": 5},
  "hints": ["Use get_region_schema for columns+samples", "trace_dependents for impact analysis"]
}
```

No cell values in the map (headers and defined-name refs are the only content strings). Include the `hints` array — it teaches the calling agent the intended navigation flow. `vis` is included only when not `visible` (hidden and veryHidden sheets are prime audit material and must be surfaced, never silently omitted). `hasVBA` reflects the presence of `xl/vbaProject.bin`.

**Degradation rules (deterministic — the map must never blow the cap):** per sheet, include at most 8 regions ordered by area descending, then a `{"more": N}` entry; `names` capped at 20 entries with the remainder in `namesMore` (hint: `list_symbols`); `externalLinks` capped at 10 with a count. The serialized map must fit the 8,000-character response cap on **every** fixture including F20 (test-enforced); S2's 1,500-token budget applies to the reference model F03. The map never paginates — that would defeat its purpose; it summarizes, and the tools drill down.

---

## 5. Module specifications (core)

Each module below lists algorithm, invariants (the mechanics reviewer checks these — §11.5), and edge cases.

### 5.1 M1 — Loader, hasher & package parser (single pass, lxml)

An `.xlsx` is a zip of XML parts. M1 owns the zip and the parse; everything downstream consumes its streams.

**Parts parsed** (namespace-tolerant throughout — LibreOffice/Google exports must survive):

- `xl/workbook.xml`: sheet list (name, sheetId, r:id, `state` → visibility), `<definedNames>` (name, `localSheetId` scope, refers-to text), `<calcPr>`, `<workbookPr date1904>`.
- `xl/_rels/workbook.xml.rels`: r:id → part path. Classify each sheet as worksheet / chartsheet / macro / dialog by part path and content type; non-worksheets get `sheets` rows with `kind` set and **no cell parse** — they must never crash the indexer.
- `xl/sharedStrings.xml`: `<si>` → text; concatenate rich-text runs (`<r><t>`); honor `xml:space="preserve"`.
- `xl/styles.xml`: custom `<numFmts>`, `<cellXfs>` (xf → numFmtId, fontId, fillId), `<fonts>` (bold flag), `<fills>`. **Date detection:** a cell is date-typed when its xf's numFmtId is in the built-in date/time set {14–22, 27–36, 45–47, 50–58} (cross-check the list against openpyxl's builtins at implementation time) OR its custom format code contains any of the tokens `y m d h s` *outside* quoted sections and outside `[...]` color/condition blocks. (`m` is ambiguous month/minute — either way it is a date/time format for dtype purposes.)
- Each `xl/worksheets/sheetN.xml`, streamed with `lxml.etree.iterparse` and element clearing (never a full DOM for big sheets): `<dimension>` (advisory only — **never trust it**; many writers omit or lie), `<sheetData>` rows and cells — attributes `r` (ref), `s` (style idx), `t` (absent/`n` | `s` | `str` | `b` | `e` | `inlineStr`); children `<v>`, `<is>`, `<f>` including `t="shared"` (`si`, `ref` on the master) and `t="array"` (`ref` span) — plus `<mergeCells>`, `<dataValidations>`, and `<tableParts>` → sheet rels → `xl/tables/table*.xml` (name/displayName, ref, headerRowCount, totalsRowCount, tableColumns).
- `xl/externalLinks/externalLink*.xml` + their rels: build the **`[n]` index → target path/URL map** consumed by `classify_ref` (§5.3) and `E_BROKEN_XLINK` (§5.6).
- Presence of `xl/vbaProject.bin` → map `hasVBA`.

**Cell record emission** — one streamed `(ref, value, value_type, formula|None, style_idx)` per non-empty cell:

- Value typing: `t` absent/`n` → number (int when integral, else float); `t="s"` → sharedStrings[i]; `t="str"` → cached string result of a formula; `t="b"` → bool; `t="e"` → error text (value_type `error`); `t="inlineStr"` → inline text. Date-formatted numerics → converted to datetime, honoring `date1904` and the 1900 leap-year quirk (serial 60) — **use openpyxl's date utilities as the reference conversion; do not hand-roll it.**
- **Value normalization (one function, used by every tool response, golden fixture, and sample):** JSON scalars only — dates as ISO-8601 strings, errors as their `#...!` text, booleans as true/false, numbers as numbers.
- **Shared formulas:** the master carries full text + `ref` + `si`; followers carry `si` only. Derive each follower's text with `Translator(master_text, origin=master_cell).translate_formula(follower_cell)`. This is the same machinery as the §5.4 block cross-check — share it.
- **Array formulas:** record the master's `t="array"` + `ref` span; member cells are attributed to the span (the editor's `E_ARRAY_FORMULA` check in §5.7 needs this).

**Hashing & incremental (unchanged in substance from v1.0):** on open, read the zip central directory; SHA-256 per part for `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `xl/sharedStrings.xml`, **`xl/styles.xml`** (added — style changes can flip date dtypes), and each sheet part; whole-file hash too. Incremental reindex compares part hashes to `sheets.part_hash` and reindexes only changed sheets, with a full names/links refresh if `workbook.xml` changed. A `sharedStrings.xml` hash change reindexes the changed sheets (strings are referenced by index; log if it ever changes alone). A `styles.xml` change re-derives dtypes for all sheets (cheap — dtype sampling only). Record `mtime_ns` + `size` in meta for the freshness fast path (§5.8).

- Invariants: **I1** reindex of an untouched workbook is a no-op (hashes short-circuit). **I2 (re-specified so it is actually testable):** per-sheet reindex and full reindex produce identical **canonical exports** for unchanged sheets — canonical export = every table ORDER BY its natural key with surrogate id columns projected out (raw rows with autoincrement ids can never be byte-identical). **I3** the loader never mutates the source file.
- Edge cases: workbook with 0 formulas; LibreOffice/Google namespace quirks (tolerate); password-protected file → structured error `E_ENCRYPTED`; file locked by Excel (Windows share violation) → retry once after 500 ms then `E_LOCKED`; **mid-save torn zip** (freshness check can race a writer) → retry once then `E_CORRUPT`; chartsheets/macro sheets (rows in `sheets`, no cells); inlineStr cells; rich-text shared strings; `date1904`; missing/false `<dimension>`.
- Oracle: every generated fixture's lxml stream must equal the openpyxl dual-load stream — see §8.1a.

### 5.2 M2 — Region detection & headers

- Build per-row non-empty column intervals while streaming (do not materialize a dense grid). Merge intervals vertically into rectangles, tolerating gaps ≤ `gap_tol` (default 1 blank row/col; configurable).
- Native Excel Tables (**from the M1 parser's tableParts output** — do not rely on openpyxl worksheet APIs) always win: emit them as `kind:"table"` with their declared range and header row; heuristic regions never overlap a ListObject.
- Header inference on candidate top rows (check up to 3): score = weighted features — top row mostly strings while body columns are typed; header values unique; format shift (bold/fill) between row and body. Style features come from the parsed `styles.xml` tables via each candidate cell's `s` index — resolve lazily, for candidate header rows only. Emit `header_rows` (0–3) and `confidence` 0–1. Weights are yours to tune; log them.
- Column dtype: sample up to 200 body cells per column; classify int/float/date/str/bool/mixed/empty; count nonnull; estimate distinct via a small hash set capped at 1,000.
- Merged cells (**from the M1 parser's `<mergeCells>` output**): value lives in the top-left; for multi-row merged headers, synthesize the header as the join of merged texts down the header rows ("Revenue / Q1").
- Caps: > 2,000,000 non-empty cells in one sheet → index structure fully but sample dtypes at lower rate and emit `warn:W_LARGE_SHEET`.
- Invariants: **I4** every non-empty cell belongs to ≤ 1 region; **I5** ListObject ranges are reproduced exactly; **I6** region ordinals are deterministic for identical content.

### 5.3 M3 — Formula parsing & reference extraction

- Tokenize with `openpyxl.formula.tokenizer.Tokenizer`. Function names are matched **after stripping `_xlfn.` / `_xlfn._xlws.` prefixes** (§2). Maintain a frozen built-in-function set: **VERIFY** the import `from openpyxl.utils import FORMULAE` (the path may differ across versions); if unavailable, vendor a static list at `core/parse/functions.py` (generated once, committed, provenance noted in a comment).
- For each OPERAND/RANGE token, classify the reference text with a dedicated `classify_ref()`:
  - Plain: `A1`, `$B$2`, `C3:D9`, `B:B`, `7:7`
  - Sheet-qualified: `Sheet2!A1`, `'My Sheet'!C3:D4` (quoted names, embedded apostrophes doubled)
  - 3-D: `Jan:Mar!B2` → expand to one edge per sheet in the span (workbook sheet order). Include the quoted-span form `'Jan 24:Mar 24'!B2` — both endpoints inside one quote pair — in the matrix and tests.
  - Defined name (operand position): bare identifier not matching A1/R1C1 grammar and not a function call → look up in `defined_names` (sheet scope first, then global). Unresolvable → edge `via='opaque:name'` + `warn:W_UNKNOWN_NAME`.
  - **Defined name (function-call position):** an identifier in call position that is **not** a built-in function is also looked up in `defined_names` — this is how LAMBDA names are invoked. Resolved → edges `via='name:<n>'` to the name's areas; for `kind='formula'|'lambda'` names, additionally parse the refs inside `refers_to`: absolute refs resolve to normal edges; **relative refs inside name definitions are context-dependent** (relative to the cell of use) → emit `via='name-relative'` opaque, no destination rect.
  - **LET / LAMBDA parameter suppression:** track function nesting with a small stack keyed on tokenizer FUNC-open/close and argument-separator tokens. Inside `LET`, the odd-positioned arguments before the final one are binding names; inside `LAMBDA`, all arguments before the final one are parameters. Identifiers matching an in-scope binding/parameter are **skipped entirely** — no lookup, no warning, no edge. Without this, every LET in a modern workbook emits a spurious `W_UNKNOWN_NAME`. F19 asserts zero spurious warnings (I20).
  - **Spill reference** `A1#` (also `Name#`): edge to the anchor cell (or the name's areas) with `via='spill'`. The spilled extent is dynamic; the anchor is the definition point, which is the correct go-to-definition answer. No diagnostic.
  - **Implicit intersection** `@ref` outside structured refs: strip the `@`, classify the inner ref normally.
  - Structured: `Table1[Col]`, `Table1[@Col]`, `Table1[[#Totals],[Col]]`, bare `[@Col]` (implicit current table). Support `#All/#Data/#Headers/#Totals/@`; anything else → opaque + warn.
  - External: `[1]Sheet1!A1` / `'[budget2025.xlsx]Q1'!A1` → resolve `[n]` through the M1 externalLinks map; edge with `dst_sheet_id NULL`, `via='external:<target>'`.
- Dynamic-reference functions (`INDIRECT`, `OFFSET`, `CHOOSE` over refs, `INDEX` used as a ref lhs): do **not** pretend to resolve. Emit `via='opaque:<FN>'` edge with no destination rectangle, mark the fblock `opaque=1`, emit `info:I_DYNAMIC_REF`. Volatile functions (`NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `RANDARRAY`, `OFFSET`, `INDIRECT`, `CELL`, `INFO`) mark the block `volatile=1`.
- Invariants: **I7** every formula cell yields ≥ 0 edges and never crashes the indexer (unparseable formula → `warn:W_PARSE` + opaque edge, keep going); **I8** quoted sheet names round-trip; **I20** on F19 — no `W_UNKNOWN_NAME` or `W_PARSE` for LET bindings, LAMBDA parameters, `_xlfn.`-prefixed calls, spill refs, or `@` intersections.

### 5.4 M4 — R1C1 normalization & formula blocks

- Implement `to_r1c1(formula: str, anchor: CellRef) -> str` using the tokenizer: rewrite each cell part relative to the anchor, honoring `$` (absolute stays `R4C2`-style, relative becomes `R[+1]C[-2]`). Non-reference tokens — including structured refs, names, and spill refs — pass through verbatim.
- Block detection: scan each sheet column-major and row-major; contiguous runs of cells whose `to_r1c1` strings are identical collapse into one `fblock` (rectangles: grow runs down, then merge horizontally-adjacent identical columns into rectangles). Parse references **once per block** using the anchor; edge rectangles for relative refs are the anchor's ref rectangle *extruded* by the block's extent (e.g., block B2:B50001 with `=RC[-1]*R1C4` yields a relative edge A2:A50001 and an absolute edge D1:D1). **Clamp extruded rectangles to sheet bounds** — a block starting at row 2 with `R[-1]C` must not emit row 0 (the corresponding real cells would be `#REF!` and are caught by E_ERRVAL, but the edge math must stay in-range).
- Cross-check for tests: `openpyxl.formula.translate.Translator(f, origin).translate_formula(target)` — two cells belong to the same block iff translating the first cell's formula to the second's coordinate reproduces it exactly.
- **Inconsistency diagnostic (re-specified — the v1.0 definition could never fire, because fblocks partition formula cells by identical R1C1, so no differing cell can lie inside a block's own rectangle):** run over each maximal contiguous vertical run of formula cells per column within a region, plus a symmetric horizontal pass per row; de-duplicate cells flagged by both. For a run of n ≥ 5 formula cells, let d = the count of the most common R1C1 string in the run. If `d/n ≥ 0.8` and `(n − d) ≤ max(3, ceil(0.05·n))`, flag each minority cell with `warn:W_INCONSISTENT_FORMULA`, `related` = {dominant fblock id, expected R1C1}. Thresholds are defaults you may tune (log it), but F07's planted tamper and benchmark task B2 must be caught with the defaults.
- Invariants: **I9** every formula cell belongs to exactly one fblock (singleton blocks allowed); **I10** `to_r1c1(f, c1) == to_r1c1(Translator-shift(f, c1→c2), c2)` (hypothesis property test); **I11** union of fblock rectangles covers exactly the set of formula cells.

### 5.5 M5 — Dependency graph & queries

- Edges as in §4.1, source granularity = fblock (or singleton cell). Precedents(query rect): edges whose *source* block intersects the rect → their destinations. Dependents(query rect): R\*Tree lookup of edges whose destination rect intersects the query rect → their source blocks. Both support `depth` (default 2, max 8) BFS with cycle guard and `max_nodes` truncation (return `truncated: true`).
- Trace output is a tree of `{symbol/ref, via, children}` with counts, not values.
- **Path query** (backs the `trace_path` tool, §6.2): up to `max_paths` (default 3) block-level paths from a source to a destination over dependent edges, depth ≤ `max_paths_depth` (default 12), found by BFS with per-node parent sets; each path is a list of `{symbol, via}`. No paths → `{"connected": false, "paths": []}` (not an error).
- Circular reference detection, two-stage (important for correctness): **(1)** Tarjan SCC on the condensed block-level graph. **(2)** Any SCC involving self-overlapping *range* edges (e.g., running total `=SUM(B$2:B2)` whose destination overlaps its own block) must be re-verified at cell level *within that SCC only* before reporting `error:E_CIRCULAR`. The cell-level stage is specified to avoid the O(n²) trap on 50k-row running totals:
  - **Stage 2a — self-inclusion test, O(cells in SCC):** for every cell c in the SCC's blocks, translate the block's R1C1 to c and test whether any destination interval **includes c itself**. Any hit → true `E_CIRCULAR` at that cell.
  - **Stage 2b — bounded expansion (multi-block SCCs with no self-inclusion hits):** cell-level BFS from ≤ 64 seed cells (block corners + evenly spaced samples), visiting ≤ 100,000 cells total. Cycle found → `E_CIRCULAR` with the path in `related`. Bound exceeded → **`warn:W_POSSIBLE_CIRCULAR`** ("verify in Excel") — never spend quadratic time proving a running total innocent.
  - Fixture contract: F09a (small true cycle) yields exactly `E_CIRCULAR`; F09b (running total) yields **nothing** — no error, no warn.
- Invariants: **I12** dependents(precedents(x)) ⊇ {x's block} on fixtures; **I13** no query allocates the full cell-level graph for the 50k fixture; **I14** whole-column refs match point queries anywhere in that column.

### 5.6 M6 — Diagnostics

`E_ERRVAL` detection keys on the cell's **error type** (XML `t="e"`), never on a text whitelist — any `t="e"` cached value is `E_ERRVAL` regardless of whether the text is recognized. The recognized-text list (`#REF!`, `#DIV/0!`, `#N/A`, `#VALUE!`, `#NAME?`, `#NUM!`, `#SPILL!`, `#CALC!`, `#BLOCKED!`) exists only to prettify messages.

Codes (severity): `E_ERRVAL`, `E_CIRCULAR`, `E_BROKEN_XLINK` (external link target missing on disk — resolve through the M1 externalLinks map relative to the workbook dir; can't check remote targets → `warn`), `W_POSSIBLE_CIRCULAR` (§5.5), `W_INCONSISTENT_FORMULA`, `W_UNKNOWN_NAME`, `W_PARSE`, `W_LARGE_SHEET`, `W_REGEX_TIMEOUT` (§6.1), `I_DYNAMIC_REF`, `I_VOLATILE`, `I_STALE` (see M7). Each diagnostic carries `ref`, `message`, and `related` (e.g., the block id, the expected R1C1). `get_diagnostics` filters by sheet/severity/code.

### 5.7 M7 — Editor (surgical XML patching) & staleness

**Why (frozen rationale — this is the part-preservation rule referenced in §0.4):** openpyxl's load/save drops charts, images, drawings, and other unmodeled content from existing files — *silently, with no repair dialog*, so a live pass keyed on repair dialogs alone would certify an editor that quietly mutilates user workbooks. A normal-mode load of the 50k fixture is also slow and memory-heavy (~1 KB/cell). Therefore the editor **never round-trips a workbook through openpyxl**. It rewrites the zip, stream-copying every untouched part byte-for-byte, and patches only:

1. The target sheet part(s).
2. `xl/workbook.xml` — ensure `<calcPr fullCalcOnLoad="1"/>` exists (create `<calcPr>` if absent). This replaces v1.0's openpyxl `wb.calculation` VERIFY: we set the XML directly.
3. **Delete `xl/calcChain.xml`**, remove its relationship from `xl/_rels/workbook.xml.rels`, and remove its `Override` from `[Content_Types].xml`. The calc chain is a recalculation-order cache; Excel rebuilds it on load. A dangling rel to a deleted part risks a repair dialog, so all three edits go together, always.
4. Nothing else. `xl/sharedStrings.xml` is never touched: written strings use inline strings (below).

**Cell mechanics (target sheet XML):** rows in `<sheetData>` are ordered by `r`; cells within a row are ordered by column. Locate or insert **in order** (Excel tolerates disorder poorly). Preserve an existing cell's `s` (style) attribute verbatim; new cells carry no `s` (default style). Value writes:

- number → `<v>` with no `t` (integers rendered without a decimal point);
- boolean → `t="b"`, `<v>0|1</v>`;
- string → **`t="inlineStr"`** with `<is><t xml:space="preserve">...</t></is>`. Inline strings avoid mutating `sharedStrings.xml`; Excel and LibreOffice read them natively. Documented as a limitation (some third-party tools handle inline strings poorly).
- null → remove `<v>`/`<f>`/`<is>`, leaving a styled empty cell;
- datetime → **`E_INVALID_VALUE`** in v1 (hint: write the serial number or an ISO string). Correct date writes require number-format surgery on `styles.xml`; that is a roadmap item, not a silent wrong answer.

Formula writes: set `<f>` to the formula text minus the leading `=`, and **remove any stale `<v>`** — the cell then reads as valueless until Excel recalculates, which the staleness system already models honestly. Writing a value over a formula removes `<f>` first.

**Shared formulas:** if the target cell belongs to a shared group (master or follower), first **expand the entire group** to explicit per-cell `<f>` elements via Translator, then apply the edit. Editing a shared master in place while followers still point at its `si` corrupts the group. **Array formulas:** a target inside a multi-cell `t="array"` span → **`E_ARRAY_FORMULA`** (mirrors Excel's own restriction; hint: rewrite the whole array or use a dynamic-array formula in the anchor).

**`<dimension>`:** recompute when written cells extend the extents; otherwise drop the element entirely — it is optional and advisory.

**Preconditions, in order:**

1. **Lockfile check:** a sibling `~$<basename>` file means the workbook is open in Excel → **`E_OPEN_IN_EXCEL`** (hint: close it in Excel first; note that stale lockfiles can survive Excel crashes and may be deleted manually). This matters most on macOS, where Excel takes no OS-level lock and a "successful" write would be silently clobbered by Excel's next save.
2. **Conflict check:** re-hash the file; if it changed since index time → **`E_CONFLICT`**, instruct the caller to `refresh` first.

Write to a temp file in the same directory, fsync, then atomic-replace (unchanged from v1.0).

`write_cells` (list of `{ref, value|formula}`, ≤ 500) and `set_column_formula` (a `col:` symbol + an R1C1 or A1-with-anchor pattern; fills the column's body range, creating/replacing the fblock) sit on these mechanics; contracts in §6.2. Edits are allowed on **all** supported extensions including macro-enabled files — the VBA project survives by construction (F16 byte-compares `xl/vbaProject.bin`).

**Post-write:** patch the changed cells in the index directly (no full reindex); bump `meta.generation`; then staleness propagation exactly as before — BFS over dependents of the written rects (block granularity, transitive, capped at 50k blocks), insert `staleness` rects. Any tool that returns values intersecting a stale rect adds `"stale": true` + `I_STALE`. Staleness clears for a sheet when a later reindex sees a changed part hash *from an external save* (i.e., Excel recalculated and saved) or when the caller passes `recalculated=true` to `refresh`.

- Invariants: **I15** an edit through Excel LSP followed by open-in-Excel produces no repair dialog (live test); **I16** after `write_cells`, dependents' staleness is set and reads reflect it; **I17** *(superseded by I18, kept as its corollary)* the editor never drops sheets, names, styles, or the VBA project it didn't touch; **I18** after any edit, **every zip part not deliberately modified is byte-identical to the source** — enforced by part-diff tests on F16 (`xl/vbaProject.bin`) and F21 (chart XML + media parts) and by a hypothesis property (§8.2).

### 5.8 M8 — Index lifecycle

`open(path)` → hash check → full or incremental reindex → map. `refresh(path)` → same, explicit. **Freshness check on every tool call:** stat the workbook (`mtime_ns`, `size`) against meta; on mismatch, hash parts and run the incremental reindex *before* serving, and include `"reindexed": true` in the response. This is the LSP-like behavior calling agents assume — the index silently tracks the file; nobody should have to remember to call `refresh` after saving in Excel. Schema version lives in `meta`; on mismatch, drop and rebuild the DB (the index is always derivable). `meta.generation` bumps on every index mutation. All core APIs are synchronous and pure w.r.t. the network.

---

## 6. MCP tool contracts

### 6.1 Global rules (testable properties)

- **No bulk data:** no tool response may contain more than 200 raw cell values, ever. `read_range` is the only tool that returns values in quantity and it enforces the cap per page.
- **Response cap:** serialized response ≤ 8,000 characters. On overflow: truncate deterministically, set `"truncated": true`, include `"cursor"` for continuation.
- **Cursors carry the index generation:** cursor = opaque base64 of `{tool, params_hash, offset, gen}`, where `gen` is `meta.generation` at issue time. Any index mutation (reindex or write) between pages invalidates outstanding cursors → **`E_STALE_CURSOR`** (hint: re-issue the original query). Silently continuing pagination over a mutated index would hand the agent a corrupt composite view.
- **Freshness:** every tool call performs the §5.8 stat check; responses include `"reindexed": true` when it fired. Agents assume the index tracks the file, like an LSP tracks a buffer.
- Every response includes `"stale"` where values are involved, and `"nextSteps"` hints where natural (map → schema → trace).
- Structured errors: `{"error": {"code": "E_...", "message": "...", "hint": "..."}}` — never a bare exception string, never a traceback.
- Server is read-only on the filesystem except through the two write tools and the index DB.
- **Tool annotations:** the 12 read tools declare `readOnlyHint: true, openWorldHint: false`; `write_cells` and `set_column_formula` declare `readOnlyHint: false, destructiveHint: true`. MCP clients gate confirmation UX on these; without them every client treats every call as potentially destructive.
- **Server instructions:** provide the MCP `instructions` field at initialize (~150 tokens): the intended navigation flow (open_workbook → get_region_schema → trace_* → read_range last and least), the no-bulk philosophy, and the stale/reindexed semantics. This teaches every connected client once; the per-response `hints` remain as reinforcement.
- **Progress:** `open_workbook`/`refresh` report per-sheet progress via the client's progress token when one is provided (FastMCP context); no-op otherwise. A silent 10-second cold index on a big workbook is exactly where client timeouts and user abandonment live.
- **Security scope:** optional `EXCEL_LSP_ROOT` — an `os.pathsep`-separated allowlist of directories. When set, every path argument is realpath-resolved (symlinks followed) and must fall under an allowed root, else **`E_PATH_DENIED`**. Default is unrestricted; the README's Security note and `SECURITY.md` state the scope plainly. An agent-facing server that reads and writes arbitrary local files owes its users one honest paragraph about that.
- **Regex guard (`find`):** pattern length ≤ 256 chars; each cell's match subject truncated to 1,000 chars; a 2-second deadline checked between cells — on expiry, return partial results with `truncated: true` and a `W_REGEX_TIMEOUT` note in the payload. Agent-supplied patterns must not be able to hang the server (catastrophic backtracking is one `(a+)+b` away).

**Canonical tool-error codes** (distinct from workbook diagnostics, §5.6; T6 exercises each at least once; the shared table lives in `core/errors.py`):

| Code | Meaning |
|---|---|
| `E_NOT_FOUND` | path does not exist |
| `E_UNSUPPORTED_FORMAT` | `.xls`/`.xlsb`/other unsupported extension |
| `E_ENCRYPTED` | password-protected workbook |
| `E_LOCKED` | OS share violation persists after one retry |
| `E_OPEN_IN_EXCEL` | `~$` lockfile present (write tools only) |
| `E_CONFLICT` | file changed since index time; `refresh` first (write tools only) |
| `E_CORRUPT` | zip/XML unreadable after one retry (e.g., torn mid-save read) |
| `E_STALE_CURSOR` | cursor generation mismatch |
| `E_INVALID_REF` | unparseable ref/range/symbol syntax |
| `E_UNKNOWN_SYMBOL` | symbol id not present in the index |
| `E_ARRAY_FORMULA` | edit targets a cell inside a multi-cell array formula |
| `E_INVALID_VALUE` | unsupported write value (e.g., datetime in v1) |
| `E_PATH_DENIED` | path outside `EXCEL_LSP_ROOT` |
| `E_INTERNAL` | sanitized unexpected failure (message logged server-side, never a traceback to the client) |

### 6.2 Tools (14)

| Tool | Input (types, defaults) | Output (keys) |
|---|---|---|
| `open_workbook` | `path` | the workbook map (§4.3) |
| `refresh` | `path`, `recalculated: bool=false` | map + `reindexedSheets` |
| `list_symbols` | `path`, `query: str=""`, `kinds: [str]=all` | matching symbol ids + one-line descriptors |
| `get_region_schema` | `path`, `region_id` | headers, dtypes, nonnull/distinct, sample rows (see cap below), per-column data-validation summaries when present, fblock summaries, `conf` |
| `read_range` | `path`, `ref`, `cursor=None`, `max_cells: int=200` | 2-D values + `truncated/cursor/stale` |
| `find` | `path`, `pattern` (regex, §6.1 guard), `in: [values,headers,formulas,names]`, `sheet=None`, `max=50` | list of `{ref, kind, snippet≤80ch}` |
| `trace_precedents` | `path`, `ref_or_symbol`, `depth: int=2 (≤8)`, `max_nodes=200` | tree (§5.5) |
| `trace_dependents` | same | tree |
| `trace_path` | `path`, `from_ref_or_symbol`, `to_ref_or_symbol`, `max_paths: int=3`, `max_depth: int=12` | up to `max_paths` block-level paths, each a list of `{symbol, via}`; `{"connected": false, "paths": []}` when none |
| `explain_formula` | `path`, `ref` | A1 + R1C1 forms, block id + extent, resolved names/structured refs, flags (volatile/opaque), diags on that cell |
| `get_diagnostics` | `path`, `sheet=None`, `severity=None`, `code=None`, `max=100` | list + counts |
| `profile_column` | `path`, `col_symbol_or_ref` | count/nonnull/sum/mean/min/max/distinct_est for numeric; top-5 values (truncated strings) otherwise; `cachedValues: false` flag + recalc hint when the file was saved without cached values |
| `write_cells` | `path`, `cells: [{ref, value?, formula?}] (≤500)` — value types per §5.7 (number/string/bool/null; datetime → `E_INVALID_VALUE`) | per-cell ok/err, `staleBlocks` count |
| `set_column_formula` | `path`, `col_symbol`, `formula` (A1 at anchor or R1C1), `overwrite: bool=false` | fblock id, cells written, `staleBlocks` |

`get_region_schema` sample cap: `sample_rows = max(0, min(3, 180 // ncols))` so samples can never breach the 200-value rule on wide tables (v1.0's flat "3 sample rows" broke it at ncols > 66). When 0, set `"samplesOmitted": "<ncols> columns; use read_range"`.

`trace_path` exists because "why does `Summary!C10` depend on `Inputs!B2`?" is the single most audit-shaped question agents ask, and reconstructing it from repeated `trace_*` calls burns turns; the graph already has the answer (§5.5).

Define each with pydantic models; FastMCP derives the JSON schema. One fully worked request/response example per tool — all 14 — goes in `docs/tool-reference.md`.

---

## 7. CLI

`excel-lsp serve` (stdio MCP), `excel-lsp map <file>`, `trace <file> <ref> --deps|--precs --depth`, `path <file> <from> <to>`, `diag <file>`, `find <file> <pattern>`, `schema <file> <region>`, `graph <file> <ref> --deps|--precs --depth N --mermaid` (emits a Mermaid flowchart of the trace — near-free atop trace output, and it makes the README demo pop), `bench` (runs §9). All read commands print the same JSON the MCP tools return (pretty-printed). This is your debugging surface and the README's proof-of-life.

---

## 8. Testing strategy

### 8.1 T1 — Unit (pytest)

Tokenizer/classify_ref over the full reference-grammar matrix **including the modern matrix** (`_xlfn.`/`_xlws.` prefixes, LET bindings, LAMBDA parameters, LAMBDA-name calls, `A1#` spill, `@ref` implicit intersection, quoted 3-D spans); `to_r1c1` known-answer tests; EdgeStore point/range/whole-column queries; header scorer on synthetic grids; hasher part-mapping; symbol id formatting; numFmt date classifier (builtin ids + custom-code heuristic incl. quote/bracket skipping); surgical-writer units by part-diffing zips (touched parts change exactly as specified, all others byte-identical; calcChain triple-edit; shared-group expansion; `E_ARRAY_FORMULA`; dimension recompute/drop); cursor generation invalidation; freshness stat path; lockfile precondition; error-table conformance (every code in §6.1 constructible and correctly shaped).

### 8.1a T-oracle — lxml vs. openpyxl cross-validation

For every generated fixture, the M1 lxml stream `(ref, value, formula)` must equal the openpyxl dual-load stream (`read_only=True`, `data_only` both ways, zipped per cell). Divergences are either lxml parser bugs (fix) or documented openpyxl deficiencies — maintain `tests/oracle/skiplist.md` recording each skip with its root cause (e.g., read-only shared-formula followers returning empty, if the P1 VERIFY shows that). Treat any *unexplained* divergence as a critical finding. This harness is what makes replacing openpyxl as the primary reader safe.

### 8.2 T2 — Property (hypothesis)

Random rectangular grids with planted regions → detection invariants I4–I6; random formulas from a small grammar → I10 round-trip vs. `Translator`; random edge sets → R\*Tree results equal brute-force interval scan; **random small workbooks + random edit scripts through the surgical writer → (a) every untouched part byte-identical (I18), (b) re-parsing yields exactly the written values/formulas**; adversarial regex patterns → the `find` guard returns within its deadline.

### 8.3 T3 — Fixture corpus (generator-built, committed as a script — one sanctioned binary exception below)

`tests/fixtures/generate.py` authors content with openpyxl (pinned, §2), then runs an **lxml post-pass** that makes the files behave like Excel-saved ones. CI regenerates all fixtures deterministically. Golden-file snapshots of maps/diagnostics/traces live in `tests/golden/` with token-count assertions. The post-pass has three jobs:

- **(a) Cached-value injection.** openpyxl never computes formulas, so without this every generated formula cell would carry no cached value — and `E_ERRVAL` detection, `profile_column`, dtype inference on formula columns, and benchmark B3 would all silently degrade to testing nothing. The generator knows every formula it plants and injects the expected `<v>` (numeric), `t="str"` + `<v>` (text results), or `t="e"` + `<v>#DIV/0!</v>` etc. (error cells). No general evaluator — the generator computes expected values in Python for the arithmetic it itself emits.
- **(b) Shared-formula groups.** Convert the designated fill-down columns in F07 into genuine `<f t="shared" ref si>` master + `si`-only followers (openpyxl writes full formulas per cell, so without this conversion no generated fixture would ever exercise the parser's shared-formula handling).
- **(c) F16 VBA injection.** Insert the committed blob `tests/fixtures/assets/vbaProject.bin` as `xl/vbaProject.bin`; add its `[Content_Types].xml` override (`application/vnd.ms-office.vbaProject`); add the workbook relationship; switch the workbook part's content type to the macroEnabled variant; emit as `.xlsm`. **The blob is the one sanctioned committed binary** — v1.0's "scripts, not binaries" rule collided with the fact that a VBA project cannot be scripted into existence. Author it once during P0 recon in desktop Excel: a module containing `Sub Stamp(): Range("Z1").Value = 42: End Sub` (used by the live pass's "macro still runs" assertion), save as `.xlsm`, extract the part, commit it with a provenance note in `tests/fixtures/README.md`. If Excel is absent at P0, pause per §0.5; F16-dependent tests skip-if-blob-missing until it exists.

| ID | File | Exercises |
|---|---|---|
| F01 | basic_single_table.xlsx | one region, clean headers |
| F02 | multi_region.xlsx | 3 islands on one sheet, gap tolerance |
| F03 | cross_sheet_model.xlsx | Inputs→Calc→Summary chains (the reference model for S2/S3 and benchmarks) |
| F04 | named_ranges.xlsx | global + sheet-scoped names, name used in formulas |
| F05 | structured_table.xlsx | ListObject, `Table1[Col]`, `[@Col]`, totals row |
| F06 | perf_50k.xlsx (+1k/10k variants) | S1 timing, fblock extrusion |
| F07 | formula_blocks.xlsx | fill-down blocks, one tampered cell (W_INCONSISTENT_FORMULA with default thresholds), genuine shared-formula group (post-pass b) |
| F08 | errors.xlsx | all `E_ERRVAL` variants incl. `#CALC!`/`#BLOCKED!` (post-pass a injects `t="e"` cached values) |
| F09a/b | circular.xlsx / running_total.xlsx | true cycle (exactly `E_CIRCULAR`) vs. SUM($B$2:B2) false-positive guard (no error, no warn) |
| F10 | external_link.xlsx | link to missing workbook, `[n]` index map |
| F11 | indirect_offset.xlsx | opaque/dynamic refs, volatility |
| F12 | merged_headers.xlsx | two-row merged headers |
| F13 | mixed_types.xlsx | dates (incl. numFmt-driven dtype from raw serials), currency, %, numbers-as-text |
| F14 | sparse.xlsx | empty sheets, lone cells |
| F15 | threeD_ref.xlsx | `SUM(Jan:Mar!B2)` |
| F16 | macro_book.xlsm | VBA part preserved byte-identical through edits (I18), edits-allowed path, macro runs post-edit (live) |
| F17 | unicode_names.xlsx | non-ASCII sheet names, quotes/apostrophes in names |
| F18 | volatile.xlsx | NOW/RAND flags |
| F19 | modern_functions.xlsx | LET bindings, LAMBDA defined name called as a function, `_xlfn.XLOOKUP`, spill `A1#` consumer, `@` implicit intersection — zero spurious warnings (I20) |
| F20 | stress_map.xlsx | 40 sheets, 300 defined names (multi-area, constant, formula, lambda kinds) — map degradation rules + 8k cap (§4.3), name `kind` handling |
| F21 | chart_image.xlsx | bar chart + embedded PNG (openpyxl *can create* these in new files) — surgical-edit part-diff proves chart/media preservation (I18) |

**21 generated + 3 live-authored (L1–L3, §8.6).**

### 8.4 T4 — Robustness (optional, network-permitting)

If you can fetch a public spreadsheet corpus from GitHub (e.g., a mirror of the Enron spreadsheet set or EUSES), run the indexer over ≤ 200 files asserting only *no crash + per-file time bound*. Skip cleanly if network/corpus unavailable; never block on this.

### 8.5 T6 — MCP conformance

Drive the built server as a subprocess over stdio using the `mcp` client SDK: initialize, list tools (assert all **14**, with annotations exactly as declared in §6.1, and a non-empty `instructions` field), call every tool happy-path + one error path, assert schema validity, response caps (S6), pagination cursors round-trip, **`E_STALE_CURSOR` after a write invalidates an outstanding cursor**, and **`E_PATH_DENIED` with `EXCEL_LSP_ROOT` set**.

### 8.6 T5 — Live Excel pass (computer use, this machine) — required, evidence committed

Phase 0 recon: detect OS; confirm desktop Excel is installed and launches. Automation lanes: **Windows** → COM via pywin32 (`win32com.client.Dispatch("Excel.Application")`) for assertions, UI + screenshots for evidence; **macOS** → AppleScript via `osascript` (`tell application "Microsoft Excel"`) for assertions, UI + screenshots for evidence. If neither lane works, pause per §0.5.

Protocol (script it as a checklist in `docs/evidence/EVIDENCE.md`, save numbered screenshots to `docs/evidence/`):

1. **Author fixtures in Excel itself.** Using the Excel UI, build three live workbooks: L1 a small 2-sheet model with fill-down formulas and a named range; L2 a table (Insert → Table) with a totals row; L3 a workbook where you deliberately break one formula (delete a referenced column to force #REF!) and hand-edit one cell of a filled column to create an inconsistency. Save natively. These complement the generated fixtures with genuinely Excel-authored files (which *will* contain `_xlfn.` prefixes and real shared-formula groups — free coverage).
2. Index L1–L3; assert maps, diagnostics (the #REF! and the inconsistency are found), and traces match what you built. Screenshot Excel next to the CLI output.
3. **Round-trip:** via Excel LSP, `write_cells` + `set_column_formula` on L1 and on F16 (.xlsm). Open each in Excel: screenshot proving no repair dialog; let Excel recalc (fullCalcOnLoad), save; `refresh(recalculated=…)`; assert recalculated values via COM/AppleScript reads match expected numbers; assert staleness cleared; for F16 confirm the macro still exists and **runs** (the `Stamp` macro writes 42 to Z1 — assert it via COM/AppleScript).
4. **Cross-validate tracing:** in Excel, use Formulas → Trace Precedents/Dependents on 3 chosen cells in L1; screenshot the arrows; assert `trace_*` returns the same sets.
5. **Open-in-Excel write refusal:** with L1 still open in Excel, call `write_cells` → expect `E_OPEN_IN_EXCEL` (lockfile) or `E_LOCKED` (Windows share violation); screenshot the structured error alongside the open workbook. This proves the silent-clobber path (§5.7) is closed.
6. **Chart preservation, visually:** edit F21 via `write_cells`, open it in Excel — chart and image intact, no repair dialog; screenshot. (The part-diff already proved the bytes; this is the human-legible evidence for S4.)
7. Record a short GIF/screen capture of an agent (or the CLI) tracing lineage for the README hero demo.

### 8.7 CI

GitHub Actions: `{ubuntu, windows, macos} × {3.11, 3.12, 3.13}` → ruff, pyright, pytest (T1–T3 incl. T-oracle, T6) with coverage gate ≥ 85% on `core/`. Fixture regeneration is deterministic across the matrix because openpyxl is pinned exactly (§2). Live-Excel tests are marked `@pytest.mark.live` and excluded from CI (they ran on this machine; evidence is committed). Optional experiment, non-blocking: headless LibreOffice recalc verification (`soffice --headless --convert-to xlsx`) — if LO's recalc-on-load can't be forced reliably, drop it and note why in the agent log.

---

## 9. Benchmarks, LLM evals, and charts

You run these yourself and/or spawn a **benchmark subagent** (recommended: one subagent owns §9 end-to-end in Phase 8). All results, raw data (CSV/JSON), and chart-generation scripts are committed under `benchmarks/`.

### 9.1 Arms

- **A — Excel LSP** (our MCP server).
- **B — Naive dump baseline**: a ~60-line MCP server you write in `benchmarks/baseline_server.py` with two tools: `read_workbook_full(path)` (all sheets as CSV text) and `read_sheet(path, sheet)`. This is the honest stand-in for "just give the model the spreadsheet."
- **C — haris-musa/excel-mcp-server** (via `uvx excel-mcp-server stdio`), optional: include if it installs and runs cleanly in ≤ 15 minutes of effort; otherwise skip and note it.

### 9.2 Task suite (deterministic answers, checker script per task)

On F03 (reference model), F07, F08, F13: **B1** lineage — enumerate the input cells that feed `Summary!C10`; **B2** audit — find the tampered formula cell; **B3** error census — list all error cells with codes; **B4** schema — report column names+dtypes of the F13 region; **B5** impact — what changes if `Inputs!B2` changes (dependent set); **B6** QA lookup — a value question answerable via find + small read. Answers are exact sets/values; `benchmarks/check.py` grades transcripts.

**Answer contract (required for gradability):** every task prompt ends with — *"The last line of your reply must be exactly: `ANSWER: <json>`"* — and documents the expected JSON shape for that task in `benchmarks/tasks/<id>.md`. `check.py` parses **only** the final `ANSWER:` line. Graders that fuzzy-match prose are how benchmark numbers become fiction.

### 9.3 Two measurement modes

- **Scripted replays (no LLM, deterministic):** for each task × arm, a fixed, reasonable tool-call sequence you define (arm B necessarily dumps). Measure payload tokens of all tool results + minimal glue text.
- **LLM evals (required):** run each task × arm with headless Claude Code as the agent: `claude -p "<task prompt>" --output-format json --bare --mcp-config benchmarks/mcp_<arm>.json --allowedTools "mcp__<server>" --max-turns 15`. `--bare` isolates the run from this project's own MCP config. Parse the result JSON for token usage and cost fields (**VERIFY** exact field names from a ping run; also confirm flags via `claude --help` — CLI flags evolve). 2 repetitions per cell; **report both repetitions individually plus an agreement column** — with n=2, disagreement is signal, not noise to average away. Grade accuracy with the checkers. Cost guard: keep total spend under ~$15 or 80 headless runs, whichever first; if `claude -p` can't authenticate, pause per §0.5. If arm B blows the context window on F06-scale inputs, record that as a result (`DNF: context overflow`) — that *is* the finding.

### 9.4 Metrics & token counting

Tokens via `tiktoken` (`o200k_base`, pinned version) as a consistent proxy for scripted replays; real usage numbers from the `claude -p` JSON for LLM evals (state both methods in `benchmarks/README.md`). Also record tool-call counts, wall time, accuracy, and index-build time for F06 variants.

### 9.5 Charts (matplotlib; PNG + SVG to `docs/assets/`; raw CSV committed; alt text in README)

1. Hero: grouped bars, total tokens per task, arms A/B(/C), log scale — this goes at the top of the README.
2. Tokens: scripted vs. LLM-eval side-by-side.
3. Tool calls per task per arm.
4. Accuracy table (rendered as a markdown table, not an image).
5. Index time vs. rows (1k/10k/50k) line chart, **plus an incremental-reindex series** — S1's "< 1 s incremental" claim deserves its own line on the chart, not just a checkbox.
6. "Cost of one audit" single-number callout computed from real usage.

---

## 10. Repository & release design

### 10.1 Tree

```
excel-lsp/
  README.md  LICENSE  CONTRIBUTING.md  CHANGELOG.md  SECURITY.md  CLAUDE.md
  KNOWN_ISSUES.md(if needed)
  pyproject.toml  uv.lock  .github/workflows/{ci.yml,release.yml}
  src/excel_lsp/{core,server,cli}/
  tests/{unit,property,golden,oracle,fixtures/generate.py,fixtures/assets/vbaProject.bin,fixtures/README.md,mcp}/
  benchmarks/{tasks/,baseline_server.py,run_scripted.py,run_llm_eval.py,check.py,results/,README.md}
  docs/{architecture.md,tool-reference.md,index-internals.md,agent-log.md,evidence/}
  docs/assets/  examples/{claude_code.mcp.json,agent-transcript.md}
  PLAN.md
```

### 10.2 README anatomy, in order (the persuasion document — repo reviewer checks this)

1. One-liner: "An LSP for Excel: semantic index + MCP server so AI agents navigate workbooks by symbols, references, and diagnostics — not by reading 50,000 rows." Immediately after it, the qualifier: *(LSP-style: the ideas — symbols, references, diagnostics, incremental index — not the LSP wire protocol.)* Hero token chart immediately below.
2. 60-second demo GIF (from the live pass) of lineage tracing.
3. Quickstart: `uvx excel-lsp serve` + `claude mcp add excel-lsp -- uvx excel-lsp serve` + the `.mcp.json` snippet (VERIFY the current `claude mcp add` syntax against Claude Code docs at write time).
4. The 14 tools, one-line each.
5. Mermaid architecture diagram (loader → index(SQLite/R\*Tree) → graph → MCP).
6. Benchmarks section with charts + link to raw results + how to reproduce (`excel-lsp bench`).
7. Honest comparison table vs. haris-musa excel-mcp-server, jwadow mcp-excel, naive dump — dimensions: persistent index, formula graph, incremental reindex, diagnostics, edit support (with fidelity: untouched parts byte-identical), token discipline.
8. How it works (3 paragraphs: regions, R1C1 blocks, rtree edges) linking to docs.
9. **Security & scope** (short): local files only; exactly what the two write tools touch (target sheet XML + calc metadata — everything else byte-preserved); `EXCEL_LSP_ROOT` confinement; link to `SECURITY.md`.
10. Limitations (honest: no recalc; INDIRECT opacity; header heuristics fallible — show the confidence field; written strings are inline strings; no datetime writes in v1; edits inside multi-cell array formulas are refused; spill extents not statically tracked) and roadmap (**rename refactoring** — sheet/column/defined-name rename with workbook-wide formula rewrite, the flagship LSP-like v1.x feature the graph makes cheap; a real LSP wire-protocol server for formula editing in editors; multi-workbook workspace — index linked workbooks and connect `external:` edges across them; value-level workbook diff; watch mode; xlsb; Google Sheets adapter).
11. Evidence section linking `docs/evidence/`. Microsoft trademark footer.

### 10.3 Release

- Tag `v0.1.0`; GitHub Release with notes generated from CHANGELOG; attach the hero chart.
- PyPI publish via `release.yml` on tag using a token from the user (pause point). Fallback path documented if no token.
- Submit/prepare listings for MCP registries: the official MCP registry, Smithery, mcp.so, PulseMCP — prepare the metadata files/PRs; where submission needs an account you don't have, generate the exact submission content into `docs/registry-submissions.md` and list it as a user follow-up.

---

## 11. Orchestration: you, subagents, and review loops

### 11.1 Roles

- **Orchestrator (you):** owns PLAN.md, agent-log, phase sequencing, commits, review budget accounting, and all pause-point communication with the user.
- **Implementation subagents:** one per phase (or per module within a phase if you judge it cleaner). Each gets a brief from the template in §14.2 — goal, interface contract, done-criteria, and an explicit out-of-scope list. Scope creep is the primary failure mode; the out-of-scope list is mandatory.
- **Benchmark subagent:** owns §9 in Phase 8 (may itself shell out to headless `claude -p` runs).
- **Reviewer subagents:** stateless; a fresh reviewer is spawned per review invocation with the rubric, the diff/artifacts, and no memory of prior reviews (prevents rubber-stamping).

### 11.2 Phases and gates

| Phase | Deliverables | Gates |
|---|---|---|
| P0 | Recon (OS, Excel, `uv`, git, `claude -p` auth, rtree VERIFY, PyPI name), scaffold, CI stub, fixture generator skeleton, PLAN.md, agent-log, **CLAUDE.md**, **pinned openpyxl/tiktoken versions recorded**, **F16 VBA blob authored + committed with provenance (Excel permitting; else pause)** | self-check only |
| P1 | M1 lxml parser + loader/hasher + M8 lifecycle + SQLite store + **T-oracle harness green on F01** (informational openpyxl read-only VERIFY logged) | R-mech, R-test |
| P2 | M2 regions/headers + workbook map (S2 token budget met; **F20 degradation test green**) | R-mech, R-test, **R-repo #1 (README skeleton + claims-to-artifacts plan)** |
| P3 | M3 ref extraction + M4 R1C1/fblocks (**modern matrix green on F19**) | R-mech, R-test |
| P4 | M5 graph, rtree edges, traces + trace_path, circular detection (bounded stage 2) | R-mech, R-test |
| P5 | M6 diagnostics complete | R-mech, R-test |
| P6 | M7 **surgical editor** + staleness; **part-diff tests green on F16/F21 (I18)**; live-Excel *smoke* (round-trip one file) | R-mech, R-test |
| P7 | MCP server (**14 tools**, annotations, instructions, progress) + CLI + T6 conformance | R-mech, R-test |
| P8 | Full live pass §8.6 (incl. steps 5–6) with evidence; benchmarks + LLM evals + charts §9 | R-test |
| P9 | README, docs, examples, CI green on matrix, release + registry prep | R-repo (iterate) |

A gate passes when the relevant reviewer returns APPROVE (see below). You fix findings between invocations yourself or via a fix-subagent.

### 11.3 Review protocol

Reviewer input: rubric + `git diff` for the phase + relevant artifacts (test output, fixtures, charts, README render). Reviewer output (strict format): `VERDICT: APPROVE | REVISE`, then findings, each `[critical|major|minor] location — issue — suggested fix`. APPROVE is only legal when there are zero critical and zero major findings. Minor findings may accompany APPROVE; fix them without a re-review. On REVISE: you fix, then re-invoke (that's a new invocation against the budget).

### 11.4 Budget (frozen)

**30 total review invocations: 10 × R-mech, 10 × R-test, 10 × R-repo.** Within each domain, iterate until a review returns **no findings at all** (a clean APPROVE) or the domain budget is exhausted — with one pragmatic release valve: an APPROVE with only minor findings unblocks the *phase gate*, but you keep spending that domain's remaining budget in later phases aiming for the clean pass. Track spend in PLAN.md as a table. Suggested pacing: ≤ 2 per gate early, reserve ≥ 3 R-mech and ≥ 3 R-test for P6–P8; **spend exactly 1 R-repo at the P2 gate** on the README skeleton plus a claims-to-artifacts plan (every §10.2 claim mapped to the phase that produces its evidence — cheap to check now, expensive to discover missing at P9), and reserve ≥ 3 R-repo for P9. **Exhaustion policy:** if any domain exhausts with unresolved critical findings, do not tag v0.1.0 — publish `v0.1.0-rc1` as a GitHub pre-release, write `KNOWN_ISSUES.md`, and stop, reporting to the user.

### 11.5 Rubrics

**R-mech (mechanics):** every invariant **I1–I20** in §5 explicitly checked against code/tests (I17 superseded by I18); correctness of `classify_ref` matrix incl. quoted/3-D/structured/external **and the modern matrix (`_xlfn.`, LET/LAMBDA suppression, spill, `@`)**; R1C1 relative/absolute handling; block extrusion math **incl. sheet-bounds clamping**; re-specified `W_INCONSISTENT_FORMULA` thresholds catch F07's tamper; two-stage circular detection with **bounded stage 2** (running-total guard present); **surgical writer: part preservation (I18), calcChain triple-edit, shared-group expansion, `E_ARRAY_FORMULA`, inlineStr mechanics, lockfile + conflict precondition order**; staleness propagation and E_CONFLICT; **cursor generation invalidation and the per-call freshness check**; **regex guard**; **error-table conformance (§6.1) with no traceback leaks through MCP**; response caps and no-bulk-data enforcement in server code; performance approach plausible for S1 (streaming, no dense grids, single pass); VERIFY items from §2/§9.3 actually verified and logged.

**R-test:** every module has unit coverage; hypothesis properties present and meaningful (not vacuous) **incl. the edit-script part-preservation property**; all **21+3** fixtures exist, are generated deterministically (pinned openpyxl), and are each asserted against at least once; **T-oracle green with a documented skip-list**; golden snapshots + token assertions; T6 covers all **14** tools incl. error paths, **annotations, instructions, and cursor invalidation**; live-Excel evidence complete per §8.6 checklist **including steps 5 (open-in-Excel refusal) and 6 (chart intact)** — screenshots present, numbered, indexed; benchmark harness reproducible from a clean checkout (`excel-lsp bench` or documented commands) **with the ANSWER-line contract enforced by `check.py`**; LLM-eval methodology sound (isolation via `--bare`, both repetitions reported, cost guard, DNFs recorded); coverage gate met.

**R-repo:** fresh-environment install dry-run actually performed (`uv venv` + install + `excel-lsp map` on a fixture, and the `uvx` path); README follows §10.2 order and every claim in it is backed by a committed artifact (charts, evidence, benchmark data) — **cross-check against the P2 claims-to-artifacts plan**; quickstart copy-paste works; comparison table fair and sourced; LICENSE/CONTRIBUTING/CHANGELOG/**SECURITY.md** present; **`uv.lock` committed; `CLAUDE.md` current; vbaProject.bin provenance note present**; **docs/tool-reference.md covers all 14 tools**; CI green on the full matrix; no committed junk (index DBs, `.xlsp.db*` sidecars, .DS_Store, dead code); registry submission content prepared; trademark footer present.

### 11.6 Subagent failure protocol

A subagent that misses its done-criteria **twice**, or violates its out-of-scope list **once**, is stopped: `git checkout`/reset the working tree to the last green commit, then re-brief with narrower scope and the failure notes appended to the brief. After a **third** failure on the same module, the orchestrator implements it directly — no fourth subagent. Every reset is an agent-log entry (what failed, what was reverted, what changed in the re-brief). Scope creep is the primary failure mode; unbounded retry churn is the second, and this protocol is its circuit breaker.

---

## 12. Risk register

| Risk | Mitigation |
|---|---|
| openpyxl load/save silently drops charts/images/drawings | **Eliminated by design:** surgical writer never round-trips through openpyxl (§5.7); I18 part-diff on F16/F21; live step 6 |
| Surgical-writer XML bugs (element ordering, cell types, calcChain triple-edit) | Part-diff unit tests; hypothesis edit-script property; re-parse equality check; live pass is the canary (S4) |
| lxml parser divergence from Excel/openpyxl semantics | T-oracle cross-validation on all fixtures with documented skip-list (§8.1a) |
| Generated fixtures lack cached formula values | generate.py lxml post-pass injects `<v>` / `t="e"` (§8.3a) — without it, E_ERRVAL/profile/B3 test nothing |
| inlineStr strings poorly handled by some third-party tools | Excel/LibreOffice read them natively; documented limitation (§10.2) |
| LET/LAMBDA workbooks drown in false `W_UNKNOWN_NAME` | Parameter suppression (§5.3) + I20 + F19 |
| R\*Tree missing from sqlite build | Detected P0; EdgeStore fallback specified §2 |
| INDIRECT/OFFSET opacity confuses users | Opaque edges + I_DYNAMIC_REF + README limitation |
| Running-total false circulars / O(n²) circular verification | Two-stage SCC with O(cells) self-inclusion + bounded stage 2b → `W_POSSIBLE_CIRCULAR`; F09a/b guard both directions |
| Header heuristics wrong on weird sheets | Confidence field, ListObject precedence, documented limitation |
| `claude -p` flags/fields drift | VERIFY via `--help` + ping run before harness code hardens |
| Excel automation flakiness (dialogs, activation) | COM/AppleScript for assertions, UI only for evidence; pause point for blockers |
| Windows locking / macOS open-in-Excel silent clobber | E_LOCKED retry; `~$` lockfile → E_OPEN_IN_EXCEL; live step 5 proves the refusal |
| Map blows the cap on many-sheet/many-name workbooks | Degradation rules (§4.3) + F20 cap test |
| PyPI name taken / no token | Fallback name; git-tag + `uvx --from git+` path |

---

## 13. Definition of Done — tick literally, in PLAN.md

- [ ] S1–S7 all verified with evidence linked (S4 includes the part-diff proof)
- [ ] CI green: 3 OS × 3 Python, lint + types + tests, coverage ≥ 85% core
- [ ] All **21** generated fixtures + 3 live-authored fixtures exercised; **T-oracle green with documented skip-list**
- [ ] Live-Excel evidence committed: authoring, round-trip (incl. .xlsm/VBA with macro-runs assertion), trace cross-validation, no-repair screenshots, **open-in-Excel refusal (step 5)**, **chart-intact screenshot (step 6)**, demo GIF
- [ ] Benchmarks: scripted + LLM-eval results committed with raw data, **ANSWER-contract checkers**, **both repetitions reported**, and 5 charts (incl. incremental series); hero chart in README
- [ ] README complete per §10.2 (**14 tools, LSP-style qualifier, Security & scope note**); fresh-env install verified
- [ ] **SECURITY.md, CLAUDE.md, uv.lock committed; `tests/fixtures/README.md` documents vbaProject.bin provenance**
- [ ] Review ledger: budget accounting complete (incl. the P2 R-repo spend); every domain ended on clean APPROVE (or exhaustion policy executed)
- [ ] `v0.1.0` tagged; PyPI published or fallback documented; registry submissions prepared
- [ ] agent-log.md tells the whole story; final summary written for the user listing follow-ups (tokens, registry accounts)

---

## 14. Appendix — templates

### 14.1 agent-log entry

```
## 2026-07-15 P3 — R1C1 block merge strategy
Decision: column-major runs then horizontal rectangle merge.
Alternatives: 2-D union-find over per-cell R1C1 (simpler, slower on F06).
Rationale: O(cells) single pass; F06 target. Revisit if I11 fails on ragged blocks.
```

### 14.2 Implementation subagent brief

```
ROLE: Implement <module> for Excel LSP, Phase <n>.
CONTEXT: Read HANDOFF §<...>, existing code in src/excel_lsp/<...>. CLAUDE.md applies.
GOAL: <one sentence>.
INTERFACE CONTRACT: <exact functions/classes + signatures to expose>.
DONE-CRITERIA: <tests that must pass, invariants I<..> satisfied, perf bound>.
OUT OF SCOPE: <explicit list — do not touch server/, do not add deps, do not refactor unrelated code>.
DELIVER: code + tests + a 5-line summary of decisions for the agent log.
```

### 14.3 Reviewer invocation

```
ROLE: You are a fresh, adversarial reviewer. You have no stake in this code.
RUBRIC: <paste R-mech | R-test | R-repo rubric from §11.5>.
MATERIALS: <diff, artifacts, paths>.
OUTPUT FORMAT (strict):
VERDICT: APPROVE | REVISE
FINDINGS:
- [critical|major|minor] <file:loc or artifact> — <issue> — <suggested fix>
Rules: APPROVE only with zero critical+major findings. Verify claims by reading
code/tests/artifacts, not summaries. Hunt for: silent scope creep, vacuous tests,
uncommitted evidence, cap violations, invariants asserted but untested.
```

### 14.4 CLAUDE.md (repo root — create at P0, keep current)

```
# Excel LSP — working conventions
One-liner: LSP-style semantic index + MCP server for Excel workbooks.

Layout: src/excel_lsp/{core,server,cli}; tests/{unit,property,golden,oracle,fixtures,mcp};
benchmarks/; docs/. Plans: PLAN.md. Decisions: docs/agent-log.md. Review ledger: PLAN.md table.

Frozen decisions: see HANDOFF §0.4. Do not renegotiate them in code.

Before touching a module, read its HANDOFF section:
  parser/loader → §5.1 | regions → §5.2 | classify_ref → §5.3 | R1C1/blocks → §5.4
  graph → §5.5 | diagnostics → §5.6 | editor → §5.7 | lifecycle → §5.8
  tools/server → §6 | tests → §8 | benchmarks → §9

Hard rules:
- NEVER load-and-save a workbook through openpyxl in src/ (HANDOFF §5.7). Edits are surgical.
- No bulk data in tool responses; caps per HANDOFF §6.1.
- Value normalization goes through the one shared function (§5.1).

Commits: conventional (feat:/fix:/test:/docs:/chore:), per milestone, on main, no force-push.
Tests: uv run pytest        Lint/type: uv run ruff check . && uv run pyright
Fixtures: uv run python tests/fixtures/generate.py
```

---

*End of handoff (v1.1). Begin with Phase 0.*
