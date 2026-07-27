"""Deterministic formula-block construction and consistency checks."""

from __future__ import annotations

from bisect import bisect_right
from collections import Counter, defaultdict
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from fractions import Fraction
from itertools import groupby

from openpyxl.formula.tokenizer import TokenizerError

from excel_lsp.core.formulas.a1 import (
    CellRef,
)
from excel_lsp.core.formulas.r1c1 import to_r1c1
from excel_lsp.core.formulas.translation import translate_a1_formula
from excel_lsp.core.models import Rect

_MAX_ROW = 1_048_576
_MAX_COLUMN = 16_384

_PatternKey = tuple[bool, str]


@dataclass(frozen=True, slots=True)
class FormulaCell:
    """One worksheet formula and its one-based coordinate."""

    row: int
    col: int
    formula: str

    def __post_init__(self) -> None:
        _validate_coordinate(self.row, self.col)
        if not self.formula:
            raise ValueError("formula must be a non-empty string")


@dataclass(frozen=True, slots=True)
class FormulaPattern:
    """One formula cell paired with its normalized pattern."""

    row: int
    col: int
    formula: str
    r1c1: str
    parsed: bool = True

    def __post_init__(self) -> None:
        _validate_coordinate(self.row, self.col)
        if not self.formula or not self.r1c1:
            raise ValueError("formula patterns must not be empty")
        if type(self.parsed) is not bool:
            raise ValueError("parsed must be a boolean")

    @property
    def key(self) -> _PatternKey:
        """Return a collision-free key for parsed and opaque patterns."""
        return self.parsed, self.r1c1


@dataclass(frozen=True, slots=True)
class FormulaBlock:
    """One exact formula rectangle with a deterministic sheet-local ordinal."""

    n: int
    rect: Rect
    r1c1: str
    anchor_formula: str
    parsed: bool = True
    volatile: bool = False
    opaque: bool = False

    def __post_init__(self) -> None:
        if type(self.n) is not int or self.n < 0:
            raise ValueError("formula-block ordinal must be a nonnegative integer")
        if not self.r1c1 or not self.anchor_formula:
            raise ValueError("formula-block text must not be empty")
        for name in ("parsed", "volatile", "opaque"):
            if type(getattr(self, name)) is not bool:
                raise ValueError(f"{name} must be a boolean")


@dataclass(frozen=True, slots=True)
class InconsistentFormula:
    """One minority-pattern cell and its deterministic expected block."""

    row: int
    col: int
    expected_r1c1: str
    dominant_block_n: int

    def __post_init__(self) -> None:
        _validate_coordinate(self.row, self.col)
        if not self.expected_r1c1:
            raise ValueError("expected R1C1 pattern must not be empty")
        if type(self.dominant_block_n) is not int or self.dominant_block_n < 0:
            raise ValueError("dominant formula-block ordinal must be nonnegative")


@dataclass(slots=True)
class _ColumnRun:
    col: int
    row_min: int
    row_max: int
    pattern: FormulaPattern


@dataclass(slots=True)
class _MutableBlock:
    row_min: int
    row_max: int
    col_min: int
    col_max: int
    pattern: FormulaPattern


@dataclass(frozen=True, slots=True)
class _Candidate:
    row: int
    col: int
    expected_r1c1: str
    dominant_block_n: int
    dominant_count: int
    run_length: int
    axis: str

    @property
    def ratio(self) -> Fraction:
        return Fraction(self.dominant_count, self.run_length)


def normalize_formula_cells(cells: Iterable[FormulaCell]) -> tuple[FormulaPattern, ...]:
    """Normalize formulas without allowing malformed input to drop a cell."""
    patterns: list[FormulaPattern] = []
    seen: set[tuple[int, int]] = set()
    for cell in cells:
        coordinate = (cell.row, cell.col)
        if coordinate in seen:
            raise ValueError("formula cell coordinates must be unique")
        seen.add(coordinate)
        try:
            normalized = to_r1c1(cell.formula, CellRef(cell.row, cell.col))
        except (IndexError, TokenizerError, ValueError):
            pattern = FormulaPattern(
                cell.row,
                cell.col,
                cell.formula,
                cell.formula,
                parsed=False,
            )
        else:
            pattern = FormulaPattern(cell.row, cell.col, cell.formula, normalized)
        patterns.append(pattern)
    return tuple(patterns)


def build_formula_blocks(patterns: Iterable[FormulaPattern]) -> tuple[FormulaBlock, ...]:
    """Grow vertical runs, then merge equal adjacent runs horizontally.

    The input is a column-major stream ordered by ``(col, row)``.  Enforcing
    that contract keeps block construction linear and avoids a second copy or
    sort of large formula sheets.
    """
    runs_by_column: dict[int, list[_ColumnRun]] = defaultdict(list)
    pattern_count = 0
    prior_coordinate: tuple[int, int] | None = None
    for column, column_group in groupby(patterns, key=lambda item: item.col):
        current: _ColumnRun | None = None
        for pattern in column_group:
            coordinate = (pattern.col, pattern.row)
            if prior_coordinate is not None and coordinate <= prior_coordinate:
                raise ValueError("formula patterns must be unique and column-major ordered")
            prior_coordinate = coordinate
            pattern_count += 1
            if (
                current is not None
                and pattern.row == current.row_max + 1
                and pattern.key == current.pattern.key
                and _translates_exactly(current.pattern, pattern)
            ):
                current.row_max = pattern.row
                continue
            current = _ColumnRun(column, pattern.row, pattern.row, pattern)
            runs_by_column[column].append(current)

    mutable: list[_MutableBlock] = []
    previous: dict[tuple[int, int, _PatternKey], _MutableBlock] = {}
    previous_column: int | None = None
    for column in sorted(runs_by_column):
        if previous_column is None or column != previous_column + 1:
            previous = {}
        current_blocks: dict[tuple[int, int, _PatternKey], _MutableBlock] = {}
        for run in runs_by_column[column]:
            key = (run.row_min, run.row_max, run.pattern.key)
            block = previous.get(key)
            if block is not None and not _translates_exactly(block.pattern, run.pattern):
                block = None
            if block is None:
                block = _MutableBlock(
                    run.row_min,
                    run.row_max,
                    column,
                    column,
                    run.pattern,
                )
                mutable.append(block)
            else:
                block.col_max = column
            current_blocks[key] = block
        previous = current_blocks
        previous_column = column

    mutable.sort(
        key=lambda block: (
            block.row_min,
            block.col_min,
            block.row_max,
            block.col_max,
            block.pattern.r1c1,
            block.pattern.parsed,
        )
    )
    result = tuple(
        FormulaBlock(
            n=index,
            rect=Rect(block.row_min, block.row_max, block.col_min, block.col_max),
            r1c1=block.pattern.r1c1,
            anchor_formula=block.pattern.formula,
            parsed=block.pattern.parsed,
            opaque=not block.pattern.parsed,
        )
        for index, block in enumerate(mutable)
    )
    covered_area = sum(_rect_area(block.rect) for block in result)
    if covered_area != pattern_count:
        raise RuntimeError("formula-block rectangles do not exactly cover formula cells")
    return result


def _translates_exactly(source: FormulaPattern, target: FormulaPattern) -> bool:
    """Verify the directional fill relationship inside one coarse R1C1 bucket."""
    if not source.parsed or not target.parsed:
        return source.key == target.key
    if (
        "#" not in source.formula
        and "#" not in target.formula
        and not _has_ascii_lowercase(source.formula)
        and not _has_ascii_lowercase(target.formula)
    ):
        # For canonical uppercase A1 spelling, the R1C1 bucket already retains
        # every lexical distinction that translation preserves. Lowercase A1
        # columns are directional (``a1`` translates to ``A2``), so any formula
        # containing lowercase text takes the exact translator path below.
        return True
    try:
        translated = translate_a1_formula(
            source.formula,
            origin=CellRef(source.row, source.col),
            target=CellRef(target.row, target.col),
            preserve_coordinate_spills=True,
        )
    except (IndexError, TokenizerError, ValueError):
        return False
    return translated == target.formula


def _has_ascii_lowercase(value: str) -> bool:
    return any("a" <= character <= "z" for character in value)


def detect_inconsistent_formulas(
    patterns: Iterable[FormulaPattern],
    regions: Sequence[Rect],
    blocks: Sequence[FormulaBlock],
) -> tuple[InconsistentFormula, ...]:
    """Detect low-frequency patterns in long row and column formula runs."""
    ordered = sorted(patterns, key=lambda item: (item.row, item.col))
    if not ordered or not regions:
        return ()
    block_by_cell = _block_owners(blocks)
    located = _patterns_by_region(ordered, regions)
    candidates: dict[tuple[int, int], _Candidate] = {}

    by_column: dict[tuple[int, int], list[FormulaPattern]] = defaultdict(list)
    by_row: dict[tuple[int, int], list[FormulaPattern]] = defaultdict(list)
    for region_index, pattern in located:
        by_column[(region_index, pattern.col)].append(pattern)
        by_row[(region_index, pattern.row)].append(pattern)

    for group in by_column.values():
        group.sort(key=lambda item: item.row)
        for run in _contiguous_runs(group, coordinate="row"):
            _collect_run_candidates(run, "vertical", block_by_cell, candidates)
    for group in by_row.values():
        group.sort(key=lambda item: item.col)
        for run in _contiguous_runs(group, coordinate="col"):
            _collect_run_candidates(run, "horizontal", block_by_cell, candidates)

    return tuple(
        InconsistentFormula(
            row=candidate.row,
            col=candidate.col,
            expected_r1c1=candidate.expected_r1c1,
            dominant_block_n=candidate.dominant_block_n,
        )
        for _coordinate, candidate in sorted(candidates.items())
    )


def _contiguous_runs(
    patterns: Sequence[FormulaPattern],
    *,
    coordinate: str,
) -> Iterable[tuple[FormulaPattern, ...]]:
    current: list[FormulaPattern] = []
    prior: int | None = None
    for pattern in patterns:
        value = getattr(pattern, coordinate)
        if prior is not None and value != prior + 1:
            yield tuple(current)
            current = []
        current.append(pattern)
        prior = value
    if current:
        yield tuple(current)


def _collect_run_candidates(
    run: Sequence[FormulaPattern],
    axis: str,
    block_by_cell: dict[tuple[int, int], FormulaBlock],
    candidates: dict[tuple[int, int], _Candidate],
) -> None:
    run_length = len(run)
    if run_length < 5:
        return
    counts = Counter(pattern.key for pattern in run)
    dominant_count = max(counts.values())
    dominant_keys = sorted(key for key, count in counts.items() if count == dominant_count)
    dominant_key = dominant_keys[0]
    minority_count = run_length - dominant_count
    if dominant_count * 5 < run_length * 4:
        return
    if minority_count > max(3, (run_length + 19) // 20):
        return

    dominant_blocks = Counter(
        block_by_cell[(pattern.row, pattern.col)].n
        for pattern in run
        if pattern.key == dominant_key
    )
    if not dominant_blocks:
        raise RuntimeError("dominant formula pattern has no formula block")
    best_block_count = max(dominant_blocks.values())
    dominant_block_n = min(
        block_n for block_n, count in dominant_blocks.items() if count == best_block_count
    )
    expected_r1c1 = next(pattern.r1c1 for pattern in run if pattern.key == dominant_key)
    for pattern in run:
        if pattern.key == dominant_key:
            continue
        candidate = _Candidate(
            row=pattern.row,
            col=pattern.col,
            expected_r1c1=expected_r1c1,
            dominant_block_n=dominant_block_n,
            dominant_count=dominant_count,
            run_length=run_length,
            axis=axis,
        )
        coordinate = (pattern.row, pattern.col)
        current = candidates.get(coordinate)
        if current is None or _candidate_is_better(candidate, current):
            candidates[coordinate] = candidate


def _candidate_is_better(candidate: _Candidate, current: _Candidate) -> bool:
    if candidate.ratio != current.ratio:
        return candidate.ratio > current.ratio
    if candidate.run_length != current.run_length:
        return candidate.run_length > current.run_length
    if candidate.axis != current.axis:
        return candidate.axis == "vertical"
    return (candidate.expected_r1c1, candidate.dominant_block_n) < (
        current.expected_r1c1,
        current.dominant_block_n,
    )


def _patterns_by_region(
    patterns: Sequence[FormulaPattern],
    regions: Sequence[Rect],
) -> tuple[tuple[int, FormulaPattern], ...]:
    ordered_regions = sorted(enumerate(regions), key=lambda item: (item[1].row_min, item[0]))
    next_region = 0
    active: list[tuple[int, Rect]] = []
    active_starts: list[int] = []
    current_row = 0
    located: list[tuple[int, FormulaPattern]] = []
    for pattern in patterns:
        if pattern.row != current_row:
            current_row = pattern.row
            active = [item for item in active if item[1].row_max >= current_row]
            while next_region < len(ordered_regions):
                item = ordered_regions[next_region]
                if item[1].row_min > current_row:
                    break
                next_region += 1
                if item[1].row_max >= current_row:
                    active.append(item)
            active.sort(key=lambda item: (item[1].col_min, item[1].col_max, item[0]))
            active_starts = [item[1].col_min for item in active]
        candidate_index = bisect_right(active_starts, pattern.col) - 1
        if candidate_index < 0:
            continue
        region_index, region = active[candidate_index]
        if region.col_min <= pattern.col <= region.col_max:
            located.append((region_index, pattern))
    return tuple(located)


def _block_owners(blocks: Sequence[FormulaBlock]) -> dict[tuple[int, int], FormulaBlock]:
    owners: dict[tuple[int, int], FormulaBlock] = {}
    for block in blocks:
        for row in range(block.rect.row_min, block.rect.row_max + 1):
            for col in range(block.rect.col_min, block.rect.col_max + 1):
                coordinate = (row, col)
                if coordinate in owners:
                    raise ValueError("formula blocks must not overlap")
                owners[coordinate] = block
    return owners


def _validate_coordinate(row: int, col: int) -> None:
    if type(row) is not int or type(col) is not int:
        raise ValueError("formula coordinates must be integers")
    if not 1 <= row <= _MAX_ROW or not 1 <= col <= _MAX_COLUMN:
        raise ValueError("formula coordinate exceeds Excel bounds")


def _rect_area(rect: Rect) -> int:
    return (rect.row_max - rect.row_min + 1) * (rect.col_max - rect.col_min + 1)


__all__ = [
    "FormulaBlock",
    "FormulaCell",
    "FormulaPattern",
    "InconsistentFormula",
    "build_formula_blocks",
    "detect_inconsistent_formulas",
    "normalize_formula_cells",
]
