"""Property checks for spatial dependency-edge lookup parity."""

from __future__ import annotations

import sqlite3
from collections.abc import Generator
from contextlib import contextmanager

from hypothesis import given, settings
from hypothesis import strategies as st

from excel_lsp.core.index.edges import EdgeStore
from excel_lsp.core.models import Rect

_ROW_COORDINATES = st.one_of(
    st.sampled_from((1, 2, 3, 1_048_574, 1_048_575, 1_048_576)),
    st.integers(min_value=1, max_value=256),
)
_COLUMN_COORDINATES = st.one_of(
    st.sampled_from((1, 2, 3, 16_382, 16_383, 16_384)),
    st.integers(min_value=1, max_value=64),
)


@st.composite
def _bounded_rectangles(draw: st.DrawFn) -> Rect:
    row_a = draw(_ROW_COORDINATES)
    row_b = draw(_ROW_COORDINATES)
    col_a = draw(_COLUMN_COORDINATES)
    col_b = draw(_COLUMN_COORDINATES)
    return Rect(
        min(row_a, row_b),
        max(row_a, row_b),
        min(col_a, col_b),
        max(col_a, col_b),
    )


_RECTANGLES = st.one_of(
    st.just(Rect(1, 1_048_576, 1, 1)),
    st.just(Rect(1, 1_048_576, 16_384, 16_384)),
    st.just(Rect(1, 1_048_576, 1, 16_384)),
    _bounded_rectangles(),
)
_EDGE_SETS = st.lists(
    st.tuples(st.integers(min_value=1, max_value=3), _RECTANGLES),
    min_size=0,
    max_size=60,
)


@contextmanager
def _edge_store(*, prefer_rtree: bool) -> Generator[EdgeStore, None, None]:
    connection = sqlite3.connect(":memory:", isolation_level=None)
    try:
        backend = EdgeStore.ensure_schema(connection, prefer_rtree=prefer_rtree)
        assert backend == ("rtree" if prefer_rtree else "interval")
        yield EdgeStore(connection)
    finally:
        connection.close()


@given(
    edges=_EDGE_SETS,
    query_sheet=st.integers(min_value=1, max_value=3),
    query_rect=_RECTANGLES,
    point_row=_ROW_COORDINATES,
    point_col=_COLUMN_COORDINATES,
    page_size=st.sampled_from((1, 2, 7, 16, 256)),
)
@settings(max_examples=80, derandomize=True, database=None, deadline=None)
def test_rtree_and_interval_results_equal_independent_brute_scan(
    edges: list[tuple[int, Rect]],
    query_sheet: int,
    query_rect: Rect,
    point_row: int,
    point_col: int,
    page_size: int,
) -> None:
    """Both backends match brute overlap at boundaries and across keyset pages."""
    expected_range = tuple(
        edge_id
        for edge_id, (sheet_id, rect) in enumerate(edges, start=1)
        if sheet_id == query_sheet and rect.intersects(query_rect)
    )
    point = Rect(point_row, point_row, point_col, point_col)
    expected_point = tuple(
        edge_id
        for edge_id, (sheet_id, rect) in enumerate(edges, start=1)
        if sheet_id == query_sheet and rect.intersects(point)
    )

    for prefer_rtree in (True, False):
        with _edge_store(prefer_rtree=prefer_rtree) as store:
            for edge_id, (sheet_id, rect) in enumerate(edges, start=1):
                store.insert(edge_id, sheet_id, rect)

            assert store.query_range(query_sheet, query_rect) == expected_range
            assert (
                tuple(store.iter_query_range(query_sheet, query_rect, page_size=page_size))
                == expected_range
            )
            assert store.query_point(query_sheet, point_row, point_col) == expected_point
            assert (
                tuple(
                    store.iter_query_point(
                        query_sheet,
                        point_row,
                        point_col,
                        page_size=page_size,
                    )
                )
                == expected_point
            )
