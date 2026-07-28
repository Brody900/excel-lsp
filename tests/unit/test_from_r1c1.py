from __future__ import annotations

import pytest

from excel_lsp.core.formulas.a1 import CellRef
from excel_lsp.core.formulas.from_r1c1 import from_r1c1


@pytest.mark.parametrize(
    ("formula", "anchor", "expected"),
    (
        ("=RC[-1]*2", CellRef(5, 3), "=B5*2"),
        ("=R[-1]C+R1C1", CellRef(5, 3), "=C4+$A$1"),
        ("=SUM(R[-2]C:R[-1]C)", CellRef(5, 3), "=SUM(C3:C4)"),
        ("=SUM(R1:R3)", CellRef(5, 3), "=SUM($1:$3)"),
        ("=SUM(C[-2]:C)", CellRef(5, 3), "=SUM(A:C)"),
        ("='Input Sheet'!R2C2+[1]Data!RC", CellRef(5, 3), "='Input Sheet'!$B$2+[1]Data!C5"),
        ('="R[-1]C"&RC', CellRef(5, 3), '="R[-1]C"&C5'),
    ),
)
def test_from_r1c1_renders_supported_reference_shapes(
    formula: str,
    anchor: CellRef,
    expected: str,
) -> None:
    assert from_r1c1(formula, anchor) == expected


def test_from_r1c1_returns_none_for_a1_or_reference_free_formula() -> None:
    assert from_r1c1("=A1*2", CellRef(5, 3)) is None
    assert from_r1c1("=1+2", CellRef(5, 3)) is None


def test_from_r1c1_rejects_mixed_modes_and_out_of_bounds_translation() -> None:
    with pytest.raises(ValueError, match="mixes A1 and R1C1"):
        from_r1c1("=A1+R[-1]C", CellRef(5, 3))
    with pytest.raises(ValueError, match="exceeds worksheet bounds"):
        from_r1c1("=R[-1]C", CellRef(1, 1))
