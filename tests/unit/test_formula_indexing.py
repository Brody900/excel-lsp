"""High-signal integration tests for pure per-sheet formula orchestration."""

from __future__ import annotations

from excel_lsp.core.formulas.blocks import FormulaCell
from excel_lsp.core.formulas.indexing import FormulaEdge, analyze_sheet_formulas
from excel_lsp.core.formulas.references import (
    ReferenceContext,
    StructuredContextRequirement,
    TableBinding,
    structured_context_requirement,
)
from excel_lsp.core.models import DefinedName, NameArea, Rect, SheetDescriptor, TableInfo
from excel_lsp.core.parse.coordinates import column_label


def _sheet(name: str, order: int) -> SheetDescriptor:
    return SheetDescriptor(
        order=order,
        name=name,
        sheet_id=order + 1,
        rel_id=f"rId{order + 1}",
        xml_part=f"xl/worksheets/sheet{order + 1}.xml",
        kind="worksheet",
    )


def test_f07_shape_builds_exact_blocks_edges_and_one_inconsistency() -> None:
    sheet = _sheet("FormulaBlocks", 0)
    context = ReferenceContext((sheet,))
    cells = tuple(
        FormulaCell(
            row,
            3,
            f"=A{row}+B{row}" if row == 12 else f"=A{row}*B{row}",
        )
        for row in range(2, 22)
    )

    result = analyze_sheet_formulas(
        sheet,
        cells,
        (Rect(1, 21, 1, 3),),
        context,
    )

    assert [(block.n, block.rect, block.r1c1) for block in result.blocks] == [
        (0, Rect(2, 11, 3, 3), "=RC[-2]*RC[-1]"),
        (1, Rect(12, 12, 3, 3), "=RC[-2]+RC[-1]"),
        (2, Rect(13, 21, 3, 3), "=RC[-2]*RC[-1]"),
    ]
    assert result.edges == (
        FormulaEdge(0, 0, Rect(2, 11, 1, 1), "ref"),
        FormulaEdge(0, 0, Rect(2, 11, 2, 2), "ref"),
        FormulaEdge(1, 0, Rect(12, 12, 1, 1), "ref"),
        FormulaEdge(1, 0, Rect(12, 12, 2, 2), "ref"),
        FormulaEdge(2, 0, Rect(13, 21, 1, 1), "ref"),
        FormulaEdge(2, 0, Rect(13, 21, 2, 2), "ref"),
    )
    assert len(result.diagnostics) == 1
    diagnostic = result.diagnostics[0]
    assert (
        diagnostic.severity,
        diagnostic.code,
        diagnostic.row,
        diagnostic.col,
        diagnostic.ref,
    ) == ("warn", "W_INCONSISTENT_FORMULA", 12, 3, "cell:FormulaBlocks!C12")
    assert dict(diagnostic.related) == {
        "dominantBlock": "fblock:FormulaBlocks:0",
        "expectedR1C1": "=RC[-2]*RC[-1]",
    }


def test_spill_edge_resolves_only_the_anchor_instead_of_extruding_the_block() -> None:
    sheet = _sheet("Spill", 0)
    context = ReferenceContext((sheet,))
    # Deliberately identical raw spill formulas form one block. Every cell
    # reads the same dynamic array anchored at A1, so the edge must stay A1.
    cells = tuple(FormulaCell(row, 2, "=A1#") for row in range(2, 6))

    result = analyze_sheet_formulas(
        sheet,
        cells,
        (Rect(1, 6, 1, 2),),
        context,
    )

    assert len(result.blocks) == 1
    assert result.blocks[0].rect == Rect(2, 5, 2, 2)
    assert result.blocks[0].r1c1 == "=A1#"
    assert result.edges == (FormulaEdge(0, 0, Rect(1, 1, 1, 1), "spill"),)
    assert result.diagnostics == ()


def test_mixed_relative_and_structured_endpoints_extrude_the_complete_block() -> None:
    sheet = _sheet("Data", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Data",
                TableInfo(
                    "SalesTable",
                    "SalesTable",
                    "A1:C6",
                    1,
                    1,
                    ("Item", "Qty", "Price"),
                ),
            ),
        ),
    )
    cells = tuple(FormulaCell(row, 5, f"=SUM(A{row}:SalesTable[Qty])") for row in range(3, 11))

    single = analyze_sheet_formulas(sheet, cells[:1], (Rect(1, 10, 1, 5),), context)
    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 10, 1, 5),), context)

    assert single.edges == (FormulaEdge(0, 0, Rect(2, 5, 1, 2), "structured:SalesTable[Qty]"),)
    assert [(block.rect, block.r1c1) for block in result.blocks] == [
        (Rect(3, 10, 5, 5), "=SUM(RC[-4]:SalesTable[Qty])")
    ]
    assert result.edges == (FormulaEdge(0, 0, Rect(2, 10, 1, 2), "structured:SalesTable[Qty]"),)
    assert result.diagnostics == ()


def test_mixed_relative_and_defined_name_endpoints_keep_name_provenance() -> None:
    sheet = _sheet("Data", 0)
    context = ReferenceContext(
        (sheet,),
        defined_names=(
            DefinedName(
                "FixedBand",
                "'Data'!$B$2:$B$5",
                None,
                "range",
                False,
                (NameArea("Data", Rect(2, 5, 2, 2)),),
            ),
        ),
    )
    cells = tuple(FormulaCell(row, 5, f"=SUM(A{row}:FixedBand)") for row in range(3, 11))

    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 10, 1, 5),), context)

    assert [(block.rect, block.r1c1) for block in result.blocks] == [
        (Rect(3, 10, 5, 5), "=SUM(RC[-4]:FixedBand)")
    ]
    assert result.edges == (FormulaEdge(0, 0, Rect(2, 10, 1, 2), "name:FixedBand"),)
    assert result.diagnostics == ()


def test_computed_name_endpoints_do_not_fabricate_static_block_hulls() -> None:
    sheet = _sheet("Data", 0)
    context = ReferenceContext(
        (sheet,),
        defined_names=(DefinedName("Pick", "=INDEX($A:$A,1)", None, "formula", False),),
    )
    cells = tuple(FormulaCell(row, 3, f"=SUM(Pick:B{row + 2})") for row in range(3, 6))

    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 7, 1, 3),), context)

    assert [(block.rect, block.opaque) for block in result.blocks] == [(Rect(3, 5, 3, 3), True)]
    assert set(result.edges) == {
        FormulaEdge(0, 0, Rect(1, 1_048_576, 1, 1), "ref"),
        FormulaEdge(0, 0, Rect(5, 7, 2, 2), "ref"),
        FormulaEdge(0, None, None, "opaque:INDEX"),
    }
    assert [(item.code, item.row, item.col) for item in result.diagnostics] == [
        ("I_DYNAMIC_REF", 3, 3)
    ]


def test_relative_composite_inside_formula_name_remains_deliberately_opaque() -> None:
    sheet = _sheet("Data", 0)
    context = ReferenceContext(
        (sheet,),
        defined_names=(
            DefinedName(
                "FixedBand",
                "'Data'!$B$2:$B$5",
                None,
                "range",
                False,
                (NameArea("Data", Rect(2, 5, 2, 2)),),
            ),
            DefinedName("MovingBand", "=A3:FixedBand", None, "formula", False),
        ),
    )
    cells = tuple(FormulaCell(row, 5, "=MovingBand") for row in range(3, 11))

    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 10, 1, 5),), context)

    assert [(block.rect, block.opaque) for block in result.blocks] == [(Rect(3, 10, 5, 5), True)]
    assert result.edges == (FormulaEdge(0, None, None, "name-relative"),)
    assert result.diagnostics == ()


def test_dynamic_unknown_and_parse_failures_are_contained_as_opaque_blocks() -> None:
    sheet = _sheet("Containment", 0)
    context = ReferenceContext((sheet,))
    cells = (
        FormulaCell(2, 2, '=INDIRECT("A1")'),
        FormulaCell(3, 2, "=MissingName+1"),
        FormulaCell(4, 2, "=SUM(A1"),
    )

    result = analyze_sheet_formulas(
        sheet,
        cells,
        (Rect(1, 5, 1, 2),),
        context,
    )

    assert [(block.n, block.volatile, block.opaque, block.parsed) for block in result.blocks] == [
        (0, True, True, True),
        (1, False, True, True),
        (2, False, True, False),
    ]
    assert result.edges == (
        FormulaEdge(0, None, None, "opaque:INDIRECT"),
        FormulaEdge(1, None, None, "opaque:name"),
        FormulaEdge(2, None, None, "opaque:parse"),
    )
    assert [(item.code, item.row, item.col) for item in result.diagnostics] == [
        ("I_DYNAMIC_REF", 2, 2),
        ("W_UNKNOWN_NAME", 3, 2),
        ("W_PARSE", 4, 2),
    ]
    assert [dict(item.related)["block"] for item in result.diagnostics] == [
        "fblock:Containment:0",
        "fblock:Containment:1",
        "fblock:Containment:2",
    ]


def test_edges_and_diagnostics_are_deduplicated_and_sorted_by_natural_keys() -> None:
    data = _sheet("Data", 0)
    calc = _sheet("Calc", 1)
    context = ReferenceContext((data, calc))
    # Input is column-major as required by block construction, but block
    # ordinals deliberately differ from that order because they are row-first.
    cells = (
        FormulaCell(5, 2, '=INDIRECT("A1")+Unknown'),
        FormulaCell(2, 3, "=Data!Z10+Data!A1+Data!Z10+Missing"),
    )

    result = analyze_sheet_formulas(calc, cells, (Rect(1, 10, 1, 3),), context)
    repeated = analyze_sheet_formulas(calc, cells, (Rect(1, 10, 1, 3),), context)

    assert result == repeated
    assert [(block.n, block.rect) for block in result.blocks] == [
        (0, Rect(2, 2, 3, 3)),
        (1, Rect(5, 5, 2, 2)),
    ]
    # Z10 occurs twice in the formula but is one natural dependency edge.
    assert result.edges == (
        FormulaEdge(0, None, None, "opaque:name"),
        FormulaEdge(0, 0, Rect(1, 1, 1, 1), "ref"),
        FormulaEdge(0, 0, Rect(10, 10, 26, 26), "ref"),
        FormulaEdge(1, None, None, "opaque:INDIRECT"),
        FormulaEdge(1, None, None, "opaque:name"),
    )
    assert [(item.row, item.col, item.code) for item in result.diagnostics] == [
        (2, 3, "W_UNKNOWN_NAME"),
        (5, 2, "I_DYNAMIC_REF"),
        (5, 2, "W_UNKNOWN_NAME"),
    ]
    assert [dict(item.related)["block"] for item in result.diagnostics] == [
        "fblock:Calc:0",
        "fblock:Calc:1",
        "fblock:Calc:1",
    ]


def test_bare_current_row_refs_keep_one_block_but_bind_each_adjacent_table() -> None:
    sheet = _sheet("Tables", 0)
    table_a = TableBinding(
        0,
        "Tables",
        TableInfo("TableA", "TableA", "A1:B4", 1, 0, ("Input", "Result")),
    )
    table_b = TableBinding(
        0,
        "Tables",
        TableInfo("TableB", "TableB", "C1:D4", 1, 0, ("Result", "Input")),
    )
    context = ReferenceContext((sheet,), tables=(table_a, table_b))
    cells = tuple(FormulaCell(row, column, "=[@Input]") for column in (2, 3) for row in range(2, 5))

    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 4, 1, 4),), context)

    assert [(block.n, block.rect, block.r1c1) for block in result.blocks] == [
        (0, Rect(2, 4, 2, 3), "=[@Input]")
    ]
    assert result.edges == (
        FormulaEdge(0, 0, Rect(2, 4, 1, 1), "structured:TableA[Input]"),
        FormulaEdge(0, 0, Rect(2, 4, 4, 4), "structured:TableB[Input]"),
    )
    assert result.diagnostics == ()


def test_bare_data_refs_bind_both_tables_without_splitting_the_fblock() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "A1:B4", 1, 0, ("Input", "Result")),
            ),
            TableBinding(
                0,
                "Tables",
                TableInfo("TableB", "TableB", "C1:D4", 1, 0, ("Result", "Input")),
            ),
        ),
    )
    cells = tuple(
        FormulaCell(row, column, "=ROWS([#Data])") for column in (2, 3) for row in range(2, 5)
    )

    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 4, 1, 4),), context)

    assert [block.rect for block in result.blocks] == [Rect(2, 4, 2, 3)]
    assert result.edges == (
        FormulaEdge(0, 0, Rect(2, 4, 1, 2), "structured:TableA"),
        FormulaEdge(0, 0, Rect(2, 4, 3, 4), "structured:TableB"),
    )
    assert result.diagnostics == ()


def test_outer_intersection_over_bare_columns_keeps_data_not_current_row_semantics() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "A1:B4", 1, 0, ("Input", "Result")),
            ),
            TableBinding(
                0,
                "Tables",
                TableInfo("TableB", "TableB", "C1:D4", 1, 0, ("Result", "Input")),
            ),
        ),
    )
    cells = tuple(FormulaCell(row, column, "=@[Input]") for column in (2, 3) for row in range(2, 5))

    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 4, 1, 4),), context)

    assert [block.rect for block in result.blocks] == [Rect(2, 4, 2, 3)]
    assert result.edges == (
        FormulaEdge(0, 0, Rect(2, 4, 1, 1), "structured:TableA[Input]"),
        FormulaEdge(0, 0, Rect(2, 4, 4, 4), "structured:TableB[Input]"),
    )
    assert result.diagnostics == ()


def test_contextual_structured_refs_do_not_fragment_ordinary_block_edges() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "A1:B4", 1, 0, ("Input", "Result")),
            ),
            TableBinding(
                0,
                "Tables",
                TableInfo("TableB", "TableB", "C1:D4", 1, 0, ("Result", "Input")),
            ),
        ),
    )
    cells = tuple(
        FormulaCell(
            row,
            column,
            f"=[@Input]+$Z$1+{column_label(column + 3)}{row}",
        )
        for column in (2, 3)
        for row in range(2, 5)
    )

    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 4, 1, 26),), context)

    assert [block.rect for block in result.blocks] == [Rect(2, 4, 2, 3)]
    assert result.edges == (
        FormulaEdge(0, 0, Rect(1, 1, 26, 26), "ref"),
        FormulaEdge(0, 0, Rect(2, 4, 1, 1), "structured:TableA[Input]"),
        FormulaEdge(0, 0, Rect(2, 4, 4, 4), "structured:TableB[Input]"),
        FormulaEdge(0, 0, Rect(2, 4, 5, 6), "ref"),
    )
    assert result.diagnostics == ()


def test_context_tiles_translate_compound_reference_horizontally() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "A1:B4", 1, 0, ("Input", "Result")),
            ),
            TableBinding(
                0,
                "Tables",
                TableInfo("TableB", "TableB", "C1:D4", 1, 0, ("Result", "Input")),
            ),
        ),
    )
    templates = (
        "=[@Input]:{far}",
        "={far}:[@Input]",
        "=([@Input]):({far})",
        "=SUM(([@Input]):({far}))",
    )
    for template in templates:
        cells = tuple(
            FormulaCell(
                row,
                column,
                template.format(far=f"{column_label(column + 4)}{row}"),
            )
            for column in (2, 3)
            for row in range(2, 5)
        )

        result = analyze_sheet_formulas(sheet, cells, (Rect(1, 4, 1, 7),), context)

        assert [block.rect for block in result.blocks] == [Rect(2, 4, 2, 3)]
        assert result.edges == (
            FormulaEdge(0, 0, Rect(2, 4, 1, 6), "structured:TableA[Input]"),
            FormulaEdge(0, 0, Rect(2, 4, 4, 7), "structured:TableB[Input]"),
        ), template
        assert result.diagnostics == (), template


def test_compound_requirement_keeps_each_qualified_current_row_table() -> None:
    sheet = _sheet("Tables", 0)
    table_a = TableBinding(
        0,
        "Tables",
        TableInfo("TableA", "TableA", "A1:B4", 1, 0, ("Input", "Result")),
    )
    table_b = TableBinding(
        0,
        "Tables",
        TableInfo("TableB", "TableB", "C1:D4", 1, 0, ("Result", "Input")),
    )
    context = ReferenceContext((sheet,), tables=(table_a, table_b))

    requirement = structured_context_requirement(
        "(TableA[@Input]:TableB[@Input]) [@Input]",
        context,
    )

    assert requirement.uses_current_table
    assert requirement.uses_current_table_row
    assert requirement.current_row_tables == (table_a, table_b)

    for context_free in (
        "Jan:Mar!A1",
        "'Jan:Mar'!A1",
        "'[Book.xlsx]Data'!A1:F2",
        "TableA[[#Data],[Input]]",
    ):
        assert (
            structured_context_requirement(context_free, context) == StructuredContextRequirement()
        )


def test_context_tiles_translate_compound_reference_vertically() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "A1:C3", 0, 0, ("Input", "X", "Y")),
            ),
            TableBinding(
                0,
                "Tables",
                TableInfo("TableB", "TableB", "A4:C6", 0, 0, ("Input", "X", "Y")),
            ),
        ),
    )
    cells = tuple(FormulaCell(row, 3, f"=[@Input]:F{row}") for row in range(1, 7))

    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 6, 1, 6),), context)

    assert [block.rect for block in result.blocks] == [Rect(1, 6, 3, 3)]
    assert result.edges == (
        FormulaEdge(0, 0, Rect(1, 3, 1, 6), "structured:TableA[Input]"),
        FormulaEdge(0, 0, Rect(4, 6, 1, 6), "structured:TableB[Input]"),
    )
    assert result.diagnostics == ()


def test_compound_context_recovers_when_block_anchor_is_outside_table() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "A2:B4", 0, 0, ("Input", "Result")),
            ),
        ),
    )
    templates = (
        "=[@Input]:F{row}",
        "=([@Input]):(F{row})",
        "=(F{row}):([@Input])",
    )
    for template in templates:
        cells = tuple(FormulaCell(row, 2, template.format(row=row)) for row in range(1, 5))

        result = analyze_sheet_formulas(sheet, cells, (Rect(1, 4, 1, 6),), context)

        assert [block.rect for block in result.blocks] == [Rect(1, 4, 2, 2)]
        assert result.edges == (
            FormulaEdge(0, None, None, "opaque:structured"),
            FormulaEdge(0, 0, Rect(1, 1, 6, 6), "ref"),
            FormulaEdge(0, 0, Rect(2, 4, 1, 6), "structured:TableA[Input]"),
        ), template
        assert [(item.code, item.row, item.col) for item in result.diagnostics] == [
            ("W_PARSE", 1, 2)
        ], template


def test_contextual_intersection_recovers_marker_from_later_tiles() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "A2:B4", 0, 0, ("Input", "Result")),
            ),
            TableBinding(
                0,
                "Tables",
                TableInfo("TableB", "TableB", "C2:D4", 0, 0, ("Result", "Input")),
            ),
        ),
    )
    cells = tuple(
        FormulaCell(row, column, f"=([@Input] {column_label(column + 4)}{row})")
        for column in (2, 3)
        for row in range(1, 5)
    )

    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 4, 1, 7),), context)

    assert [block.rect for block in result.blocks] == [Rect(1, 4, 2, 3)]
    assert result.edges == (
        FormulaEdge(0, None, None, "opaque:ref"),
        FormulaEdge(0, None, None, "opaque:structured"),
        FormulaEdge(0, 0, Rect(1, 4, 6, 7), "ref"),
        FormulaEdge(0, 0, Rect(2, 4, 1, 1), "structured:TableA[Input]"),
        FormulaEdge(0, 0, Rect(2, 4, 4, 4), "structured:TableB[Input]"),
    )
    assert [(item.code, item.row, item.col) for item in result.diagnostics] == [("W_PARSE", 1, 2)]


def test_grouped_compound_context_retains_endpoints_after_leaving_table() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "A1:B4", 1, 0, ("Input", "Result")),
            ),
        ),
    )
    templates = (
        "=([@Input]):({far})",
        "=({far}):([@Input])",
    )
    for template in templates:
        cells = tuple(
            FormulaCell(
                row,
                column,
                template.format(far=f"{column_label(column + 4)}{row}"),
            )
            for column in (2, 3)
            for row in range(2, 7)
        )

        result = analyze_sheet_formulas(sheet, cells, (Rect(1, 6, 1, 7),), context)

        assert [block.rect for block in result.blocks] == [Rect(2, 6, 2, 3)]
        assert result.edges == (
            FormulaEdge(0, None, None, "opaque:structured"),
            FormulaEdge(0, 0, Rect(2, 4, 1, 6), "structured:TableA[Input]"),
            FormulaEdge(0, 0, Rect(2, 6, 7, 7), "ref"),
            FormulaEdge(0, 0, Rect(5, 6, 6, 6), "ref"),
        ), template
        assert [(item.code, item.row, item.col) for item in result.diagnostics] == [
            ("W_PARSE", 2, 3),
        ], template


def test_context_tiles_preserve_compound_coordinate_spill_anchor() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "A1:B4", 1, 0, ("Input", "Result")),
            ),
            TableBinding(
                0,
                "Tables",
                TableInfo("TableB", "TableB", "C1:D4", 1, 0, ("Result", "Input")),
            ),
        ),
    )
    cells = tuple(
        FormulaCell(row, column, "=[@Input]:F2#") for column in (2, 3) for row in range(2, 5)
    )

    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 4, 1, 6),), context)

    assert [block.rect for block in result.blocks] == [Rect(2, 4, 2, 3)]
    assert result.edges == (
        FormulaEdge(0, 0, Rect(2, 4, 1, 1), "structured:TableA[Input]"),
        FormulaEdge(0, 0, Rect(2, 4, 4, 4), "structured:TableB[Input]"),
        FormulaEdge(0, 0, Rect(2, 2, 6, 6), "spill"),
    )
    assert result.diagnostics == ()


def test_contextual_structured_refs_inside_names_keep_name_relative_opacity() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        defined_names=(DefinedName("CurrentInput", "=[@Input]", None, "formula", False),),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "B1:C4", 1, 0, ("Input", "Result")),
            ),
        ),
    )
    cells = tuple(
        FormulaCell(row, column, "=CurrentInput") for column in (1, 2) for row in range(2, 5)
    )

    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 4, 1, 3),), context)

    assert [(block.rect, block.opaque) for block in result.blocks] == [(Rect(2, 4, 1, 2), True)]
    assert result.edges == (
        FormulaEdge(0, None, None, "name-relative"),
        FormulaEdge(0, None, None, "opaque:structured"),
    )
    assert [(item.code, item.row, item.col) for item in result.diagnostics] == [("W_PARSE", 2, 1)]


def test_qualified_current_row_block_is_tiled_at_table_data_boundaries() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "A1:B4", 1, 0, ("Input", "Result")),
            ),
        ),
    )
    cells = tuple(FormulaCell(row, 3, "=TableA[@Input]") for row in range(1, 6))

    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 5, 1, 3),), context)

    assert [(block.rect, block.opaque) for block in result.blocks] == [(Rect(1, 5, 3, 3), True)]
    assert result.edges == (
        FormulaEdge(0, None, None, "opaque:structured"),
        FormulaEdge(0, 0, Rect(2, 4, 1, 1), "structured:TableA[Input]"),
    )
    assert [(item.code, item.row, item.col) for item in result.diagnostics] == [("W_PARSE", 1, 3)]


def test_context_tiles_union_dynamic_index_semantics_across_one_block() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "B1:C4", 1, 0, ("Input", "Result")),
            ),
        ),
    )
    cells = tuple(
        FormulaCell(row, col, "=SUM(INDEX([#Data],1):$Z$1)")
        for col in (1, 2)
        for row in range(2, 5)
    )

    result = analyze_sheet_formulas(
        sheet,
        cells,
        (Rect(1, 4, 1, 26),),
        context,
    )

    assert [(block.rect, block.opaque) for block in result.blocks] == [(Rect(2, 4, 1, 2), True)]
    assert result.edges == (
        FormulaEdge(0, None, None, "opaque:INDEX"),
        FormulaEdge(0, None, None, "opaque:structured"),
        FormulaEdge(0, 0, Rect(1, 1, 26, 26), "ref"),
        FormulaEdge(0, 0, Rect(2, 4, 2, 3), "structured:TableA"),
    )
    assert [(item.code, item.row, item.col) for item in result.diagnostics] == [
        ("W_PARSE", 2, 1),
        ("I_DYNAMIC_REF", 2, 2),
    ]


def test_context_tiles_union_dynamic_choose_semantics_across_table_rows() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "A1:B4", 1, 0, ("Input", "Result")),
            ),
        ),
    )
    cells = tuple(FormulaCell(row, 3, "=CHOOSE(1,TableA[@Input],10)") for row in range(1, 5))

    result = analyze_sheet_formulas(
        sheet,
        cells,
        (Rect(1, 4, 1, 3),),
        context,
    )

    assert [(block.rect, block.opaque) for block in result.blocks] == [(Rect(1, 4, 3, 3), True)]
    assert result.edges == (
        FormulaEdge(0, None, None, "opaque:CHOOSE"),
        FormulaEdge(0, None, None, "opaque:structured"),
        FormulaEdge(0, 0, Rect(2, 4, 1, 1), "structured:TableA[Input]"),
    )
    assert [(item.code, item.row, item.col) for item in result.diagnostics] == [
        ("W_PARSE", 1, 3),
        ("I_DYNAMIC_REF", 2, 3),
    ]


def test_bare_current_row_block_distinguishes_data_body_from_total_row() -> None:
    sheet = _sheet("Tables", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Tables",
                TableInfo("TableA", "TableA", "A1:B5", 1, 1, ("Input", "Result")),
            ),
        ),
    )
    cells = tuple(FormulaCell(row, 2, "=[@Input]") for row in range(2, 6))

    result = analyze_sheet_formulas(sheet, cells, (Rect(1, 5, 1, 2),), context)

    assert [(block.rect, block.opaque) for block in result.blocks] == [(Rect(2, 5, 2, 2), True)]
    assert result.edges == (
        FormulaEdge(0, None, None, "opaque:structured"),
        FormulaEdge(0, 0, Rect(2, 4, 1, 1), "structured:TableA[Input]"),
    )
    assert [(item.code, item.row, item.col) for item in result.diagnostics] == [("W_PARSE", 5, 2)]


def test_many_adjacent_table_contexts_remain_one_block_with_all_edges() -> None:
    sheet = _sheet("ManyTables", 0)
    table_count = 128
    tables = tuple(
        TableBinding(
            0,
            "ManyTables",
            TableInfo(
                f"T{column:03d}",
                f"T{column:03d}",
                f"{column_label(column)}1:{column_label(column)}3",
                1,
                0,
                ("Value",),
            ),
        )
        for column in range(1, table_count + 1)
    )
    cells = tuple(
        FormulaCell(row, column, "=[@Value]")
        for column in range(1, table_count + 1)
        for row in (2, 3)
    )

    result = analyze_sheet_formulas(
        sheet,
        cells,
        (Rect(1, 3, 1, table_count),),
        ReferenceContext((sheet,), tables=tables),
    )

    assert [block.rect for block in result.blocks] == [Rect(2, 3, 1, table_count)]
    assert len(result.edges) == table_count
    assert result.edges[0] == FormulaEdge(
        0,
        0,
        Rect(2, 3, 1, 1),
        "structured:T001[Value]",
    )
    assert result.edges[-1] == FormulaEdge(
        0,
        0,
        Rect(2, 3, table_count, table_count),
        "structured:T128[Value]",
    )
    assert result.diagnostics == ()
