from __future__ import annotations

import asyncio
from collections.abc import Iterable
from concurrent.futures import ThreadPoolExecutor
from contextvars import copy_context
from hashlib import sha256
from time import perf_counter
from types import MappingProxyType

import pytest

import excel_lsp.core.formulas.analysis as analysis_module
import excel_lsp.core.formulas.references as references_module
from excel_lsp.core.formulas import (
    BUILTIN_FUNCTIONS,
    MODERN_FUNCTIONS,
    FormulaAnchor,
    FormulaIssue,
    FormulaSyntaxError,
    FormulaToken,
    ReferenceContext,
    TableBinding,
    analyze_formula,
    classify_ref,
    compatibility_function_identifier,
    normalize_function_name,
    resolve_reference,
    tokenize_formula,
)
from excel_lsp.core.models import (
    DefinedName,
    NameArea,
    Rect,
    SheetDescriptor,
    TableInfo,
)


def _sheet(name: str, order: int) -> SheetDescriptor:
    return SheetDescriptor(
        order=order,
        name=name,
        sheet_id=order + 1,
        rel_id=f"rId{order + 1}",
        xml_part=f"xl/worksheets/sheet{order + 1}.xml",
        kind="worksheet",
    )


@pytest.fixture
def context() -> ReferenceContext:
    sheets = (
        _sheet("Data", 0),
        _sheet("Calc", 1),
        _sheet("My Sheet", 2),
        _sheet("O'Brien", 3),
        _sheet("Jan 24", 4),
        _sheet("Feb 24", 5),
        _sheet("Mar 24", 6),
    )
    names = (
        DefinedName(
            "Rate",
            "'Data'!$B$2",
            None,
            "range",
            False,
            (NameArea("Data", Rect(2, 2, 2, 2)),),
        ),
        DefinedName(
            "Rate",
            "'Calc'!$C$3",
            1,
            "range",
            False,
            (NameArea("Calc", Rect(3, 3, 3, 3)),),
        ),
        DefinedName(
            "Buckets",
            "'Data'!$A$2:$A$3,'Data'!$C$2:$C$3",
            None,
            "multi_range",
            False,
            (
                NameArea("Data", Rect(2, 3, 1, 1)),
                NameArea("Data", Rect(2, 3, 3, 3)),
            ),
        ),
        DefinedName("PiValue", "=3.14", None, "constant", False),
        DefinedName("AbsoluteFormula", "=$A$1+'Data'!$B$2", None, "formula", False),
        DefinedName("RelativeFormula", "=A1+$B$2", None, "formula", False),
        DefinedName("Double", "=_xlfn.LAMBDA(x,x*$B$2)", None, "lambda", False),
    )
    table = TableInfo(
        name="SalesTable",
        display_name="SalesTable",
        ref="A1:C6",
        header_rows=1,
        totals_rows=1,
        columns=("Item", "Qty", "Price"),
    )
    return ReferenceContext(
        sheets,
        names,
        (TableBinding(0, "Data", table),),
        MappingProxyType(
            {
                1: "budget.xlsx",
                2: "https://user:secret@example.test/private/payroll.xlsx?token=abc#fragment",
            }
        ),
    )


def _anchor(sheet: str = "Data", order: int = 0, row: int = 3, col: int = 2) -> FormulaAnchor:
    return FormulaAnchor(order, sheet, row, col)


def _codes(issues: Iterable[FormulaIssue]) -> list[str]:
    return [issue.code for issue in issues]


def test_tokenizer_preserves_spills_structured_hashes_quotes_and_whitespace() -> None:
    formula = "=A1#   +\n Name# + Table1[#All] + 'Hash#Sheet'!A1 + \"x#y\""

    tokens = tokenize_formula(formula)

    assert "".join(token.value for token in tokens) == formula[1:]
    assert [token.value for token in tokens if token.subtype == "RANGE"] == [
        "A1#",
        "Name#",
        "Table1[#All]",
        "'Hash#Sheet'!A1",
    ]
    assert [token.value for token in tokens if token.type == "WHITE-SPACE"] == [
        "   ",
        "\n ",
        " ",
        " ",
        " ",
        " ",
        " ",
        " ",
    ]


def test_tokenizer_rejects_spill_garbage_and_unclosed_functions() -> None:
    with pytest.raises(FormulaSyntaxError):
        tokenize_formula("=A1##")
    with pytest.raises(FormulaSyntaxError, match="unclosed"):
        tokenize_formula("=SUM(")


def test_function_normalization_and_frozen_modern_supplement() -> None:
    assert normalize_function_name("_xlfn._xlws.FILTER(") == "FILTER"
    assert normalize_function_name("@_xlfn.SINGLE(") == "SINGLE"
    assert normalize_function_name("A1:@INDEX(") == "INDEX"
    assert compatibility_function_identifier("_xlfn._xlws.DoubleIt(") == "DoubleIt"
    assert {
        "LET",
        "LAMBDA",
        "XLOOKUP",
        "FILTER",
        "RANDARRAY",
        "IFS",
        "SWITCH",
        "CONCAT",
    } <= BUILTIN_FUNCTIONS


def test_frozen_modern_function_inventory_has_stable_size_and_digest() -> None:
    serialized = "\n".join(sorted(MODERN_FUNCTIONS)) + "\n"

    assert len(MODERN_FUNCTIONS) == 171
    assert sha256(serialized.encode()).hexdigest() == (
        "e45b2862249d78b5b65a7ffe96c5681c9a97aeee94335adb323be330d03bdb0f"
    )


def test_every_frozen_modern_function_avoids_unknown_name_classification(
    context: ReferenceContext,
) -> None:
    for function_name in sorted(MODERN_FUNCTIONS):
        result = analyze_formula(
            f"=_xlfn.{function_name}(A1)",
            anchor=_anchor(),
            context=context,
        )
        assert "W_UNKNOWN_NAME" not in _codes(result.issues), function_name
        assert "opaque:name" not in {reference.via for reference in result.references}, (
            function_name
        )


def test_unknown_prefixed_callable_still_uses_defined_name_resolution(
    context: ReferenceContext,
) -> None:
    result = analyze_formula(
        "=_xlfn.NotARealFunction(A1)",
        anchor=_anchor(),
        context=context,
    )

    assert _codes(result.issues) == ["W_UNKNOWN_NAME"]
    assert [reference.via for reference in result.references] == [
        "opaque:name",
        "ref",
    ]


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("A1", Rect(1, 1, 1, 1)),
        ("$B$2", Rect(2, 2, 2, 2)),
        ("C3:D9", Rect(3, 9, 3, 4)),
        ("B:B", Rect(1, 1_048_576, 2, 2)),
        ("7:7", Rect(7, 7, 1, 16_384)),
    ],
)
def test_plain_reference_matrix_resolves_exactly(
    context: ReferenceContext,
    text: str,
    expected: Rect,
) -> None:
    anchor = _anchor()
    result = classify_ref(text, anchor=anchor, context=context)

    assert result.issues == ()
    assert len(result.references) == 1
    reference = result.references[0]
    assert reference.dst_sheet_name == "Data"
    assert reference.geometry is not None
    assert resolve_reference(reference.geometry, anchor.cell) == expected


def test_quoted_sheets_apostrophes_and_quoted_3d_spans(context: ReferenceContext) -> None:
    anchor = _anchor()
    quoted = classify_ref("'My Sheet'!C3:D4", anchor=anchor, context=context)
    apostrophe = classify_ref("'O''Brien'!$A$1", anchor=anchor, context=context)
    three_d = classify_ref("'Mar 24:Jan 24'!B2", anchor=anchor, context=context)

    assert quoted.references[0].dst_sheet_name == "My Sheet"
    assert apostrophe.references[0].dst_sheet_name == "O'Brien"
    assert [reference.dst_sheet_name for reference in three_d.references] == [
        "Jan 24",
        "Feb 24",
        "Mar 24",
    ]
    assert {reference.via for reference in three_d.references} == {"3d"}


def test_unquoted_3d_span_expands_in_workbook_order() -> None:
    context = ReferenceContext(
        (
            _sheet("Jan", 0),
            _sheet("Feb", 1),
            _sheet("Mar", 2),
        )
    )

    result = classify_ref("Jan:Mar!B2", anchor=_anchor("Jan", 0), context=context)

    assert result.issues == ()
    assert [reference.dst_sheet_name for reference in result.references] == [
        "Jan",
        "Feb",
        "Mar",
    ]
    assert {reference.via for reference in result.references} == {"3d"}


def test_external_references_use_numeric_map_or_preserve_direct_book_label(
    context: ReferenceContext,
) -> None:
    numeric = classify_ref("[1]Data!A1", anchor=_anchor(), context=context)
    direct = classify_ref("'[budget 2025.xlsx]Q1'!A1", anchor=_anchor(), context=context)
    credentialed = classify_ref("[2]Data!A1", anchor=_anchor(), context=context)

    assert [(reference.geometry, reference.via) for reference in numeric.references] == [
        (None, "external:[budget.xlsx]")
    ]
    assert direct.references[0].via == "external:[budget 2025.xlsx]"
    assert credentialed.references[0].via == "external:[payroll.xlsx]"
    assert not any(
        secret in credentialed.references[0].via
        for secret in ("user", "secret", "private", "token", "fragment")
    )
    assert not numeric.opaque and not direct.opaque


def test_defined_names_are_scope_aware_and_multi_area(context: ReferenceContext) -> None:
    global_rate = classify_ref("Rate", anchor=_anchor(), context=context)
    local_rate = classify_ref("Rate", anchor=_anchor("Calc", 1), context=context)
    buckets = classify_ref("Buckets", anchor=_anchor(), context=context)
    constant = classify_ref("PiValue", anchor=_anchor(), context=context)

    assert global_rate.references[0].dst_sheet_name == "Data"
    assert local_rate.references[0].dst_sheet_name == "Calc"
    assert len(buckets.references) == 2
    assert {reference.via for reference in buckets.references} == {"name:Buckets"}
    assert constant.references == ()


def test_formula_names_resolve_absolute_refs_and_make_relative_refs_opaque(
    context: ReferenceContext,
) -> None:
    absolute = classify_ref("AbsoluteFormula", anchor=_anchor(), context=context)
    relative = classify_ref("RelativeFormula", anchor=_anchor(), context=context)

    assert [reference.via for reference in absolute.references] == ["ref", "ref"]
    assert [reference.via for reference in relative.references] == [
        "name-relative",
        "ref",
    ]
    assert not absolute.opaque
    assert relative.opaque


def test_spill_and_implicit_intersection_target_the_definition_anchor(
    context: ReferenceContext,
) -> None:
    anchor = _anchor()
    cell_spill = classify_ref("A1#", anchor=anchor, context=context)
    name_spill = classify_ref("Rate#", anchor=anchor, context=context)
    intersection = classify_ref("@A1", anchor=anchor, context=context)

    assert cell_spill.references[0].via == "spill"
    assert name_spill.references[0].via == "spill"
    assert intersection.references[0].via == "ref"
    assert not _codes((*cell_spill.issues, *name_spill.issues, *intersection.issues))


@pytest.mark.parametrize(
    ("text", "expected"),
    (
        ("A1#:B5", (("spill", Rect(1, 1, 1, 1)), ("ref", Rect(5, 5, 2, 2)))),
        ("A1:B5#", (("ref", Rect(1, 1, 1, 1)), ("spill", Rect(5, 5, 2, 2)))),
        ("@A1:B5", (("ref", Rect(1, 5, 1, 2)),)),
        ("A1:@B5", (("ref", Rect(1, 5, 1, 2)),)),
        (
            "@A1#:@B5#",
            (("spill", Rect(1, 1, 1, 1)), ("spill", Rect(5, 5, 2, 2))),
        ),
    ),
)
def test_modern_operators_bind_to_either_range_endpoint_without_warnings(
    context: ReferenceContext,
    text: str,
    expected: tuple[tuple[str, Rect], ...],
) -> None:
    anchor = _anchor()
    result = classify_ref(text, anchor=anchor, context=context)

    assert result.issues == ()
    actual = tuple(
        (reference.via, resolve_reference(reference.geometry, anchor.cell))
        for reference in result.references
        if reference.geometry is not None
    )
    assert actual == expected


def test_modern_range_operators_expand_quoted_3d_qualifiers_per_endpoint(
    context: ReferenceContext,
) -> None:
    anchor = _anchor()
    result = classify_ref(
        "'Jan 24:Mar 24'!A1#:@B5",
        anchor=anchor,
        context=context,
    )

    assert result.issues == ()
    assert [(reference.dst_sheet_name, reference.via) for reference in result.references] == [
        ("Jan 24", "spill"),
        ("Feb 24", "spill"),
        ("Mar 24", "spill"),
        ("Jan 24", "3d"),
        ("Feb 24", "3d"),
        ("Mar 24", "3d"),
    ]


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        ("SalesTable[Qty]", Rect(2, 5, 2, 2)),
        ("SalesTable[@Qty]", Rect(3, 3, 2, 2)),
        ("SalesTable[[#Totals],[Qty]]", Rect(6, 6, 2, 2)),
        ("SalesTable[#Headers]", Rect(1, 1, 1, 3)),
        ("SalesTable[#All]", Rect(1, 6, 1, 3)),
        ("[@Qty]", Rect(3, 3, 2, 2)),
    ],
)
def test_structured_reference_matrix(
    context: ReferenceContext,
    text: str,
    expected: Rect,
) -> None:
    anchor = _anchor()
    result = classify_ref(text, anchor=anchor, context=context)

    assert result.issues == ()
    reference = result.references[0]
    assert reference.geometry is not None
    assert resolve_reference(reference.geometry, anchor.cell) == expected
    assert reference.via.startswith("structured:SalesTable")


def test_invalid_structured_columns_and_empty_table_sections_are_contained() -> None:
    sheet = _sheet("Data", 0)
    tables = (
        TableBinding(
            0,
            "Data",
            TableInfo("NoData", "NoData", "A1:B1", 1, 0, ("A", "B")),
        ),
        TableBinding(
            0,
            "Data",
            TableInfo("NoHeader", "NoHeader", "D1:E2", 0, 0, ("D", "E")),
        ),
        TableBinding(
            0,
            "Data",
            TableInfo("NoTotals", "NoTotals", "G1:H2", 1, 0, ("G", "H")),
        ),
    )
    context = ReferenceContext((sheet,), tables=tables)

    for text in (
        "NoData[Missing]",
        "NoData[[A]:[B]]",
        "NoData[#Data]",
        "NoHeader[#Headers]",
        "NoTotals[#Totals]",
    ):
        result = classify_ref(text, anchor=_anchor(), context=context)

        assert [reference.via for reference in result.references] == ["opaque:structured"]
        assert _codes(result.issues) == ["W_PARSE"]
        assert result.opaque


def test_qualified_current_row_reference_works_beside_table_on_matching_row(
    context: ReferenceContext,
) -> None:
    beside = classify_ref(
        "SalesTable[@Qty]",
        anchor=_anchor(row=3, col=4),
        context=context,
    )
    below = classify_ref(
        "SalesTable[@Qty]",
        anchor=_anchor(row=7, col=4),
        context=context,
    )

    assert beside.references[0].geometry is not None
    assert resolve_reference(beside.references[0].geometry, _anchor(row=3, col=4).cell) == Rect(
        3, 3, 2, 2
    )
    assert beside.issues == ()
    assert _codes(below.issues) == ["W_PARSE"]


@pytest.mark.parametrize(
    ("header", "text", "expected"),
    [
        ("A]B", "Esc[[A']B]]", Rect(2, 4, 1, 1)),
        ("A]B", "Esc[[#Data],[A']B]]", Rect(2, 4, 1, 1)),
        ("A]B", "Esc[@[A']B]]", Rect(3, 3, 1, 1)),
        ("A]B", "[@[A']B]]", Rect(3, 3, 1, 1)),
        ("A#B", "Esc[[A'#B]]", Rect(2, 4, 1, 1)),
        ("A@B", "Esc[[A'@B]]", Rect(2, 4, 1, 1)),
        ("A'B", "Esc[[A''B]]", Rect(2, 4, 1, 1)),
        ("A[B", "Esc[[A'[B]]", Rect(2, 4, 1, 1)),
        ("A[B", "Esc[A'[B]", Rect(2, 4, 1, 1)),
    ],
)
def test_structured_column_special_characters_require_and_decode_escapes(
    header: str,
    text: str,
    expected: Rect,
) -> None:
    sheet = _sheet("Escaped", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Escaped",
                TableInfo("Esc", "Esc", "A1:B5", 1, 1, (header, "Other")),
            ),
        ),
    )
    anchor = FormulaAnchor(0, "Escaped", 3, 2)

    result = classify_ref(text, anchor=anchor, context=context)

    assert result.issues == ()
    assert result.references[0].geometry is not None
    assert resolve_reference(result.references[0].geometry, anchor.cell) == expected

    analyzed = analyze_formula(f"={text}+A1#", anchor=anchor, context=context)

    assert analyzed.issues == ()
    assert [(reference.token, reference.via) for reference in analyzed.references] == [
        (text, f"structured:Esc[{header}]"),
        ("A1#", "spill"),
    ]
    assert analyzed.references[0].geometry is not None
    assert resolve_reference(analyzed.references[0].geometry, anchor.cell) == expected


@pytest.mark.parametrize(
    ("header", "text"),
    [
        ("A]B", "Esc[A]B]"),
        ("A#B", "Esc[A#B]"),
        ("A@B", "Esc[A@B]"),
        ("A'B", "Esc[A'B]"),
        ("A#B", "Esc[[A#B]]"),
        ("A@B", "Esc[[A@B]]"),
        ("A'B", "Esc[[A'B]]"),
    ],
)
def test_unescaped_structured_column_special_characters_are_contained(
    header: str,
    text: str,
) -> None:
    sheet = _sheet("Escaped", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Escaped",
                TableInfo("Esc", "Esc", "A1:B5", 1, 1, (header, "Other")),
            ),
        ),
    )

    result = classify_ref(
        text,
        anchor=FormulaAnchor(0, "Escaped", 3, 2),
        context=context,
    )

    assert _codes(result.issues) == ["W_PARSE"]
    assert [reference.via for reference in result.references] == ["opaque:structured"]


@pytest.mark.parametrize(
    ("table_name", "header", "formula"),
    (
        ("NewlineTable", "Line\nBreak", "=SUM(NewlineTable[Line\nBreak])"),
        ("QuoteTable", 'He"ader', '=SUM(QuoteTable[He"ader])+SUM(D1#)'),
    ),
)
def test_structured_header_literals_do_not_corrupt_later_formula_lexing(
    table_name: str,
    header: str,
    formula: str,
) -> None:
    sheet = _sheet("Structured", 0)
    context = ReferenceContext(
        (sheet,),
        tables=(
            TableBinding(
                0,
                "Structured",
                TableInfo(table_name, table_name, "A1:A5", 1, 1, (header,)),
            ),
        ),
    )
    anchor = FormulaAnchor(0, "Structured", 3, 1)

    result = analyze_formula(formula, anchor=anchor, context=context)

    assert result.issues == ()
    assert result.references[0].token == f"{table_name}[{header}]"
    assert result.references[0].via == f"structured:{table_name}[{header}]"
    if table_name == "QuoteTable":
        assert result.references[1].token == "D1#"
        assert result.references[1].via == "spill"


def test_unknown_names_and_unsupported_structured_selectors_are_contained(
    context: ReferenceContext,
) -> None:
    unknown = classify_ref("Missing", anchor=_anchor(), context=context)
    unsupported = classify_ref("SalesTable[#Bogus]", anchor=_anchor(), context=context)

    assert _codes(unknown.issues) == ["W_UNKNOWN_NAME"]
    assert unknown.references[0].via == "opaque:name"
    assert _codes(unsupported.issues) == ["W_PARSE"]
    assert unsupported.references[0].via == "opaque:structured"

    outside = classify_ref(
        "SalesTable[@Qty]",
        anchor=_anchor("Calc", 1),
        context=context,
    )
    assert _codes(outside.issues) == ["W_PARSE"]


@pytest.mark.parametrize(
    "formula",
    [
        "=_xlfn.LET(x,A1,y,x+1,x+y)",
        "=_xlfn.LAMBDA(x,y,x+y)(A1,B1)",
        "=LET(x,A1,LAMBDA(y,x+y)(B1))",
        "=LET(x,{1,2;3,4},SUM(x))",
        "=LET(f,LAMBDA(x,x+1),f(A1))",
        "=_xlfn.XLOOKUP(A1,B:B,C:C)",
        "=_xlfn._xlws.FILTER(A1:B5,B1:B5>0)",
        "=A1#+@B1",
    ],
)
def test_i20_modern_constructs_have_no_spurious_parse_or_name_warnings(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert not ({"W_PARSE", "W_UNKNOWN_NAME"} & set(_codes(result.issues)))


def test_let_binding_is_not_visible_inside_its_own_value() -> None:
    sheet = _sheet("Data", 0)
    outer_name = DefinedName(
        "x",
        "'Data'!$A$10",
        None,
        "range",
        False,
        (NameArea("Data", Rect(10, 10, 1, 1)),),
    )
    context = ReferenceContext((sheet,), (outer_name,))

    direct = analyze_formula("=LET(x,x+1,x)", anchor=_anchor(), context=context)
    renamed = analyze_formula("=LET(y,x+1,y)", anchor=_anchor(), context=context)

    for result in (direct, renamed):
        assert [(reference.token, reference.via) for reference in result.references] == [
            ("x", "name:x")
        ]
        assert result.issues == ()


def test_let_bindings_are_lexical_and_visible_to_later_values() -> None:
    sheet = _sheet("Data", 0)
    outer_name = DefinedName(
        "x",
        "'Data'!$A$10",
        None,
        "range",
        False,
        (NameArea("Data", Rect(10, 10, 1, 1)),),
    )
    context = ReferenceContext((sheet,), (outer_name,))

    nested = analyze_formula(
        "=LET(x,A1,LET(x,x+1,x))",
        anchor=_anchor(),
        context=context,
    )
    later_value = analyze_formula(
        "=LET(x,A1,y,x+1,y)",
        anchor=_anchor(),
        context=context,
    )

    for result in (nested, later_value):
        assert [(reference.token, reference.via) for reference in result.references] == [
            ("A1", "ref")
        ]
        assert result.issues == ()


@pytest.mark.parametrize(
    ("formula", "dynamic_via"),
    (
        ("=LET(r,A1,SUM(INDEX(r,1):D5))", "opaque:INDEX"),
        ("=SUM(LET(r,A1,CHOOSE(1,r,10)):D5)", "opaque:CHOOSE"),
        ("=SUM(INDEX(LET(r,A1,r),1):D5)", "opaque:INDEX"),
        ("=SUM(CHOOSE(1,LET(r,A1,r),10):D5)", "opaque:CHOOSE"),
    ),
)
def test_let_preserves_typed_reference_identity(
    context: ReferenceContext,
    formula: str,
    dynamic_via: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert dynamic_via in [reference.via for reference in result.references]
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=LET(r,42,SUM(INDEX(r,1):D5))",
        "=SUM(CHOOSE(1,LET(r,42,r),10):D5)",
    ),
)
def test_let_scalar_bindings_do_not_promote_reference_functions(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert not any(reference.via.startswith("opaque:") for reference in result.references)
    assert result.issues == ()


def test_let_accepts_missing_final_value_but_rejects_invalid_or_duplicate_names(
    context: ReferenceContext,
) -> None:
    missing_final = analyze_formula(
        "=LET(x,1,)",
        anchor=_anchor(),
        context=context,
    )
    dotted = analyze_formula(
        "=LET(a.b,1,a.b)",
        anchor=_anchor(),
        context=context,
    )
    duplicate_let = analyze_formula(
        "=LET(x,1,x,2,x)",
        anchor=_anchor(),
        context=context,
    )
    duplicate_lambda = analyze_formula(
        "=LAMBDA(x,x,x)",
        anchor=_anchor(),
        context=context,
    )

    assert missing_final.references == ()
    assert missing_final.issues == ()
    assert not missing_final.opaque
    for result in (dotted, duplicate_let, duplicate_lambda):
        assert "opaque:parse" in [reference.via for reference in result.references]
        assert "W_PARSE" in _codes(result.issues)
        assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=LET(R1C1,$B$2,R1C1)",
        "=LET(RC,$B$2,RC)",
        "=LET(R1C,$B$2,R1C)",
        "=LET(c,$B$2,c)",
        "=LET(r,$B$2,r)",
        "=LET(rr,$B$2,rr)",
        "=LET(cc,$B$2,cc)",
        "=LET(XFE1,$B$2,XFE1)",
        "=LET(XFD1048577,$B$2,XFD1048577)",
        "=LET(ZZZ9999999,$B$2,ZZZ9999999)",
        "=LET(foo_2,$B$2,foo_2)",
        "=LET(_,$B$2,_)",
        "=LET(_x,$B$2,_x)",
        "=LET(é,$B$2,é)",
        "=LET(λ,$B$2,λ)",
        "=LET(名,$B$2,名)",
        "=LET(\\,$B$2,\\)",
        "=LAMBDA(R1C1,R1C1)($B$2)",
        "=LAMBDA(RC,RC)($B$2)",
        "=LAMBDA(c,c)($B$2)",
        "=LAMBDA(r,r)($B$2)",
        "=LAMBDA(XFE1,XFE1)($B$2)",
    ),
)
def test_local_declarations_accept_r1c1_spelling_that_is_not_an_a1_reference(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert [(reference.token, reference.via) for reference in result.references] == [
        ("$B$2", "ref")
    ]
    assert result.issues == ()
    assert not result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=LET(A1,$B$2,A1)",
        "=LET(foo2,$B$2,foo2)",
        "=LET(abc123,$B$2,abc123)",
        "=LET(RC1,$B$2,RC1)",
        "=LET(XFD1048576,$B$2,XFD1048576)",
        "=LAMBDA(A1,A1)($B$2)",
        "=LAMBDA(XFD1048576,XFD1048576)($B$2)",
    ),
)
def test_local_declarations_reject_valid_a1_cell_spelling(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert "opaque:parse" in {reference.via for reference in result.references}
    assert "W_PARSE" in _codes(result.issues)
    assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=LET(1x,$B$2,1x)",
        "=LET(foo bar,$B$2,foo bar)",
        "=LET(a-b,$B$2,a-b)",
        "=LET(a+b,$B$2,a+b)",
        "=LET(.x,$B$2,.x)",
        "=LET(x.,$B$2,x.)",
        "=LET(@x,$B$2,@x)",
        "=LET(x#,$B$2,x#)",
        "=LET(a.b,$B$2,a.b)",
        "=LET(x,1,X,2,x)",
        "=_xlfn.LET(_xlpm.x,1,_xlpm.X,2,_xlpm.x)",
        "=_xlfn.LET(_xlpm.x,1,x,2,_xlpm.x)",
    ),
)
def test_local_declarations_reject_ui_invalid_or_duplicate_spellings(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert "opaque:parse" in {reference.via for reference in result.references}
    assert "W_PARSE" in _codes(result.issues)
    assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=_xlfn.LET(_xlpm.x,$B$2,_xlpm.x)",
        "=_xlfn.LET(_xlpm.A1,$B$2,_xlpm.A1)",
        "=_xlfn.LET(_xlpm.R1C1,$B$2,_xlpm.R1C1)",
        "=_xlfn.LAMBDA(_xlpm.x,_xlpm.x)($B$2)",
        "=_xlfn.LAMBDA(_xlpm.A1,_xlpm.A1)($B$2)",
        "=_xlfn.LAMBDA(_xlpm.R1C1,_xlpm.R1C1)($B$2)",
    ),
)
def test_excel_stored_local_parameter_prefix_is_suppressed(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert [(reference.token, reference.via) for reference in result.references] == [
        ("$B$2", "ref")
    ]
    assert result.issues == ()
    assert not result.opaque


def test_excel_stored_callable_prefix_is_private_implementation_detail(
    context: ReferenceContext,
) -> None:
    result = analyze_formula(
        "=_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.r,_xlpm.r),SUM(CHOOSE(1,_xlpm.f(A1),10):D5))",
        anchor=_anchor(),
        context=context,
    )

    assert result.function_calls == ("LET", "LAMBDA", "SUM", "CHOOSE", "F")
    assert "_XLPM" not in repr(result)
    assert "opaque:CHOOSE" in {reference.via for reference in result.references}
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]


@pytest.mark.parametrize(
    "formula",
    (
        "=LET(r,B2,SUM(r:C3))",
        "=LET(r,B2,SUM(C3:r))",
        "=LET(r,B2,SUM(r C3))",
        "=_xlfn.LET(_xlpm.r,B2,SUM(_xlpm.r:C3))",
        "=_xlfn.LET(_xlpm.r,B2,SUM(C3:_xlpm.r))",
        "=_xlfn.LET(_xlpm.r,B2,SUM(_xlpm.r C3))",
    ),
)
def test_local_reference_composites_remain_visible_and_conservative(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    vias = [reference.via for reference in result.references]
    assert vias.count("ref") == 2
    assert vias.count("opaque:ref") == 1
    assert result.issues == ()
    assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=LET(r,B2,s,C3,SUM(r:s))",
        "=LET(rr,B2,ss,C3,SUM(rr:ss))",
        "=LET(r,B2,SUM(r:S))",
        "=LET(s,C3,SUM(R:s))",
    ),
)
def test_column_shaped_raw_local_colons_take_lexical_precedence(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert "opaque:ref" in {reference.via for reference in result.references}
    assert result.issues == ()
    assert result.opaque


def test_qualified_3d_colon_is_not_reinterpreted_as_a_lexical_local() -> None:
    context = ReferenceContext(
        (_sheet("Data", 0), _sheet("SheetOne", 1), _sheet("SheetTwo", 2)),
    )
    result = analyze_formula(
        "=LET(SheetOne,B2,SUM(SheetOne:SheetTwo!A1))",
        anchor=_anchor(),
        context=context,
    )

    assert [reference.via for reference in result.references].count("ref") == 1
    assert [reference.via for reference in result.references].count("3d") == 2
    assert "opaque:ref" not in {reference.via for reference in result.references}
    assert result.issues == ()


@pytest.mark.parametrize(
    "formula",
    (
        "=SUM(LET(r,42,s,43,r:s):D5)",
        "=SUM(LET(rr,42,ss,43,rr:ss):D5)",
        "=SUM(LET(r,42,r:S):D5)",
    ),
)
def test_scalar_column_shaped_raw_local_colons_do_not_gain_reference_identity(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert "opaque:ref" in {reference.via for reference in result.references}
    assert "opaque:LET" not in {reference.via for reference in result.references}
    assert "I_DYNAMIC_REF" not in _codes(result.issues)


@pytest.mark.parametrize(
    "formula",
    (
        "=SUM(LET(r,B2,s,C3,r:s):D5)",
        "=SUM(LET(rr,B2,ss,C3,rr:ss):D5)",
        "=SUM(LET(r,B2,r:S):D5)",
    ),
)
def test_reference_column_shaped_raw_local_colons_retain_outer_attribution(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert "opaque:ref" in {reference.via for reference in result.references}
    assert "opaque:LET" in {reference.via for reference in result.references}
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]


@pytest.mark.parametrize(
    "formula",
    (
        "=LET(r,B2,SUM((r):C3))",
        "=LET(r,B2,SUM(((r)):C3))",
        "=LET(r,B2,SUM((C3):r))",
        "=LET(r,B2,SUM((r):S))",
        "=LET(r,B2,SUM(S:(r)))",
        "=LET(r,B2,SUM((r):3))",
        "=LET(r,B2,SUM(3:(r)))",
        "=LET(r,B2,SUM(D5:(r)))",
        "=LET(r,B2,SUM(D5:((r))))",
        "=LET(r,B2,s,C3,SUM((r):(s)))",
        "=LET(r,B2,s,C3,SUM(r:(s)))",
        "=LET(r,B2,s,C3,SUM(((r)):((s))))",
        "=_xlfn.LET(_xlpm.r,B2,SUM((_xlpm.r):C3))",
        "=_xlfn.LET(_xlpm.r,B2,SUM((C3):_xlpm.r))",
        "=_xlfn.LET(_xlpm.r,B2,SUM(D5:((_xlpm.r))))",
        "=_xlfn.LET(_xlpm.r,B2,_xlpm.s,C3,SUM((_xlpm.r):(_xlpm.s)))",
    ),
)
def test_grouped_local_colons_remain_visible_and_conservative(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    vias = [reference.via for reference in result.references]
    assert vias.count("ref") == 2
    assert vias.count("opaque:ref") == 1
    assert result.issues == ()
    assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=LET(r,B2,SUM((r) C3))",
        "=LET(r,B2,SUM(((r)) C3))",
        "=LET(r,B2,SUM(C3 (r)))",
        "=LET(r,B2,SUM((C3) ((r))))",
        "=_xlfn.LET(_xlpm.r,B2,SUM(((_xlpm.r)) (C3)))",
        "=LAMBDA(r,SUM((r) (C3)))(B2)",
    ),
)
def test_grouped_local_intersections_remain_visible_and_conservative(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    vias = [reference.via for reference in result.references]
    assert vias.count("ref") == 2
    assert vias.count("opaque:ref") == 1
    assert result.issues == ()
    assert result.opaque


def test_grouped_local_intersections_reuse_one_matching_group_map(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original = analysis_module._matching_groups
    calls = 0

    def counted(tokens: tuple[FormulaToken, ...]) -> dict[int, int]:
        nonlocal calls
        calls += 1
        return original(tokens)

    monkeypatch.setattr(analysis_module, "_matching_groups", counted)
    formula = "=LET(r,A1,SUM(" + " ".join(["(r)"] * 190) + "))"

    result = analyze_formula(
        formula,
        anchor=_anchor(),
        context=ReferenceContext((_sheet("Data", 0),)),
    )

    assert calls <= 4
    assert [reference.via for reference in result.references].count("opaque:ref") == 189


@pytest.mark.parametrize(
    "formula",
    (
        "=LET(Pick,42,Pick C3)",
        "=LAMBDA(Pick,Pick C3)(42)",
        "=LET(Pick,42,(Pick) C3)",
        "=LAMBDA(Pick,(Pick) (C3))(42)",
    ),
)
def test_local_intersection_suppresses_same_spelled_formula_name_prepass(
    formula: str,
) -> None:
    context = ReferenceContext(
        (_sheet("Data", 0),),
        (DefinedName("Pick", "=INDEX($A:$A,1)", None, "formula", False),),
    )

    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert [reference.via for reference in result.references].count("opaque:ref") == 1
    assert "opaque:INDEX" not in {reference.via for reference in result.references}
    assert result.issues == ()


def test_stored_local_intersection_namespace_remains_exact() -> None:
    context = ReferenceContext(
        (_sheet("Data", 0),),
        (DefinedName("Pick", "=INDEX($A:$A,1)", None, "formula", False),),
    )
    local = analyze_formula(
        "=_xlfn.LET(_xlpm.Pick,42,_xlpm.Pick C3)",
        anchor=_anchor(),
        context=context,
    )
    global_name = analyze_formula(
        "=_xlfn.LET(_xlpm.Pick,42,Pick C3)",
        anchor=_anchor(),
        context=context,
    )

    assert [reference.via for reference in local.references].count("opaque:ref") == 1
    assert "opaque:INDEX" not in {reference.via for reference in local.references}
    assert "opaque:INDEX" in {reference.via for reference in global_name.references}
    assert _codes(global_name.issues) == ["I_DYNAMIC_REF"]


def test_stored_local_prefix_does_not_shadow_unprefixed_builtin_or_name(
    context: ReferenceContext,
) -> None:
    builtin = analyze_formula(
        "=_xlfn.LET(_xlpm.SUM,_xlfn.LAMBDA(_xlpm.x,_xlpm.x),SUM(_xlpm.SUM(A1):D5))",
        anchor=_anchor(),
        context=context,
    )
    global_name = analyze_formula(
        "=_xlfn.LET(_xlpm.Rate,42,Rate+_xlpm.Rate)",
        anchor=_anchor(),
        context=context,
    )

    assert builtin.function_calls == ("LET", "LAMBDA", "SUM", "SUM")
    assert builtin.returns_reference is False
    assert [reference.via for reference in builtin.references].count("opaque:SUM") == 1
    assert _codes(builtin.issues) == ["I_DYNAMIC_REF"]
    assert [(reference.token, reference.via) for reference in global_name.references] == [
        ("Rate", "name:Rate")
    ]
    assert global_name.issues == ()


def test_unprefixed_builtin_call_wins_over_same_spelled_raw_local() -> None:
    result = analyze_formula(
        "=SUM(LET(SUM,LAMBDA(x,x),SUM(A1)):D5)",
        anchor=_anchor(),
        context=ReferenceContext((_sheet("Data", 0),)),
    )

    assert result.function_calls == ("SUM", "LET", "LAMBDA", "SUM")
    assert not any(reference.via.startswith("opaque:") for reference in result.references)
    assert result.issues == ()
    assert not result.volatile


@pytest.mark.parametrize(
    "local_name",
    ("LET", "LAMBDA", "CHOOSE", "INDEX", "OFFSET", "INDIRECT", "IF", "SINGLE", "RAND"),
)
def test_stored_builtin_spelled_local_callable_uses_prefixed_namespace(
    local_name: str,
) -> None:
    stored = f"_xlpm.{local_name}"
    result = analyze_formula(
        f"=_xlfn.LET({stored},_xlfn.LAMBDA(_xlpm.x,_xlpm.x),SUM({stored}(A1):D5))",
        anchor=_anchor(),
        context=ReferenceContext((_sheet("Data", 0),)),
    )

    assert result.function_calls == ("LET", "LAMBDA", "SUM", local_name)
    assert f"opaque:{local_name}" in {reference.via for reference in result.references}
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert not result.volatile


@pytest.mark.parametrize(
    "formula",
    (
        "=SUM(LET(r,42,r:C3):D5)",
        "=SUM(LET(r,42,@r:C3):D5)",
        "=SUM(LET(r,42,r#:C3):D5)",
        "=SUM(LET(r,42,(r):C3):D5)",
        "=SUM(LET(r,42,D5:(r)):E5)",
        "=SUM(LET(r,42,(r):S):D5)",
        "=SUM(LET(r,42,s,43,(r):(s)):E5)",
        "=SUM(_xlfn.LET(_xlpm.r,42,_xlpm.r:C3):D5)",
    ),
)
def test_scalar_local_composites_do_not_gain_false_dynamic_reference_identity(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert "opaque:ref" in {reference.via for reference in result.references}
    assert "opaque:LET" not in {reference.via for reference in result.references}
    assert "I_DYNAMIC_REF" not in _codes(result.issues)
    assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=SUM(LET(r,B2,r:C3):D5)",
        "=SUM(LET(r,B2,@r:C3):D5)",
        "=SUM(LET(r,B2,r#:C3):D5)",
        "=SUM(LET(r,B2,(r):C3):D5)",
        "=SUM(LET(r,B2,D5:(r)):E5)",
        "=SUM(LET(r,B2,(r):S):D5)",
        "=SUM(LET(r,B2,s,C3,(r):(s)):E5)",
        "=SUM(_xlfn.LET(_xlpm.r,B2,_xlpm.r:C3):D5)",
    ),
)
def test_reference_local_composites_retain_outer_dynamic_attribution(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert "opaque:LET" in {reference.via for reference in result.references}
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


def test_lambda_defined_name_resolves_body_and_call_arguments(context: ReferenceContext) -> None:
    result = analyze_formula("=Double(A1)", anchor=_anchor(), context=context)

    assert [(reference.token, reference.via) for reference in result.references] == [
        ("$B$2", "ref"),
        ("A1", "ref"),
    ]
    assert result.function_calls == ("DOUBLE", "LAMBDA")
    assert result.issues == ()


@pytest.mark.parametrize(
    ("formula", "dynamic_via"),
    (
        ("=LAMBDA(r,SUM(INDEX(r,1):D5))(A1)", "opaque:INDEX"),
        ("=SUM(CHOOSE(1,LAMBDA(r,r)(A1),10):D5)", "opaque:CHOOSE"),
        ("=SUM(INDEX(LAMBDA(r,r)(A1),1):D5)", "opaque:INDEX"),
    ),
)
def test_inline_lambda_propagates_argument_reference_identity(
    context: ReferenceContext,
    formula: str,
    dynamic_via: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert dynamic_via in [reference.via for reference in result.references]
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=LAMBDA(r,SUM(INDEX(r,1):D5))(42)",
        "=SUM(CHOOSE(1,LAMBDA(r,r)(42),10):D5)",
        "=SUM(INDEX(LAMBDA(r,r)(42),1):D5)",
    ),
)
def test_inline_lambda_scalar_arguments_do_not_promote_reference_functions(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert not any(reference.via.startswith("opaque:") for reference in result.references)
    assert result.issues == ()


@pytest.mark.parametrize(
    "formula",
    (
        "=SUM(CHOOSE(1,LAMBDA(f,f)(LAMBDA(r,r))(A1),10):D5)",
        "=LET(make,LAMBDA(x,LAMBDA(y,x)),SUM(CHOOSE(1,make(A1)(42),10):D5))",
        "=SUM(CHOOSE(1,(LET(f,LAMBDA(r,r),f))(A1),10):D5)",
        "=LAMBDA(f,SUM(CHOOSE(1,f(A1),10):D5))(LAMBDA(r,r))",
    ),
)
def test_higher_order_lambdas_retain_callable_and_reference_results(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert [reference.via for reference in result.references].count("opaque:CHOOSE") == 1
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=SUM(CHOOSE(1,LAMBDA(f,f)(LAMBDA(r,42))(A1),10):D5)",
        "=LET(make,LAMBDA(x,LAMBDA(y,42)),SUM(CHOOSE(1,make(A1)(42),10):D5))",
        "=SUM(CHOOSE(1,(LET(f,LAMBDA(r,42),f))(A1),10):D5)",
    ),
)
def test_higher_order_lambda_scalar_results_do_not_promote_choose(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert "opaque:CHOOSE" not in [reference.via for reference in result.references]
    assert result.issues == ()
    assert not result.opaque


@pytest.mark.parametrize(
    "selector",
    (
        "CHOOSE(1,LAMBDA(r,r),LAMBDA(r,42))",
        "IF(TRUE,LAMBDA(r,r),LAMBDA(r,42))",
    ),
)
def test_callable_selector_results_retain_conservative_reference_flow(
    context: ReferenceContext,
    selector: str,
) -> None:
    result = analyze_formula(
        f"=LET(f,{selector},SUM(CHOOSE(1,f(A1),10):D5))",
        anchor=_anchor(),
        context=context,
    )

    assert [reference.via for reference in result.references].count("opaque:CHOOSE") == 1
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]


def test_nested_callable_choices_are_deduplicated_before_repeated_invocation(
    context: ReferenceContext,
) -> None:
    repeated_selector_calls = "(s)" * 14
    formula = (
        "=LET(s,IF(TRUE,LAMBDA(f,f),LAMBDA(f,f)),"
        f"SUM(CHOOSE(1,s{repeated_selector_calls}(A1),10):D5))"
    )

    started = perf_counter()
    result = analyze_formula(formula, anchor=_anchor(), context=context)
    elapsed = perf_counter() - started

    assert [reference.via for reference in result.references].count("opaque:CHOOSE") == 1
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque
    # The unbounded duplicate choice tree takes several seconds by depth 14;
    # normalized choices complete with ample headroom even under coverage.
    assert elapsed < 1.0


def test_callable_choice_cap_keeps_discarded_reference_branches_conservative(
    context: ReferenceContext,
) -> None:
    scalar_branches = ",".join(f"LAMBDA(x,{value})" for value in range(40))
    formula = f"=LET(f,CHOOSE(1,{scalar_branches},LAMBDA(x,x)),SUM(CHOOSE(1,f(A1),10):D5))"

    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert [reference.via for reference in result.references].count("opaque:CHOOSE") == 1
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


def test_computed_name_alias_chain_is_bounded_and_keeps_conservative_edges() -> None:
    sheet = _sheet("Data", 0)
    names: list[DefinedName] = [
        DefinedName(
            "Fixed",
            "'Data'!$A$2:$A$3",
            None,
            "range",
            False,
            (NameArea("Data", Rect(2, 3, 1, 1)),),
        )
    ]
    prior = "Fixed"
    for index in range(32):
        current = f"_F{index}"
        names.append(DefinedName(current, f"={prior}", None, "formula", False))
        prior = current
    context = ReferenceContext((sheet,), tuple(names))

    started = perf_counter()
    result = analyze_formula(
        f"=SUM({prior}:B5)",
        anchor=_anchor(col=3),
        context=context,
    )
    elapsed = perf_counter() - started

    assert elapsed < 1.0
    assert result.issues == ()
    assert result.opaque
    assert {reference.via for reference in result.references} == {
        "name:Fixed",
        "ref",
        "opaque:ref",
    }


def test_defined_lambda_propagates_typed_argument_reference_identity() -> None:
    sheet = _sheet("Data", 0)
    context = ReferenceContext(
        (sheet,),
        (
            DefinedName(
                "Pick",
                "=_xlfn.LAMBDA(r,r)",
                None,
                "lambda",
                False,
            ),
        ),
    )

    for formula, dynamic_via in (
        ("=SUM(CHOOSE(1,Pick(A1),10):D5)", "opaque:CHOOSE"),
        ("=SUM(INDEX(Pick(A1),1):D5)", "opaque:INDEX"),
    ):
        result = analyze_formula(formula, anchor=_anchor(), context=context)
        assert dynamic_via in [reference.via for reference in result.references]
        assert _codes(result.issues) == ["I_DYNAMIC_REF"]

    for formula in (
        "=SUM(CHOOSE(1,Pick(A1+0),10):D5)",
        "=SUM(INDEX(Pick(A1+0),1):D5)",
    ):
        result = analyze_formula(formula, anchor=_anchor(), context=context)
        assert not any(reference.via.startswith("opaque:") for reference in result.references)
        assert result.issues == ()


def test_defined_lambda_preserves_callable_arguments() -> None:
    sheet = _sheet("Data", 0)
    context = ReferenceContext(
        (sheet,),
        (
            DefinedName(
                "Apply",
                "=_xlfn.LAMBDA(f,x,f(x))",
                None,
                "lambda",
                False,
            ),
            DefinedName(
                "Pick",
                "=_xlfn.LAMBDA(r,r)",
                None,
                "lambda",
                False,
            ),
            DefinedName(
                "Make",
                "=_xlfn.LAMBDA(x,LAMBDA(y,x))",
                None,
                "lambda",
                False,
            ),
        ),
    )

    for formula, dynamic_via in (
        (
            "=SUM(CHOOSE(1,Apply(LAMBDA(r,r),A1),10):D5)",
            "opaque:CHOOSE",
        ),
        (
            "=SUM(INDEX(Apply(LAMBDA(r,r),A1),1):D5)",
            "opaque:INDEX",
        ),
        (
            "=SUM(CHOOSE(1,Apply(Pick,A1),10):D5)",
            "opaque:CHOOSE",
        ),
        (
            "=SUM(CHOOSE(1,Make(A1)(42),10):D5)",
            "opaque:CHOOSE",
        ),
        (
            "=LET(f,Pick,SUM(CHOOSE(1,Apply(f,A1),10):D5))",
            "opaque:CHOOSE",
        ),
    ):
        result = analyze_formula(formula, anchor=_anchor(), context=context)
        assert [reference.via for reference in result.references].count(dynamic_via) == 1
        assert _codes(result.issues) == ["I_DYNAMIC_REF"]

    scalar = analyze_formula(
        "=SUM(CHOOSE(1,Apply(LAMBDA(r,r),A1+0),10):D5)",
        anchor=_anchor(),
        context=context,
    )
    assert not any(reference.via.startswith("opaque:") for reference in scalar.references)
    assert scalar.issues == ()


def test_prefixed_lambda_defined_name_uses_the_stripped_callable_identifier(
    context: ReferenceContext,
) -> None:
    result = analyze_formula("=_xlfn.Double(A1)", anchor=_anchor(), context=context)

    assert [(reference.token, reference.via) for reference in result.references] == [
        ("$B$2", "ref"),
        ("A1", "ref"),
    ]
    assert result.function_calls == ("DOUBLE", "LAMBDA")
    assert result.issues == ()


def test_dynamic_functions_retain_explicit_refs_and_emit_opaque_edges(
    context: ReferenceContext,
) -> None:
    result = analyze_formula(
        '=INDIRECT(A1)+OFFSET(B2,1,0)+CHOOSE(1,C1,D1)+CELL("address",E1)',
        anchor=_anchor(),
        context=context,
    )

    assert result.volatile and result.opaque
    assert {reference.via for reference in result.references} == {
        "ref",
        "opaque:INDIRECT",
        "opaque:OFFSET",
        "opaque:CHOOSE",
    }
    assert _codes(result.issues) == ["I_DYNAMIC_REF", "I_DYNAMIC_REF", "I_DYNAMIC_REF"]
    assert "I_VOLATILE" not in _codes(result.issues)


def test_choose_and_index_are_only_opaque_in_reference_context(
    context: ReferenceContext,
) -> None:
    scalar_choose = analyze_formula("=CHOOSE(A1,10,20)", anchor=_anchor(), context=context)
    constant_choose = analyze_formula(
        "=CHOOSE(1,PiValue,SUM(A1)+1)",
        anchor=_anchor(),
        context=context,
    )
    scalar_index = analyze_formula("=INDEX(A:A,1)+1", anchor=_anchor(), context=context)
    range_index = analyze_formula(
        "=SUM(INDEX(A:A,1):INDEX(A:A,5))",
        anchor=_anchor(),
        context=context,
    )

    assert not scalar_choose.opaque
    assert not constant_choose.opaque
    assert not scalar_index.opaque
    assert [reference.via for reference in range_index.references].count("opaque:INDEX") == 2
    assert range_index.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=SUM((INDEX(A:A,1)):A5)",
        "=SUM(((INDEX(A:A,1))):A5)",
        "=SUM(A1:(INDEX(A:A,5)))",
        "=SUM(A1:((INDEX(A:A,5))))",
        "=SUM((A1):(INDEX(A:A,5)))",
    ),
)
def test_parenthesized_index_range_endpoints_remain_dynamic(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert [reference.via for reference in result.references].count("opaque:INDEX") == 1
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


def test_grouped_range_endpoints_do_not_create_empty_callable_warnings(
    context: ReferenceContext,
) -> None:
    grouped = analyze_formula("=SUM(A1:(A5))", anchor=_anchor(), context=context)
    grouped_both = analyze_formula("=SUM((A1):(A5))", anchor=_anchor(), context=context)
    arithmetic = analyze_formula(
        "=(INDEX(A:A,1)+1):A5",
        anchor=_anchor(),
        context=context,
    )

    for result in (grouped, grouped_both):
        assert [reference.token for reference in result.references] == ["A1:A5"]
        geometry = result.references[0].geometry
        assert geometry is not None
        assert resolve_reference(geometry, _anchor().cell) == Rect(1, 5, 1, 1)
        assert result.issues == ()
        assert not result.opaque
    assert "opaque:INDEX" not in [reference.via for reference in arithmetic.references]
    assert arithmetic.issues == ()
    assert not arithmetic.opaque


def test_static_reference_intersection_is_indexed_as_the_exact_overlap(
    context: ReferenceContext,
) -> None:
    for formula, expected in (
        ("=SUM(A1:C3 B2:D4)", Rect(2, 3, 2, 3)),
        ("=SUM((A1:C3) (B2:D4))", Rect(2, 3, 2, 3)),
        ("=SUM(A:A 5:5)", Rect(5, 5, 1, 1)),
    ):
        result = analyze_formula(
            formula,
            anchor=_anchor(),
            context=context,
        )

        assert result.issues == ()
        assert not result.opaque
        assert len(result.references) == 1
        geometry = result.references[0].geometry
        assert geometry is not None
        assert resolve_reference(geometry, _anchor().cell) == expected


def test_unrepresentable_static_intersection_is_visibly_conservative(
    context: ReferenceContext,
) -> None:
    result = analyze_formula(
        "=SUM($A$1:$C$3 B2:D4)",
        anchor=_anchor(),
        context=context,
    )

    assert "opaque:ref" in [reference.via for reference in result.references]
    assert result.issues == ()
    assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=(1+INDEX(A:A,1)):A5",
        "=(A1+INDEX(A:A,1)):A5",
        "=(INDEX(A:A,1)+INDEX(B:B,1)):A5",
        "=A1:(INDEX(A:A,1)+1)",
        "=A1:((INDEX(A:A,1))+1)",
    ),
)
def test_scalar_grouped_index_expressions_are_not_reference_endpoints(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert "opaque:INDEX" not in [reference.via for reference in result.references]
    assert result.issues == ()
    assert not result.opaque


def test_choose_detects_pure_nested_reference_returning_branches(
    context: ReferenceContext,
) -> None:
    nested = analyze_formula(
        "=SUM(CHOOSE(1,INDEX(A:A,1),INDEX(A:A,2)))",
        anchor=_anchor(),
        context=context,
    )
    scalar = analyze_formula(
        "=CHOOSE(1,INDEX(A:A,1)+1,10)",
        anchor=_anchor(),
        context=context,
    )
    array_source = analyze_formula(
        "=CHOOSE(1,INDEX({1,2},1),10)",
        anchor=_anchor(),
        context=context,
    )
    calculated_source = analyze_formula(
        "=CHOOSE(1,INDEX(A:A+0,1),10)",
        anchor=_anchor(),
        context=context,
    )
    intersected = analyze_formula(
        "=CHOOSE(1,@INDEX(A:A,1),10)",
        anchor=_anchor(),
        context=context,
    )
    offset_branch = analyze_formula(
        "=CHOOSE(1,OFFSET(A1,1,0),10)",
        anchor=_anchor(),
        context=context,
    )
    indirect_branch = analyze_formula(
        '=CHOOSE(1,INDIRECT("A1"),10)',
        anchor=_anchor(),
        context=context,
    )

    assert [reference.via for reference in nested.references].count("opaque:CHOOSE") == 1
    assert _codes(nested.issues) == ["I_DYNAMIC_REF"]
    assert nested.opaque
    for result in (scalar, array_source, calculated_source):
        assert "opaque:CHOOSE" not in [reference.via for reference in result.references]
        assert result.issues == ()
        assert not result.opaque
    assert "opaque:CHOOSE" in [reference.via for reference in intersected.references]
    assert _codes(intersected.issues) == ["I_DYNAMIC_REF"]
    assert intersected.opaque
    for result, inner_via in (
        (offset_branch, "opaque:OFFSET"),
        (indirect_branch, "opaque:INDIRECT"),
    ):
        vias = [reference.via for reference in result.references]
        assert inner_via in vias
        assert "opaque:CHOOSE" in vias
        assert _codes(result.issues) == ["I_DYNAMIC_REF", "I_DYNAMIC_REF"]
        assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=SUM(INDEX((A1:A3,C1:C3),1,1,2):B5)",
        "=SUM(INDEX(A1:C3 B2:D4,1,1):E5)",
        "=SUM(INDEX(A1 : B3,1,1):C5)",
        "=SUM(INDEX((A1):(B3),1,1):C5)",
        "=SUM(INDEX(A1:((B3)),1,1):C5)",
    ),
)
def test_index_reference_form_accepts_pure_union_and_intersection_sources(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert [reference.via for reference in result.references].count("opaque:INDEX") == 1
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


@pytest.mark.parametrize(
    ("formula", "expected_vias"),
    (
        (
            "=SUM(INDEX(INDEX(A1:C3,0,1),1):B5)",
            {"opaque:INDEX"},
        ),
        (
            "=SUM(INDEX(OFFSET(A1,1,0),1):B5)",
            {"opaque:INDEX", "opaque:OFFSET"},
        ),
        (
            "=SUM(INDEX(CHOOSE(1,A1:A3,C1:C3),1):B5)",
            {"opaque:INDEX", "opaque:CHOOSE"},
        ),
    ),
)
def test_index_reference_form_propagates_reference_returning_function_sources(
    context: ReferenceContext,
    formula: str,
    expected_vias: set[str],
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    vias = {reference.via for reference in result.references}
    assert expected_vias <= vias
    assert _codes(result.issues) == ["I_DYNAMIC_REF"] * len(expected_vias)
    assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=SUM(INDEX(CHOOSE(1,10,20),1):B5)",
        "=SUM(INDEX(SUM(A1:A3),1):B5)",
    ),
)
def test_index_reference_form_rejects_scalar_function_sources(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert "opaque:INDEX" not in [reference.via for reference in result.references]
    assert "I_DYNAMIC_REF" not in _codes(result.issues)


def test_index_reference_form_rejects_scalar_absorbed_range_prefix(
    context: ReferenceContext,
) -> None:
    result = analyze_formula(
        "=SUM(INDEX(PiValue:INDEX(B:B,1),1):$Z$1)",
        anchor=_anchor(),
        context=context,
    )

    assert [reference.via for reference in result.references].count("opaque:INDEX") == 1
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]


@pytest.mark.parametrize(
    "formula",
    (
        "=CHOOSE(1,(A1:A3,C1:C3),10)",
        "=CHOOSE(1,A1:C3 B2:D4,10)",
    ),
)
def test_choose_recognizes_pure_reference_union_and_intersection_branches(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert [reference.via for reference in result.references].count("opaque:CHOOSE") == 1
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=SUM(INDEX((A:A,,B:B),1):C5)",
        "=SUM(INDEX((A:A,B:B,),1):C5)",
        "=CHOOSE(1,(A:A,,B:B),10)",
        "=CHOOSE(1,(A:A,B:B,),10)",
    ),
)
def test_malformed_reference_unions_do_not_emit_dynamic_edges(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert not any(reference.via.startswith("opaque:") for reference in result.references)
    assert "I_DYNAMIC_REF" not in _codes(result.issues)


def test_choose_union_and_intersection_reject_scalar_name_operands(
    context: ReferenceContext,
) -> None:
    for formula in (
        "=CHOOSE(1,(A1,PiValue),10)",
        "=CHOOSE(1,A1 PiValue,10)",
        "=CHOOSE(1,(A1,AbsoluteFormula),10)",
        "=CHOOSE(1,A1 AbsoluteFormula,10)",
        "=LET(x,42,CHOOSE(1,(A1 x),10))",
        "=LAMBDA(x,CHOOSE(1,(A1 x),10))(42)",
    ):
        result = analyze_formula(formula, anchor=_anchor(), context=context)

        assert "opaque:CHOOSE" not in [reference.via for reference in result.references]
        assert "I_DYNAMIC_REF" not in _codes(result.issues)


@pytest.mark.parametrize(
    ("formula", "dynamic_via"),
    (
        ("=CHOOSE(1,@A1,10)", "opaque:CHOOSE"),
        ("=CHOOSE(1,@INDEX(B:B,5),10)", "opaque:CHOOSE"),
        ("=SUM(A1:(@INDEX(B:B,5)))", "opaque:INDEX"),
    ),
)
def test_implicit_intersection_preserves_single_cell_reference_result_type(
    context: ReferenceContext,
    formula: str,
    dynamic_via: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert dynamic_via in [reference.via for reference in result.references]
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=CHOOSE(1,@A1+0,10)",
        "=CHOOSE(1,@INDEX(B:B,5)+0,10)",
    ),
)
def test_value_operators_coerce_implicit_intersections_to_scalars(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert "opaque:CHOOSE" not in [reference.via for reference in result.references]
    assert result.issues == ()


@pytest.mark.parametrize(
    ("formula", "dynamic_vias"),
    (
        (
            "=SUM(CHOOSE(1,_xlfn.SINGLE(A1),10):D5)",
            ["opaque:CHOOSE"],
        ),
        (
            "=SUM(A1:(_xlfn.SINGLE(INDEX(B:B,5))))",
            ["opaque:INDEX"],
        ),
        (
            "=CHOOSE(1,_xlfn.SINGLE(INDEX(B:B,5)),10)",
            ["opaque:CHOOSE"],
        ),
    ),
)
def test_single_compatibility_wrapper_preserves_reference_identity(
    context: ReferenceContext,
    formula: str,
    dynamic_vias: list[str],
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    actual = [
        reference.via for reference in result.references if reference.via.startswith("opaque:")
    ]
    assert actual == dynamic_vias
    assert _codes(result.issues) == ["I_DYNAMIC_REF"] * len(dynamic_vias)


def test_single_wrapper_with_value_expression_is_scalar(
    context: ReferenceContext,
) -> None:
    result = analyze_formula(
        "=CHOOSE(1,_xlfn.SINGLE(INDEX(B:B,5)+0),10)",
        anchor=_anchor(),
        context=context,
    )

    assert not any(reference.via.startswith("opaque:") for reference in result.references)
    assert result.issues == ()


@pytest.mark.parametrize(
    ("formula", "dynamic_via"),
    (
        ("=SUM(_xlfn.SINGLE(A1):D5)", "opaque:SINGLE"),
        ("=SUM(LET(r,A1,r):D5)", "opaque:LET"),
        ("=SUM(LET(f,LAMBDA(r,r),f)(A1):D5)", "opaque:LET"),
        ("=SUM((LET(f,LAMBDA(r,r),f))(A1):D5)", "opaque:LET"),
        ("=SUM(LAMBDA(r,r)(A1):D5)", "opaque:LAMBDA"),
        ("=SUM((LAMBDA(r,r))(A1):D5)", "opaque:LAMBDA"),
    ),
)
def test_computed_reference_endpoints_are_conservatively_opaque(
    context: ReferenceContext,
    formula: str,
    dynamic_via: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert [reference.via for reference in result.references].count(dynamic_via) == 1
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


@pytest.mark.parametrize(
    "formula",
    (
        "=SUM(LET(r,INDEX(B:B,5),r):D5)",
        "=SUM(LAMBDA(r,INDEX(r,5))(B:B):D5)",
        "=SUM(_xlfn.SINGLE(LET(r,INDEX(B:B,5),r)):D5)",
    ),
)
def test_index_provenance_survives_reference_preserving_wrappers(
    context: ReferenceContext,
    formula: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    opaque = [
        reference.via for reference in result.references if reference.via.startswith("opaque:")
    ]
    assert opaque == ["opaque:INDEX"]
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]


@pytest.mark.parametrize(
    ("formula", "dynamic_via"),
    (
        (
            "=SUM(A1:(LAMBDA(r,r))(INDEX(B:B,5)))",
            "opaque:LAMBDA",
        ),
        (
            "=SUM(A1:(LET(f,LAMBDA(r,r),f))(INDEX(B:B,5)))",
            "opaque:LET",
        ),
    ),
)
def test_synthetic_left_range_group_connects_callable_invocation(
    context: ReferenceContext,
    formula: str,
    dynamic_via: str,
) -> None:
    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert [reference.via for reference in result.references].count(dynamic_via) == 1
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


@pytest.mark.parametrize(
    ("formula", "dynamic_via"),
    (
        ("=SUM(A1:(Pick)(D5))", "opaque:PICK"),
        ("=SUM(A1:((Pick))(D5))", "opaque:PICK"),
        ("=SUM(A1:((Pick)(D5)))", "opaque:PICK"),
        ("=SUM(A1:(((Pick)(D5))))", "opaque:PICK"),
        ("=SUM(A1:(Make)(D5)(42))", "opaque:MAKE"),
        ("=SUM(A1:((Make)(D5)(42)))", "opaque:MAKE"),
        ("=SUM((Pick)(A1):D5)", "opaque:PICK"),
        ("=SUM(((Pick)(A1)):D5)", "opaque:PICK"),
        ("=LET(f,LAMBDA(r,r),SUM(A1:(f)(D5)))", "opaque:F"),
        ("=LET(f,LAMBDA(r,r),SUM(A1:((f)(D5))))", "opaque:F"),
    ),
)
def test_grouped_named_and_lexical_callable_range_endpoints_are_opaque(
    formula: str,
    dynamic_via: str,
) -> None:
    context = ReferenceContext(
        (_sheet("Data", 0),),
        (
            DefinedName(
                "Pick",
                "=_xlfn.LAMBDA(r,r)",
                None,
                "lambda",
                False,
            ),
            DefinedName(
                "Make",
                "=_xlfn.LAMBDA(x,LAMBDA(y,x))",
                None,
                "lambda",
                False,
            ),
        ),
    )

    result = analyze_formula(formula, anchor=_anchor(), context=context)

    assert [reference.via for reference in result.references].count(dynamic_via) == 1
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


def test_choose_distinguishes_range_names_from_scalar_formula_names(
    context: ReferenceContext,
) -> None:
    range_name = analyze_formula(
        "=CHOOSE(1,Rate,10)",
        anchor=_anchor(),
        context=context,
    )
    scalar_formula = analyze_formula(
        "=CHOOSE(1,AbsoluteFormula,10)",
        anchor=_anchor(),
        context=context,
    )
    range_index = analyze_formula(
        "=SUM(INDEX(Rate,1):B5)",
        anchor=_anchor(),
        context=context,
    )
    scalar_index = analyze_formula(
        "=SUM(INDEX(AbsoluteFormula,1):B5)",
        anchor=_anchor(),
        context=context,
    )

    assert "opaque:CHOOSE" in [reference.via for reference in range_name.references]
    assert _codes(range_name.issues) == ["I_DYNAMIC_REF"]
    assert range_name.opaque
    assert "opaque:CHOOSE" not in [reference.via for reference in scalar_formula.references]
    assert scalar_formula.issues == ()
    assert not scalar_formula.opaque
    assert "opaque:INDEX" in [reference.via for reference in range_index.references]
    assert _codes(range_index.issues) == ["I_DYNAMIC_REF"]
    assert range_index.opaque
    assert "opaque:INDEX" not in [reference.via for reference in scalar_index.references]
    assert scalar_index.issues == ()
    assert not scalar_index.opaque


def test_both_parenthesized_index_range_endpoints_are_dynamic(
    context: ReferenceContext,
) -> None:
    result = analyze_formula(
        "=SUM((INDEX(A:A,1)):(INDEX(A:A,5)))",
        anchor=_anchor(),
        context=context,
    )

    assert [reference.via for reference in result.references].count("opaque:INDEX") == 2
    assert _codes(result.issues) == ["I_DYNAMIC_REF", "I_DYNAMIC_REF"]
    assert result.opaque


def test_index_over_an_external_reference_remains_dynamic(
    context: ReferenceContext,
) -> None:
    result = analyze_formula(
        "=SUM(INDEX([1]Data!A:A,1):B5)",
        anchor=_anchor(),
        context=context,
    )

    vias = [reference.via for reference in result.references]
    assert "opaque:INDEX" in vias
    assert "external:[budget.xlsx]" in vias
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


def test_volatile_modern_function_marks_flag_without_p3_diagnostic(
    context: ReferenceContext,
) -> None:
    result = analyze_formula("=_xlfn.RANDARRAY(2,2)+NOW()", anchor=_anchor(), context=context)

    assert result.volatile
    assert result.issues == ()
    assert result.function_calls == ("RANDARRAY", "NOW")


@pytest.mark.parametrize(
    "formula",
    (
        "=NOW()",
        "=TODAY()",
        "=RAND()",
        "=RANDBETWEEN(1,2)",
        "=RANDARRAY(1,1)",
        "=OFFSET(A1,0,0)",
        '=INDIRECT("A1")',
        '=CELL("address",A1)',
        '=INFO("directory")',
    ),
)
def test_each_frozen_volatile_function_sets_the_flag(
    context: ReferenceContext,
    formula: str,
) -> None:
    assert analyze_formula(formula, anchor=_anchor(), context=context).volatile


def test_computed_defined_names_remain_conservative_composite_endpoints() -> None:
    sheet = _sheet("Data", 0)
    context = ReferenceContext(
        (sheet,),
        (
            DefinedName("Pick", "=INDEX($A:$A,1)", None, "formula", False),
            DefinedName("PickLet", "=LET(r,INDEX($A:$A,1),r)", None, "formula", False),
            DefinedName("PickAlias", "=PickLet", None, "formula", False),
            DefinedName("Scalar", "=ABS($A$1)", None, "formula", False),
            DefinedName("Alias", "=Rate", None, "formula", False),
            DefinedName("BareLambda", "=LAMBDA(r,r)", None, "lambda", False),
            DefinedName("Pi", "=3.14", None, "constant", False),
            DefinedName(
                "Rate",
                "'Data'!$A$2",
                None,
                "range",
                False,
                (NameArea("Data", Rect(2, 2, 1, 1)),),
            ),
        ),
    )
    expected = {
        "Pick": ({Rect(1, 1_048_576, 1, 1), Rect(5, 5, 2, 2)}, "opaque:INDEX"),
        "PickLet": (
            {Rect(1, 1_048_576, 1, 1), Rect(5, 5, 2, 2)},
            "opaque:INDEX",
        ),
        "PickAlias": (
            {Rect(1, 1_048_576, 1, 1), Rect(5, 5, 2, 2)},
            "opaque:INDEX",
        ),
        "Scalar": ({Rect(1, 1, 1, 1), Rect(5, 5, 2, 2)}, "opaque:ref"),
        "Alias": ({Rect(2, 2, 1, 1), Rect(5, 5, 2, 2)}, "opaque:ref"),
        "BareLambda": ({Rect(5, 5, 2, 2)}, "opaque:ref"),
        "Pi": ({Rect(5, 5, 2, 2)}, "opaque:ref"),
    }

    for name, (expected_rectangles, marker) in expected.items():
        for formula in (
            f"=SUM({name}:B5)",
            f"=SUM(B5:{name})",
            f"=SUM({name} B5)",
            f"=SUM(({name}):B5)",
            f"=SUM((({name})) (B5))",
        ):
            result = analyze_formula(formula, anchor=_anchor(col=3), context=context)
            rectangles = {
                resolve_reference(reference.geometry, _anchor(col=3).cell)
                for reference in result.references
                if reference.geometry is not None
            }

            assert rectangles == expected_rectangles
            assert marker in {reference.via for reference in result.references}
            assert _codes(result.issues) == (["I_DYNAMIC_REF"] if marker == "opaque:INDEX" else [])
            assert result.opaque

    exact = analyze_formula("=SUM(Rate:B5)", anchor=_anchor(col=3), context=context)
    assert not exact.opaque
    assert {
        resolve_reference(reference.geometry, _anchor(col=3).cell)
        for reference in exact.references
        if reference.geometry is not None
    } == {Rect(2, 5, 1, 2)}


@pytest.mark.parametrize(
    ("formula", "marker"),
    (
        ("=SUM(INDEX(A:A,1) B5)", "opaque:INDEX"),
        ("=SUM(B5 INDEX(A:A,1))", "opaque:INDEX"),
        ("=SUM((INDEX(A:A,1)) B5)", "opaque:INDEX"),
        ("=SUM(Fn(1) B5)", "opaque:FN"),
        ("=SUM((Fn)(1) B5)", "opaque:FN"),
        ("=SUM(B5 (Fn)(1))", "opaque:FN"),
    ),
)
def test_computed_reference_intersections_emit_dynamic_attribution(
    formula: str,
    marker: str,
) -> None:
    sheet = _sheet("Data", 0)
    context = ReferenceContext(
        (sheet,),
        (
            DefinedName(
                "Fn",
                "=LAMBDA(r,INDEX(A:A,r))",
                None,
                "lambda",
                False,
            ),
        ),
    )

    result = analyze_formula(formula, anchor=_anchor(col=3), context=context)

    assert [reference.via for reference in result.references].count(marker) == 1
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.opaque


def test_every_unfolded_intersection_retains_its_own_conservative_attribution() -> None:
    sheet = _sheet("Data", 0)
    context = ReferenceContext(
        (sheet,),
        (
            DefinedName("Pick", "=INDEX($A:$A,1)", None, "formula", False),
            DefinedName("Pick2", "=INDEX($B:$B,2)", None, "formula", False),
        ),
    )

    mixed = analyze_formula(
        "=SUM($A$1:$C$3 B2:D4,Pick C5)",
        anchor=_anchor(col=3),
        context=context,
    )
    assert {reference.via for reference in mixed.references} >= {
        "opaque:ref",
        "opaque:INDEX",
    }
    assert _codes(mixed.issues) == ["I_DYNAMIC_REF"]
    assert mixed.opaque

    repeated = analyze_formula(
        "=SUM(Pick B5,Pick2 C5)",
        anchor=_anchor(col=3),
        context=context,
    )
    dynamic = [reference for reference in repeated.references if reference.via == "opaque:INDEX"]
    assert [reference.token for reference in dynamic] == ["Pick B5", "Pick2 C5"]
    assert _codes(repeated.issues) == ["I_DYNAMIC_REF", "I_DYNAMIC_REF"]
    assert repeated.opaque


def test_nested_computed_name_endpoint_deduplicates_dynamic_issue() -> None:
    sheet = _sheet("Data", 0)
    context = ReferenceContext(
        (sheet,),
        (
            DefinedName("Pick", "=INDEX($A:$A,1)", None, "formula", False),
            DefinedName("Both", "=LET(x,Pick B5,x)", None, "formula", False),
        ),
    )

    result = analyze_formula("=SUM(Both:C6)", anchor=_anchor(col=3), context=context)

    assert [reference.via for reference in result.references].count("opaque:INDEX") == 1
    assert _codes(result.issues) == ["I_DYNAMIC_REF"]
    assert result.issues[0].related["function"] == "INDEX"
    assert result.opaque


def test_name_expansion_cache_is_isolated_across_inherited_execution_contexts() -> None:
    sheet = _sheet("Data", 0)
    context = ReferenceContext(
        (sheet,),
        (DefinedName("Pick", "=INDEX($A:$A,1)", None, "formula", False),),
    )

    def cache():
        state = references_module._NAME_EXPANSION_CACHE.get()
        assert state is not None
        assert state.owner[0] is references_module.current_thread()
        return state.cache

    def expand():
        with references_module.name_expansion_scope():
            classify_ref("Pick", anchor=_anchor(), context=context)
            return cache()

    async def expand_in_task():
        return expand()

    async def run_tasks():
        return tuple(await asyncio.gather(*(expand_in_task() for _ in range(8))))

    with references_module.name_expansion_scope():
        parent_cache = cache()
        task_caches = asyncio.run(run_tasks())
        with ThreadPoolExecutor(max_workers=4) as pool:
            futures = [pool.submit(copy_context().run, expand) for _ in range(8)]
            thread_caches = tuple(future.result() for future in futures)

        assert not parent_cache
        assert all(child is not parent_cache for child in (*task_caches, *thread_caches))
        assert len({id(child) for child in task_caches}) == len(task_caches)
        assert len({id(child) for child in thread_caches}) == len(thread_caches)
        assert all(child for child in (*task_caches, *thread_caches))


def test_malformed_formulas_return_one_opaque_parse_edge(context: ReferenceContext) -> None:
    unclosed = analyze_formula("=SUM(", anchor=_anchor(), context=context)
    malformed_let = analyze_formula("=LET(x,A1)", anchor=_anchor(), context=context)
    missing_pairs = analyze_formula("=LET(x)", anchor=_anchor(), context=context)

    assert [reference.via for reference in unclosed.references] == ["opaque:parse"]
    assert _codes(unclosed.issues) == ["W_PARSE"]
    assert "opaque:parse" in [reference.via for reference in malformed_let.references]
    assert "W_PARSE" in _codes(malformed_let.issues)
    assert "W_PARSE" in _codes(missing_pairs.issues)


def test_recursive_defined_name_is_bounded_and_contained() -> None:
    sheet = _sheet("Data", 0)
    context = ReferenceContext(
        (sheet,),
        (DefinedName("Loop", "=Loop", None, "formula", False),),
    )

    result = analyze_formula("=Loop", anchor=_anchor(), context=context)

    assert result.opaque
    assert [reference.via for reference in result.references] == ["opaque:name"]
    assert _codes(result.issues) == ["W_PARSE"]


def test_tokenizer_normalizes_implicit_intersection_groups_without_empty_calls() -> None:
    grouped = tokenize_formula("=@('My Sheet'!A1:B5)")
    left_endpoint = tokenize_formula("=A1:@(B5)")
    both_endpoints = tokenize_formula("=SUM(@(A1):@(B5))")

    assert [(token.value, token.type, token.subtype) for token in grouped] == [
        ("@", "OPERATOR-PREFIX", ""),
        ("(", "PAREN", "OPEN"),
        ("'My Sheet'!A1:B5", "OPERAND", "RANGE"),
        (")", "PAREN", "CLOSE"),
    ]
    assert [(token.value, token.type, token.subtype) for token in left_endpoint] == [
        ("A1", "OPERAND", "RANGE"),
        (":", "OPERAND", "RANGE"),
        ("@", "OPERATOR-PREFIX", ""),
        ("(", "PAREN", "OPEN"),
        ("B5", "OPERAND", "RANGE"),
        (")", "PAREN", "CLOSE"),
    ]
    assert [(token.value, token.type, token.subtype) for token in both_endpoints] == [
        ("SUM(", "FUNC", "OPEN"),
        ("@", "OPERATOR-PREFIX", ""),
        ("(", "PAREN", "OPEN"),
        ("A1", "OPERAND", "RANGE"),
        (")", "PAREN", "CLOSE"),
        (":", "OPERAND", "RANGE"),
        ("@", "OPERATOR-PREFIX", ""),
        ("(", "PAREN", "OPEN"),
        ("B5", "OPERAND", "RANGE"),
        (")", "PAREN", "CLOSE"),
        (")", "FUNC", "CLOSE"),
    ]
    assert "".join(token.value for token in grouped) == "@('My Sheet'!A1:B5)"
    assert "".join(token.value for token in left_endpoint) == "A1:@(B5)"
    assert "".join(token.value for token in both_endpoints) == "SUM(@(A1):@(B5))"


@pytest.mark.parametrize(
    "formula",
    (
        "=@'My Sheet'!A1:B5",
        "=@'O''Brien'!A1",
        "=@'[budget 2025.xlsx]Q1'!A1:B5",
    ),
)
def test_tokenizer_shields_implicit_intersection_before_quoted_qualifiers(
    formula: str,
) -> None:
    tokens = tokenize_formula(formula)

    assert "".join(token.value for token in tokens) == formula[1:]
    assert len(tokens) == 1
    assert tokens[0].type == "OPERAND"
    assert tokens[0].subtype == "RANGE"


def test_implicit_intersection_groups_are_references_not_empty_callables(
    context: ReferenceContext,
) -> None:
    grouped = analyze_formula("=@('My Sheet'!A1:B5)", anchor=_anchor(), context=context)
    choose = analyze_formula("=CHOOSE(1,@(A1:B5),10)", anchor=_anchor(), context=context)
    endpoint = analyze_formula("=SUM(A1:@(B5))", anchor=_anchor(), context=context)

    assert grouped.function_calls == ()
    assert grouped.issues == ()
    grouped_geometry = grouped.references[0].geometry
    assert grouped_geometry is not None
    assert resolve_reference(grouped_geometry, _anchor().cell) == Rect(1, 5, 1, 2)
    assert "opaque:CHOOSE" in [reference.via for reference in choose.references]
    assert _codes(choose.issues) == ["I_DYNAMIC_REF"]
    assert endpoint.issues == ()
    assert [
        resolve_reference(reference.geometry, _anchor().cell)
        for reference in endpoint.references
        if reference.geometry is not None
    ] == [Rect(1, 5, 1, 2)]


@pytest.mark.parametrize(
    ("formula", "expected", "expected_vias"),
    (
        (
            "=SUM(@(A1):@(B5))",
            Rect(1, 5, 1, 2),
            {"ref"},
        ),
        (
            "=SUM(@((A1)):@(((B5))))",
            Rect(1, 5, 1, 2),
            {"ref"},
        ),
        (
            "=SUM(@(Rate):@(SalesTable[@Qty]))",
            Rect(2, 3, 2, 2),
            {"name:Rate", "structured:SalesTable[Qty]"},
        ),
        (
            "=SUM(@((SalesTable[@Qty])):@(((Rate))))",
            Rect(2, 3, 2, 2),
            {"name:Rate", "structured:SalesTable[Qty]"},
        ),
        (
            "=SUM(@(A1):@(SalesTable[Qty]))",
            Rect(1, 5, 1, 2),
            {"structured:SalesTable[Qty]"},
        ),
    ),
)
def test_both_grouped_implicit_intersection_range_endpoints_are_exact(
    context: ReferenceContext,
    formula: str,
    expected: Rect,
    expected_vias: set[str],
) -> None:
    anchor = FormulaAnchor(0, "Data", 3, 5)
    result = analyze_formula(formula, anchor=anchor, context=context)

    assert result.issues == ()
    assert not result.opaque
    assert "" not in result.function_calls
    assert {reference.via for reference in result.references} == expected_vias
    assert {
        resolve_reference(reference.geometry, anchor.cell)
        for reference in result.references
        if reference.geometry is not None
    } == {expected}


@pytest.mark.parametrize(
    ("formula", "anchor", "expected"),
    (
        (
            "=SUM(A1:@(SalesTable[@Qty]))",
            FormulaAnchor(0, "Data", 3, 5),
            Rect(1, 3, 1, 2),
        ),
        (
            "=SUM(A1:@([@Qty]))",
            FormulaAnchor(0, "Data", 3, 3),
            Rect(1, 3, 1, 2),
        ),
        (
            "=SUM(A1:SalesTable[Qty])",
            FormulaAnchor(0, "Data", 3, 5),
            Rect(1, 5, 1, 2),
        ),
    ),
)
def test_structured_range_endpoints_synthesize_exact_compatible_bounds(
    context: ReferenceContext,
    formula: str,
    anchor: FormulaAnchor,
    expected: Rect,
) -> None:
    result = analyze_formula(formula, anchor=anchor, context=context)

    assert result.issues == ()
    assert not result.opaque
    assert len(result.references) == 1
    geometry = result.references[0].geometry
    assert geometry is not None
    assert resolve_reference(geometry, anchor.cell) == expected


def test_grouped_whole_column_range_endpoints_keep_exact_bounds(
    context: ReferenceContext,
) -> None:
    direct = analyze_formula(
        "=SUM(A1:(B:B))",
        anchor=_anchor(),
        context=context,
    )
    structured = analyze_formula(
        "=SUM((B:B):(SalesTable[@Qty]))",
        anchor=FormulaAnchor(0, "Data", 3, 5),
        context=context,
    )

    assert direct.issues == structured.issues == ()
    assert not direct.opaque
    assert not structured.opaque
    assert len(direct.references) == len(structured.references) == 1
    direct_geometry = direct.references[0].geometry
    structured_geometry = structured.references[0].geometry
    assert direct_geometry is not None
    assert structured_geometry is not None
    assert resolve_reference(direct_geometry, _anchor().cell) == Rect(
        1,
        1_048_576,
        1,
        2,
    )
    assert resolve_reference(
        structured_geometry,
        FormulaAnchor(0, "Data", 3, 5).cell,
    ) == Rect(1, 1_048_576, 2, 2)


@pytest.mark.parametrize(
    ("formula", "expected", "structured_via"),
    (
        (
            "=SUM(Rate:SalesTable[@Qty])",
            Rect(2, 3, 1, 2),
            "structured:SalesTable[Qty]",
        ),
        (
            "=SUM(Rate:SalesTable[Qty])",
            Rect(2, 5, 1, 2),
            "structured:SalesTable[Qty]",
        ),
    ),
)
def test_name_and_structured_range_endpoints_share_exact_compatible_bounds(
    context: ReferenceContext,
    formula: str,
    expected: Rect,
    structured_via: str,
) -> None:
    exact_context = ReferenceContext(
        context.sheets,
        (
            DefinedName(
                "Rate",
                "'Data'!$A$2:$A$3",
                None,
                "range",
                False,
                (NameArea("Data", Rect(2, 3, 1, 1)),),
            ),
        ),
        context.tables,
        context.external_links,
    )
    anchor = FormulaAnchor(0, "Data", 3, 5)

    result = analyze_formula(formula, anchor=anchor, context=exact_context)

    assert result.issues == ()
    assert not result.opaque
    assert {reference.via for reference in result.references} == {
        "name:Rate",
        structured_via,
    }
    assert {
        resolve_reference(reference.geometry, anchor.cell)
        for reference in result.references
        if reference.geometry is not None
    } == {expected}


def test_mixed_name_a1_ranges_synthesize_excel_single_area_bounds(
    context: ReferenceContext,
) -> None:
    exact_context = ReferenceContext(
        context.sheets,
        (
            DefinedName(
                "Rate",
                "'Data'!$A$2:$A$3",
                None,
                "range",
                False,
                (NameArea("Data", Rect(2, 3, 1, 1)),),
            ),
            DefinedName(
                "Block",
                "'Data'!$A$2:$B$3",
                None,
                "range",
                False,
                (NameArea("Data", Rect(2, 3, 1, 2)),),
            ),
            DefinedName(
                "Start",
                "'Data'!$A$1",
                None,
                "range",
                False,
                (NameArea("Data", Rect(1, 1, 1, 1)),),
            ),
            DefinedName(
                "End",
                "'Data'!$D$5",
                None,
                "range",
                False,
                (NameArea("Data", Rect(5, 5, 4, 4)),),
            ),
        ),
    )
    expected = {
        "Rate:B5": Rect(2, 5, 1, 2),
        "A1:Rate": Rect(1, 3, 1, 1),
        "Block:D5": Rect(2, 5, 1, 4),
        "D1:Block": Rect(1, 3, 1, 4),
    }

    for text, bounds in expected.items():
        result = classify_ref(text, anchor=_anchor(), context=exact_context)

        assert result.issues == ()
        assert not result.opaque
        assert len(result.references) == 1
        assert result.references[0].via in {"name:Rate", "name:Block"}
        geometry = result.references[0].geometry
        assert geometry is not None
        assert resolve_reference(geometry, _anchor().cell) == bounds

    direct_names = classify_ref("Start:End", anchor=_anchor(), context=exact_context)
    assert direct_names.issues == ()
    assert not direct_names.opaque
    assert [reference.via for reference in direct_names.references] == [
        "name:Start",
        "name:End",
    ]
    direct_bounds: set[Rect] = set()
    for reference in direct_names.references:
        geometry = reference.geometry
        assert geometry is not None
        direct_bounds.add(resolve_reference(geometry, _anchor().cell))
    assert direct_bounds == {Rect(1, 5, 1, 4)}

    whole_column = classify_ref("Start:B:B", anchor=_anchor(), context=exact_context)
    whole_row = classify_ref("Start:5:5", anchor=_anchor(), context=exact_context)
    reverse_column = classify_ref("B:B:Start", anchor=_anchor(), context=exact_context)
    ambiguous_name = classify_ref("A:A:End", anchor=_anchor(), context=exact_context)
    reverse_ambiguous = classify_ref("End:A:A", anchor=_anchor(), context=exact_context)
    assert (
        whole_column.issues
        == whole_row.issues
        == reverse_column.issues
        == ambiguous_name.issues
        == reverse_ambiguous.issues
        == ()
    )
    assert not any(
        result.opaque
        for result in (
            whole_column,
            whole_row,
            reverse_column,
            ambiguous_name,
            reverse_ambiguous,
        )
    )
    for result, bounds in (
        (whole_column, Rect(1, 1_048_576, 1, 2)),
        (whole_row, Rect(1, 5, 1, 16_384)),
        (reverse_column, Rect(1, 1_048_576, 1, 2)),
        (ambiguous_name, Rect(1, 1_048_576, 1, 4)),
        (reverse_ambiguous, Rect(1, 1_048_576, 1, 4)),
    ):
        assert len(result.references) == 1
        geometry = result.references[0].geometry
        assert geometry is not None
        assert resolve_reference(geometry, _anchor().cell) == bounds


def test_parenthesized_name_range_endpoints_keep_exact_range_semantics(
    context: ReferenceContext,
) -> None:
    exact_context = ReferenceContext(
        context.sheets,
        (
            DefinedName(
                "Start",
                "'Data'!$A$1",
                None,
                "range",
                False,
                (NameArea("Data", Rect(1, 1, 1, 1)),),
            ),
            DefinedName(
                "End",
                "'Data'!$D$5",
                None,
                "range",
                False,
                (NameArea("Data", Rect(5, 5, 4, 4)),),
            ),
        ),
    )

    for formula in ("=Start:End", "=Start:(End)", "=(Start):End"):
        result = analyze_formula(formula, anchor=_anchor(), context=exact_context)

        assert result.issues == ()
        assert not result.opaque
        assert {
            resolve_reference(reference.geometry, _anchor().cell)
            for reference in result.references
            if reference.geometry is not None
        } == {Rect(1, 5, 1, 4)}


def test_mixed_name_a1_spills_keep_endpoint_semantics(
    context: ReferenceContext,
) -> None:
    left = classify_ref("rAtE#:B5", anchor=_anchor(), context=context)
    right = classify_ref("A1:rAtE#", anchor=_anchor(), context=context)
    a1_left = classify_ref("A1#:rAtE", anchor=_anchor(), context=context)
    a1_right = classify_ref("rAtE:A1#", anchor=_anchor(), context=context)

    assert left.issues == right.issues == a1_left.issues == a1_right.issues == ()
    assert not any(result.opaque for result in (left, right, a1_left, a1_right))
    assert [(reference.token, reference.via) for reference in left.references] == [
        ("rAtE#:B5", "spill"),
        ("rAtE#:B5", "ref"),
    ]
    assert [(reference.token, reference.via) for reference in right.references] == [
        ("A1:rAtE#", "ref"),
        ("A1:rAtE#", "spill"),
    ]
    assert [reference.via for reference in a1_left.references] == ["spill", "name:Rate"]
    assert [reference.via for reference in a1_right.references] == ["name:Rate", "spill"]


def test_mixed_ranges_contain_multi_area_and_cross_sheet_uncertainty(
    context: ReferenceContext,
) -> None:
    multi = classify_ref("Buckets:B5", anchor=_anchor(), context=context)
    cross_sheet = classify_ref("'My Sheet'!A1:Rate", anchor=_anchor(), context=context)
    qualified_right = classify_ref("Rate:'My Sheet'!B5", anchor=_anchor(), context=context)

    assert multi.issues == cross_sheet.issues == qualified_right.issues == ()
    assert multi.opaque and cross_sheet.opaque and qualified_right.opaque
    assert len(multi.references) == 3
    assert {reference.dst_sheet_name for reference in cross_sheet.references} == {
        "Data",
        "My Sheet",
    }
    assert {reference.dst_sheet_name for reference in qualified_right.references} == {
        "Data",
        "My Sheet",
    }


def test_outer_implicit_intersection_dispatches_bare_table_and_external_refs(
    context: ReferenceContext,
) -> None:
    bare_column = classify_ref("@[Qty]", anchor=_anchor(), context=context)
    current_row = classify_ref("[@Qty]", anchor=_anchor(), context=context)
    external = classify_ref("@[1]Data!A1", anchor=_anchor(), context=context)
    quoted = classify_ref("@'My Sheet'!A1:B5", anchor=_anchor(), context=context)

    assert bare_column.issues == current_row.issues == external.issues == quoted.issues == ()
    bare_geometry = bare_column.references[0].geometry
    current_geometry = current_row.references[0].geometry
    quoted_geometry = quoted.references[0].geometry
    assert bare_geometry is not None
    assert current_geometry is not None
    assert quoted_geometry is not None
    assert resolve_reference(bare_geometry, _anchor().cell) == Rect(2, 5, 2, 2)
    assert resolve_reference(current_geometry, _anchor().cell) == Rect(3, 3, 2, 2)
    assert external.references[0].via == "external:[budget.xlsx]"
    assert external.references[0].geometry is None
    assert quoted.references[0].dst_sheet_name == "My Sheet"
    assert resolve_reference(quoted_geometry, _anchor().cell) == Rect(1, 5, 1, 2)
