"""Compact workbook-map ordering, content, and degradation tests."""

from __future__ import annotations

import json
from collections import Counter
from pathlib import Path
from typing import cast

import pytest

import excel_lsp.core.workbook_map as workbook_map_module
from excel_lsp.core.index import IndexStore
from excel_lsp.core.models import IndexUpdate
from excel_lsp.core.workbook_map import build_workbook_map, serialize_workbook_map


def test_workbook_map_uses_stable_ids_caps_content_and_surfaces_visibility(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook, database = _seed_map_index(tmp_path)
    _replace_refresh(monkeypatch, workbook, database)

    result = build_workbook_map(workbook)
    serialized = serialize_workbook_map(result)

    assert len(serialized) <= 8_000
    assert "DO_NOT_LEAK_BODY_VALUE" not in serialized
    assert result["workbook"] == "model.xlsx"
    assert result["sheets"] == 2
    assert result["hasVBA"] is True
    assert result["namesMore"] == 5
    assert result["externalLinksMore"] == 2
    assert cast(list[str], result["externalLinks"])[0] == "[Book01.xlsx]"
    assert result["diagCounts"] == {"error": 1, "warn": 2}

    sheet_list = cast(list[dict[str, object]], result["sheetList"])
    assert [sheet["sheet"] for sheet in sheet_list] == ["Visible", "Audit"]
    assert "vis" not in sheet_list[0]
    assert sheet_list[1]["vis"] == "veryHidden"
    visible_regions = cast(list[dict[str, object]], sheet_list[0]["regions"])
    assert [region.get("id") for region in visible_regions[:-1]] == [
        f"region:Visible:{n}" for n in range(11, 3, -1)
    ]
    assert visible_regions[-1] == {"more": 4}


def test_workbook_map_has_a_deterministic_generic_character_fallback(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook, database = _seed_map_index(tmp_path)
    _replace_refresh(monkeypatch, workbook, database)

    first = build_workbook_map(workbook, character_cap=700)
    second = build_workbook_map(workbook, character_cap=700)

    assert first == second
    assert len(serialize_workbook_map(first)) <= 700
    assert first["namesMore"] == 20
    assert len(cast(list[dict[str, str]], first["names"])) == 5
    sheet_list = cast(list[dict[str, object]], first["sheetList"])
    assert all(
        not cast(list[dict[str, object]], sheet["regions"])
        or set(cast(list[dict[str, object]], sheet["regions"])[0]) == {"more"}
        for sheet in sheet_list
    )


def test_character_degradation_truncates_wide_columns_with_exact_remainder(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook, database = _seed_wide_map_index(tmp_path)
    _replace_refresh(monkeypatch, workbook, database)

    first = build_workbook_map(workbook, character_cap=1_600)
    second = build_workbook_map(workbook, character_cap=1_600)
    serialized = serialize_workbook_map(first)

    assert first == second
    assert len(serialized) <= 1_600
    assert "sheetListMore" not in first
    sheet_list = cast(list[dict[str, object]], first["sheetList"])
    assert [sheet["sheet"] for sheet in sheet_list] == ["Wide"]
    regions = cast(list[dict[str, object]], sheet_list[0]["regions"])
    assert len(regions) == 1
    assert regions[0]["id"] == "region:Wide:0"
    columns = cast(list[dict[str, str]], regions[0]["cols"])
    assert [column["h"] for column in columns] == [
        f"Wide Header {index:02d} {'X' * 24}" for index in range(16)
    ]
    assert regions[0]["colsMore"] == 14


def test_character_degradation_omits_sheet_detail_but_prioritizes_hidden_sheets(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook, database = _seed_many_sheet_map_index(tmp_path)
    _replace_refresh(monkeypatch, workbook, database)

    first = build_workbook_map(workbook, character_cap=1_450)
    second = build_workbook_map(workbook, character_cap=1_450)
    serialized = serialize_workbook_map(first)

    assert first == second
    assert len(serialized) <= 1_450
    assert first["sheets"] == 50
    assert first["sheetListMore"] == 30
    sheet_list = cast(list[dict[str, object]], first["sheetList"])
    assert [sheet["sheet"] for sheet in sheet_list] == [
        *(f"Sheet{number:02d}" for number in range(1, 19)),
        "Sheet45",
        "Sheet50",
    ]
    assert all(sheet["regions"] == [] for sheet in sheet_list)
    assert "vis" not in sheet_list[0]
    assert sheet_list[-2]["vis"] == "hidden"
    assert sheet_list[-1]["vis"] == "veryHidden"
    assert first["names"] == []
    assert first["namesMore"] == 0
    assert first["externalLinks"] == []
    assert first["externalLinksMore"] == 0


def test_workbook_map_rejects_an_impossible_custom_cap(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook, database = _seed_map_index(tmp_path)
    _replace_refresh(monkeypatch, workbook, database)

    with pytest.raises(ValueError, match="too small"):
        build_workbook_map(workbook, character_cap=10)
    with pytest.raises(ValueError, match="cannot exceed 8000"):
        build_workbook_map(workbook, character_cap=8_001)


def test_external_link_labels_strip_url_secrets_and_unsupported_schemes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook, database = _seed_map_index(tmp_path)
    with IndexStore(database) as store:
        store.set_meta(
            "external_links",
            json.dumps(
                {
                    "1": "https://example.com/budget.xlsx?sig=SECRET_QUERY",
                    "2": "https://user:pass@example.com",
                    "3": (
                        "https://example.com/models/model.xlsx;SECRET_PARAM"
                        "?sig=SECRET_QUERY_TWO#SECRET_FRAGMENT"
                    ),
                    "4": ("https://example.com/encoded.xlsx%253Fsig%253DSECRET_DOUBLE_ENCODED"),
                    "5": "ftp://user:SECRET_FTP@example.com/private.xlsx",
                    "6": "file:///C:/models/local.xlsx?sig=SECRET_FILE",
                    "7": r"C:\Models\windows.xlsx",
                    "8": r"\\server\share\unc.xlsx",
                    "9": "../relative/relative.xlsx#SECRET_RELATIVE",
                    "10": "http://example.com/http.xlsx",
                },
                separators=(",", ":"),
            ),
        )
    _replace_refresh(monkeypatch, workbook, database)

    result = build_workbook_map(workbook)
    serialized = serialize_workbook_map(result)

    assert result["externalLinks"] == [
        "[budget.xlsx]",
        "[external-workbook]",
        "[model.xlsx]",
        "[encoded.xlsx]",
        "[external-workbook]",
        "[local.xlsx]",
        "[windows.xlsx]",
        "[unc.xlsx]",
        "[relative.xlsx]",
        "[http.xlsx]",
    ]
    assert result["externalLinksMore"] == 0
    for sentinel in (
        "SECRET_QUERY",
        "user:pass",
        "SECRET_PARAM",
        "SECRET_FRAGMENT",
        "SECRET_DOUBLE_ENCODED",
        "SECRET_FTP",
        "SECRET_FILE",
        "SECRET_RELATIVE",
    ):
        assert sentinel not in serialized

    malformed_targets = {
        "1": "https:SECRET_CREDENTIAL",
        "2": "https:/user:SECRET_MALFORMED@example.com",
        "3": "http:///SECRET_HOSTNAME",
        "4": "https%3A%2F%2Fuser%3ASECRET_ENCODED%40example.com",
        "5": "https%253A%252F%252Fuser%253ASECRET_DOUBLE%2540example.com",
        "6": "file:user:SECRET_FILE_URI@example.com",
        "7": "https ://user:SECRET_SPACED@example.com",
        "8": "example.com",
        "9": ("https://example.com/x.xlsx%2525253Fsig%2525253DSECRET_QUAD"),
    }
    with IndexStore(database) as store:
        store.set_meta("external_links", json.dumps(malformed_targets))
    malformed_result = build_workbook_map(workbook)
    malformed_serialized = serialize_workbook_map(malformed_result)
    assert malformed_result["externalLinks"] == [
        *(["[external-workbook]"] * 8),
        "[x.xlsx]",
    ]
    assert malformed_result["externalLinksMore"] == 0
    for sentinel in (
        "SECRET_CREDENTIAL",
        "SECRET_MALFORMED",
        "SECRET_HOSTNAME",
        "SECRET_ENCODED",
        "SECRET_DOUBLE",
        "SECRET_FILE_URI",
        "SECRET_SPACED",
        "example.com",
        "SECRET_QUAD",
    ):
        assert sentinel not in malformed_serialized

    with IndexStore(database) as store:
        store.set_meta(
            "external_links",
            json.dumps(
                {
                    "1": "budget.xlsx%26sig%3DSECRET_SUFFIX_QUERY.xlsx",
                    "2": "sig%3DSECRET_PREFIX%26budget.xlsx",
                    "3": "../api_key=SECRET_RAW+budget.xlsx",
                }
            ),
        )
    query_shaped_result = build_workbook_map(workbook)
    query_shaped_serialized = serialize_workbook_map(query_shaped_result)
    assert query_shaped_result["externalLinks"] == ["[external-workbook]"] * 3
    for sentinel in ("SECRET_SUFFIX_QUERY", "SECRET_PREFIX", "SECRET_RAW"):
        assert sentinel not in query_shaped_serialized

    with IndexStore(database) as store:
        store.set_meta(
            "external_links",
            '{"1":"[Already.xlsx]","2":"../data.csv","3":"../model.ods"}',
        )
    bracketed_result = build_workbook_map(workbook)
    assert bracketed_result["externalLinks"] == [
        "[Already.xlsx]",
        "[data.csv]",
        "[model.ods]",
    ]


@pytest.mark.parametrize(
    ("metadata", "expected_links", "expected_total", "sentinels"),
    [
        (
            (
                '{"1":{"password":"SECRET_OBJECT"},"2":null,"3":42,'
                '"alpha":"SECRET_KEY","01":"SECRET_LEADING_ZERO","4":"../safe.xlsx"}'
            ),
            ["[safe.xlsx]"],
            1,
            ("SECRET_OBJECT", "SECRET_KEY", "SECRET_LEADING_ZERO"),
        ),
        ('[42,"SECRET_ARRAY"]', [], 0, ("SECRET_ARRAY",)),
        ('"SECRET_SCALAR"', [], 0, ("SECRET_SCALAR",)),
        ("null", [], 0, ()),
        ('{"not-a-number":"SECRET_NONNUMERIC"}', [], 0, ("SECRET_NONNUMERIC",)),
        ("{SECRET_INVALID", [], 0, ("SECRET_INVALID",)),
    ],
)
def test_external_link_projection_ignores_invalid_json_shapes_and_values(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    metadata: str,
    expected_links: list[str],
    expected_total: int,
    sentinels: tuple[str, ...],
) -> None:
    workbook, database = _seed_map_index(tmp_path)
    with IndexStore(database) as store:
        store.set_meta("external_links", metadata)
    _replace_refresh(monkeypatch, workbook, database)

    result = build_workbook_map(workbook)
    serialized = serialize_workbook_map(result)

    assert result["externalLinks"] == expected_links
    assert (
        len(cast(list[object], result["externalLinks"])) + cast(int, result["externalLinksMore"])
        == expected_total
    )
    for sentinel in sentinels:
        assert sentinel not in serialized


def test_extreme_degradation_summarizes_omitted_visibility_exactly(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook, database = _seed_visibility_pressure_index(tmp_path)
    _replace_refresh(monkeypatch, workbook, database)

    first = build_workbook_map(workbook, character_cap=700)
    second = build_workbook_map(workbook, character_cap=700)
    serialized = serialize_workbook_map(first)

    assert first == second
    assert len(serialized) <= 700
    sheet_list = cast(list[dict[str, object]], first["sheetList"])
    rendered = Counter(str(sheet.get("vis", "visible")) for sheet in sheet_list)
    assert first["sheetListMore"] == 100 - len(sheet_list)
    assert first["sheetListMoreByVis"] == {
        visibility: total - rendered[visibility]
        for visibility, total in (
            ("visible", 20),
            ("hidden", 60),
            ("veryHidden", 20),
        )
        if total > rendered[visibility]
    }
    omitted_visibility = cast(dict[str, int], first["sheetListMoreByVis"])
    assert omitted_visibility["hidden"] > 0
    assert omitted_visibility["veryHidden"] == 20


def test_full_cap_preserves_41_empty_sheet_details_without_remainder(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook, database = _seed_41_sheet_map_index(tmp_path)
    _replace_refresh(monkeypatch, workbook, database)

    result = build_workbook_map(workbook)
    sheet_list = cast(list[dict[str, object]], result["sheetList"])

    assert len(serialize_workbook_map(result)) <= 8_000
    assert len(sheet_list) == 41
    assert [sheet["sheet"] for sheet in sheet_list] == [f"S{number:02d}" for number in range(1, 42)]
    assert "sheetListMore" not in result
    assert "sheetListMoreByVis" not in result


def test_full_cap_preserves_17_short_columns_without_remainder(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook, database = _seed_17_column_map_index(tmp_path)
    _replace_refresh(monkeypatch, workbook, database)

    result = build_workbook_map(workbook)
    sheet = cast(list[dict[str, object]], result["sheetList"])[0]
    region = cast(list[dict[str, object]], sheet["regions"])[0]

    assert len(serialize_workbook_map(result)) <= 8_000
    assert len(cast(list[dict[str, str]], region["cols"])) == 17
    assert "colsMore" not in region


def test_batched_formula_block_counts_match_region_intersections(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook, database = _seed_fblock_map_index(tmp_path)
    _replace_refresh(monkeypatch, workbook, database)

    result = build_workbook_map(workbook)
    sheets = cast(list[dict[str, object]], result["sheetList"])
    counts = {
        cast(str, region["id"]): cast(int, region["fblocks"])
        for sheet in sheets
        for region in cast(list[dict[str, object]], sheet["regions"])
        if "id" in region
    }

    assert counts == {
        "region:One:0": 2,
        "region:One:1": 2,
        "region:Two:0": 1,
    }


def test_source_loading_has_fixed_queries_and_bounds_at_100_plus_sheet_scale(
    tmp_path: Path,
) -> None:
    tiny_workbook, tiny_database = _seed_map_index(tmp_path)
    large_workbook, large_database = _seed_scaling_map_index(tmp_path)

    tiny_source, tiny_query_count = _load_source_with_query_count(
        tiny_workbook,
        tiny_database,
    )
    source, large_query_count = _load_source_with_query_count(
        large_workbook,
        large_database,
    )

    assert tiny_query_count == large_query_count == 7
    assert tiny_source.sheet_count == 2
    assert len(tiny_source.sheets) == 2
    assert tiny_source.name_count == 25
    assert len(tiny_source.names) == 20
    assert tiny_source.external_link_count == 12
    assert len(tiny_source.external_links) == 10

    assert source.sheet_count == 121
    assert source.name_count == 500
    assert source.external_link_count == 15
    assert len(source.sheets) == 121
    assert [sheet.name for sheet in source.sheets[-3:]] == [
        "Scale119",
        "Scale120",
        "Scale121",
    ]
    assert sum(len(sheet.regions) for sheet in source.sheets) == 80
    assert all(sheet.region_count == 12 for sheet in source.sheets)
    assert all(len(sheet.regions) == 1 for sheet in source.sheets[:80])
    assert all(not sheet.regions for sheet in source.sheets[80:])
    assert all(region.column_count == 40 for sheet in source.sheets for region in sheet.regions)
    loaded_column_counts = [
        len(region.columns) for sheet in source.sheets for region in sheet.regions
    ]
    assert all(16 <= count <= 40 for count in loaded_column_counts)
    assert sum(loaded_column_counts) == 80 * 16 + 512
    assert len(source.names) == 20
    assert len(source.external_links) == 10

    first = workbook_map_module._bounded_render(source, character_cap=700)
    second = workbook_map_module._bounded_render(source, character_cap=700)
    serialized = serialize_workbook_map(first)
    sheet_list = cast(list[dict[str, object]], first["sheetList"])

    assert first == second
    assert len(serialized) <= 700
    assert first["sheets"] == 121
    assert first["sheetListMore"] == 121 - len(sheet_list)
    assert first["namesMore"] == 500 - len(cast(list[object], first["names"]))
    assert first["externalLinksMore"] == 15 - len(cast(list[object], first["externalLinks"]))


def test_sheet_source_loading_is_bounded_above_the_cap_proof_limit(tmp_path: Path) -> None:
    workbook, database = _seed_over_source_sheet_limit_index(tmp_path)

    source, query_count = _load_source_with_query_count(workbook, database)

    assert query_count == 7
    assert source.sheet_count == 205
    assert source.visibility_counts == (("visible", 204), ("hidden", 0), ("veryHidden", 1))
    assert len(source.sheets) == 200
    assert [sheet.name for sheet in source.sheets[-2:]] == ["Bound199", "Bound205"]


def _load_source_with_query_count(
    workbook: Path,
    database: Path,
) -> tuple[workbook_map_module._MapSource, int]:
    statements: list[str] = []
    with IndexStore(database) as store:
        store.connection.set_trace_callback(
            lambda statement: (
                statements.append(statement)
                if statement.lstrip().upper().startswith(("SELECT", "WITH"))
                else None
            )
        )
        source = workbook_map_module._load_source(
            workbook,
            store,
        )
        store.connection.set_trace_callback(None)
    return source, len(statements)


def _replace_refresh(
    monkeypatch: pytest.MonkeyPatch,
    workbook: Path,
    database: Path,
) -> None:
    def refresh(path: str | Path, *, index_dir: str | Path | None = None) -> IndexUpdate:
        del index_dir
        assert Path(path).resolve() == workbook.resolve()
        return IndexUpdate(str(workbook), str(database), 7, False, ())

    monkeypatch.setattr(workbook_map_module, "ensure_fresh", refresh)


def _seed_map_index(tmp_path: Path) -> tuple[Path, Path]:
    workbook = tmp_path / "model.xlsx"
    workbook.write_bytes(b"map-test-placeholder")
    database = tmp_path / "model.xlsp.db"
    with IndexStore(database) as store:
        store.connection.executemany(
            """
            INSERT INTO sheets(
                id, name, xml_part, part_hash, kind, visibility, max_row, max_col
            ) VALUES (?, ?, ?, ?, 'worksheet', ?, ?, ?)
            """,
            (
                (1, "Visible", "xl/worksheets/sheet1.xml", "one", "visible", 500, 26),
                (2, "Audit", "xl/worksheets/sheet2.xml", "two", "veryHidden", 0, 0),
            ),
        )
        for n in range(12):
            row_min = n * 20 + 1
            row_max = row_min + n + 1
            cursor = store.connection.execute(
                """
                INSERT INTO regions(
                    sheet_id, n, row_min, row_max, col_min, col_max,
                    header_rows, kind, list_object_name, confidence
                ) VALUES (1, ?, ?, ?, 1, 2, 1, 'region', NULL, ?)
                """,
                (n, row_min, row_max, 0.8 + n / 100),
            )
            assert cursor.lastrowid is not None
            region_id = cursor.lastrowid
            store.connection.executemany(
                """
                INSERT INTO columns(
                    region_id, idx, header, norm_header, dtype,
                    nonnull, distinct_est, formula_block_id
                ) VALUES (?, ?, ?, ?, ?, 1, 1, NULL)
                """,
                (
                    (region_id, 0, f"Header {n} A", f"header_{n}_a", "str"),
                    (region_id, 1, f"Header {n} B", f"header_{n}_b", "float"),
                ),
            )
        store.connection.executemany(
            """
            INSERT INTO defined_names(
                name, scope_sheet_id, refers_to, kind, is_builtin
            ) VALUES (?, NULL, ?, 'range', 0)
            """,
            ((f"Name{index:02d}", f"Visible!$A${index + 1}") for index in range(25)),
        )
        store.connection.executemany(
            """
            INSERT INTO diagnostics(
                severity, code, sheet_id, row, col, ref, message, related
            ) VALUES (?, ?, 1, NULL, NULL, '', 'test', '{}')
            """,
            (("error", "E_TEST"), ("warn", "W_ONE"), ("warn", "W_TWO")),
        )
        store.connection.execute(
            """
            INSERT INTO cells(
                sheet_id, row, col, ref, value, value_type, formula, style_idx,
                formula_kind, shared_index, array_ref, data_table
            ) VALUES (1, 500, 26, 'Z500', 'DO_NOT_LEAK_BODY_VALUE',
                      'string', NULL, 0, NULL, NULL, NULL, NULL)
            """
        )
        store.set_meta_many(
            {
                "indexed_at": "2026-07-15T12:00:00+00:00",
                "has_vba": 1,
                "external_links": json.dumps(
                    {str(index): f"../links/Book{index:02d}.xlsx" for index in range(1, 13)},
                    separators=(",", ":"),
                ),
            }
        )
    return workbook, database


def _seed_wide_map_index(tmp_path: Path) -> tuple[Path, Path]:
    workbook = tmp_path / "wide.xlsx"
    workbook.write_bytes(b"wide-map-test-placeholder")
    database = tmp_path / "wide.xlsp.db"
    with IndexStore(database) as store:
        store.connection.execute(
            """
            INSERT INTO sheets(
                id, name, xml_part, part_hash, kind, visibility, max_row, max_col
            ) VALUES (1, 'Wide', 'xl/worksheets/sheet1.xml', 'wide',
                      'worksheet', 'visible', 4, 30)
            """
        )
        cursor = store.connection.execute(
            """
            INSERT INTO regions(
                sheet_id, n, row_min, row_max, col_min, col_max,
                header_rows, kind, list_object_name, confidence
            ) VALUES (1, 0, 1, 4, 1, 30, 1, 'region', NULL, 0.88)
            """
        )
        assert cursor.lastrowid is not None
        store.connection.executemany(
            """
            INSERT INTO columns(
                region_id, idx, header, norm_header, dtype,
                nonnull, distinct_est, formula_block_id
            ) VALUES (?, ?, ?, ?, 'str', 3, 3, NULL)
            """,
            (
                (
                    cursor.lastrowid,
                    index,
                    f"Wide Header {index:02d} {'X' * 24}",
                    f"wide_header_{index:02d}",
                )
                for index in range(30)
            ),
        )
        store.set_meta_many(
            {
                "indexed_at": "2026-07-16T00:00:00+00:00",
                "has_vba": 0,
                "external_links": "{}",
            }
        )
    return workbook, database


def _seed_many_sheet_map_index(tmp_path: Path) -> tuple[Path, Path]:
    workbook = tmp_path / "many-sheets.xlsx"
    workbook.write_bytes(b"many-sheet-map-test-placeholder")
    database = tmp_path / "many-sheets.xlsp.db"
    with IndexStore(database) as store:
        store.connection.executemany(
            """
            INSERT INTO sheets(
                id, name, xml_part, part_hash, kind, visibility, max_row, max_col
            ) VALUES (?, ?, ?, ?, 'worksheet', ?, 0, 0)
            """,
            (
                (
                    number,
                    f"Sheet{number:02d}",
                    f"xl/worksheets/sheet{number}.xml",
                    f"hash-{number:02d}",
                    ("hidden" if number == 45 else "veryHidden" if number == 50 else "visible"),
                )
                for number in range(1, 51)
            ),
        )
        store.set_meta_many(
            {
                "indexed_at": "2026-07-16T00:00:00+00:00",
                "has_vba": 0,
                "external_links": "{}",
            }
        )
    return workbook, database


def _seed_visibility_pressure_index(tmp_path: Path) -> tuple[Path, Path]:
    workbook = tmp_path / "visibility-pressure.xlsx"
    workbook.write_bytes(b"visibility-pressure-map-test-placeholder")
    database = tmp_path / "visibility-pressure.xlsp.db"
    with IndexStore(database) as store:
        store.connection.executemany(
            """
            INSERT INTO sheets(
                id, name, xml_part, part_hash, kind, visibility, max_row, max_col
            ) VALUES (?, ?, ?, ?, 'worksheet', ?, 0, 0)
            """,
            (
                (
                    number,
                    f"Sheet{number:03d}",
                    f"xl/worksheets/sheet{number}.xml",
                    f"hash-{number:03d}",
                    ("visible" if number <= 20 else "hidden" if number <= 80 else "veryHidden"),
                )
                for number in range(1, 101)
            ),
        )
        store.set_meta_many(
            {
                "indexed_at": "2026-07-16T00:00:00+00:00",
                "has_vba": 0,
                "external_links": "{}",
            }
        )
    return workbook, database


def _seed_over_source_sheet_limit_index(tmp_path: Path) -> tuple[Path, Path]:
    workbook = tmp_path / "over-sheet-source-limit.xlsx"
    workbook.write_bytes(b"over-sheet-source-limit-map-test-placeholder")
    database = tmp_path / "over-sheet-source-limit.xlsp.db"
    with IndexStore(database) as store:
        store.connection.executemany(
            """
            INSERT INTO sheets(
                id, name, xml_part, part_hash, kind, visibility, max_row, max_col
            ) VALUES (?, ?, ?, ?, 'worksheet', ?, 0, 0)
            """,
            (
                (
                    sheet_id,
                    f"Bound{sheet_id:03d}",
                    f"xl/worksheets/sheet{sheet_id}.xml",
                    f"hash-{sheet_id:03d}",
                    "veryHidden" if sheet_id == 205 else "visible",
                )
                for sheet_id in range(1, 206)
            ),
        )
        store.set_meta_many(
            {
                "indexed_at": "2026-07-16T00:00:00+00:00",
                "has_vba": 0,
                "external_links": "{}",
            }
        )
    return workbook, database


def _seed_41_sheet_map_index(tmp_path: Path) -> tuple[Path, Path]:
    workbook = tmp_path / "forty-one.xlsx"
    workbook.write_bytes(b"forty-one-sheet-map-test-placeholder")
    database = tmp_path / "forty-one.xlsp.db"
    with IndexStore(database) as store:
        store.connection.executemany(
            """
            INSERT INTO sheets(
                id, name, xml_part, part_hash, kind, visibility, max_row, max_col
            ) VALUES (?, ?, ?, ?, 'worksheet', 'visible', 0, 0)
            """,
            (
                (
                    number,
                    f"S{number:02d}",
                    f"xl/worksheets/sheet{number}.xml",
                    f"hash-{number:02d}",
                )
                for number in range(1, 42)
            ),
        )
        store.set_meta_many(
            {
                "indexed_at": "2026-07-16T00:00:00+00:00",
                "has_vba": 0,
                "external_links": "{}",
            }
        )
    return workbook, database


def _seed_17_column_map_index(tmp_path: Path) -> tuple[Path, Path]:
    workbook = tmp_path / "seventeen-columns.xlsx"
    workbook.write_bytes(b"seventeen-column-map-test-placeholder")
    database = tmp_path / "seventeen-columns.xlsp.db"
    with IndexStore(database) as store:
        store.connection.execute(
            """
            INSERT INTO sheets(
                id, name, xml_part, part_hash, kind, visibility, max_row, max_col
            ) VALUES (
                1, 'Cols', 'xl/worksheets/sheet1.xml', 'cols',
                'worksheet', 'visible', 2, 17
            )
            """
        )
        cursor = store.connection.execute(
            """
            INSERT INTO regions(
                sheet_id, n, row_min, row_max, col_min, col_max,
                header_rows, kind, list_object_name, confidence
            ) VALUES (1, 0, 1, 2, 1, 17, 1, 'region', NULL, 0.9)
            """
        )
        assert cursor.lastrowid is not None
        store.connection.executemany(
            """
            INSERT INTO columns(
                region_id, idx, header, norm_header, dtype,
                nonnull, distinct_est, formula_block_id
            ) VALUES (?, ?, ?, ?, 'str', 1, 1, NULL)
            """,
            ((cursor.lastrowid, index, f"H{index}", f"h{index}") for index in range(17)),
        )
        store.set_meta_many(
            {
                "indexed_at": "2026-07-16T00:00:00+00:00",
                "has_vba": 0,
                "external_links": "{}",
            }
        )
    return workbook, database


def _seed_fblock_map_index(tmp_path: Path) -> tuple[Path, Path]:
    workbook = tmp_path / "fblocks.xlsx"
    workbook.write_bytes(b"fblock-map-test-placeholder")
    database = tmp_path / "fblocks.xlsp.db"
    with IndexStore(database) as store:
        store.connection.executemany(
            """
            INSERT INTO sheets(
                id, name, xml_part, part_hash, kind, visibility, max_row, max_col
            ) VALUES (?, ?, ?, ?, 'worksheet', 'visible', 7, 7)
            """,
            (
                (1, "One", "xl/worksheets/sheet1.xml", "one"),
                (2, "Two", "xl/worksheets/sheet2.xml", "two"),
            ),
        )
        region_rows = (
            (1, 0, 1, 5, 1, 2),
            (1, 1, 1, 5, 4, 5),
            (2, 0, 1, 3, 1, 3),
        )
        for sheet_id, n, row_min, row_max, col_min, col_max in region_rows:
            cursor = store.connection.execute(
                """
                INSERT INTO regions(
                    sheet_id, n, row_min, row_max, col_min, col_max,
                    header_rows, kind, list_object_name, confidence
                ) VALUES (?, ?, ?, ?, ?, ?, 1, 'region', NULL, 0.9)
                """,
                (sheet_id, n, row_min, row_max, col_min, col_max),
            )
            assert cursor.lastrowid is not None
            store.connection.executemany(
                """
                INSERT INTO columns(
                    region_id, idx, header, norm_header, dtype,
                    nonnull, distinct_est, formula_block_id
                ) VALUES (?, ?, ?, ?, 'float', 1, 1, NULL)
                """,
                (
                    (cursor.lastrowid, index, f"C{index}", f"c{index}")
                    for index in range(col_max - col_min + 1)
                ),
            )
        store.connection.executemany(
            """
            INSERT INTO fblocks(
                sheet_id, n, r1c1, row_min, row_max, col_min, col_max,
                volatile, opaque
            ) VALUES (?, ?, '=RC', ?, ?, ?, ?, 0, 0)
            """,
            (
                (1, 0, 2, 3, 2, 4),
                (1, 1, 2, 4, 1, 1),
                (1, 2, 2, 4, 5, 5),
                (1, 3, 1, 2, 7, 7),
                (2, 0, 2, 2, 2, 3),
                (2, 1, 1, 2, 5, 5),
            ),
        )
        store.set_meta_many(
            {
                "indexed_at": "2026-07-16T00:00:00+00:00",
                "has_vba": 0,
                "external_links": "{}",
            }
        )
    return workbook, database


def _seed_scaling_map_index(tmp_path: Path) -> tuple[Path, Path]:
    workbook = tmp_path / "scaling.xlsx"
    workbook.write_bytes(b"scaling-map-test-placeholder")
    database = tmp_path / "scaling.xlsp.db"
    with IndexStore(database) as store:
        store.connection.executemany(
            """
            INSERT INTO sheets(
                id, name, xml_part, part_hash, kind, visibility, max_row, max_col
            ) VALUES (?, ?, ?, ?, 'worksheet', ?, 500, 500)
            """,
            (
                (
                    sheet_id,
                    f"Scale{sheet_id:03d}",
                    f"xl/worksheets/sheet{sheet_id}.xml",
                    f"hash-{sheet_id:03d}",
                    (
                        "hidden"
                        if sheet_id in {119, 120}
                        else "veryHidden"
                        if sheet_id == 121
                        else "visible"
                    ),
                )
                for sheet_id in range(1, 122)
            ),
        )
        for sheet_id in range(1, 122):
            for region_n in range(12):
                row_min = region_n * 20 + 1
                cursor = store.connection.execute(
                    """
                    INSERT INTO regions(
                        sheet_id, n, row_min, row_max, col_min, col_max,
                        header_rows, kind, list_object_name, confidence
                    ) VALUES (?, ?, ?, ?, 1, 40, 1, 'region', NULL, 0.8)
                    """,
                    (sheet_id, region_n, row_min, row_min + 10),
                )
                assert cursor.lastrowid is not None
                store.connection.executemany(
                    """
                    INSERT INTO columns(
                        region_id, idx, header, norm_header, dtype,
                        nonnull, distinct_est, formula_block_id
                    ) VALUES (?, ?, ?, ?, 'str', 10, 10, NULL)
                    """,
                    (
                        (
                            cursor.lastrowid,
                            column_index,
                            f"Header {sheet_id:03d}-{region_n:02d}-{column_index:02d}",
                            f"h_{sheet_id:03d}_{region_n:02d}_{column_index:02d}",
                        )
                        for column_index in range(40)
                    ),
                )
        store.connection.executemany(
            """
            INSERT INTO defined_names(
                name, scope_sheet_id, refers_to, kind, is_builtin
            ) VALUES (?, NULL, ?, 'range', 0)
            """,
            ((f"ScaleName{index:03d}", f"Scale001!$A${index + 1}") for index in range(500)),
        )
        store.set_meta_many(
            {
                "indexed_at": "2026-07-16T00:00:00+00:00",
                "has_vba": 0,
                "external_links": json.dumps(
                    {str(index): f"../links/ScaleBook{index:02d}.xlsx" for index in range(1, 16)},
                    separators=(",", ":"),
                ),
            }
        )
    return workbook, database
