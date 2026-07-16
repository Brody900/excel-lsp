"""SpreadsheetML style catalog and date-format classification."""

# pyright: reportAttributeAccessIssue=false, reportMissingTypeStubs=false
# pyright: reportUnknownArgumentType=false, reportUnknownMemberType=false
# pyright: reportUnknownParameterType=false, reportUnknownVariableType=false

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from types import MappingProxyType

from lxml import etree

from excel_lsp.core.parse._xml import (
    attr_by_local,
    child_by_local,
    children_by_local,
    local_name,
    parse_bool,
)

BUILTIN_DATE_FORMAT_IDS = frozenset(
    (*range(14, 23), *range(27, 37), *range(45, 48), *range(50, 59))
)


@dataclass(frozen=True, slots=True)
class FontStyle:
    """Font metadata needed by header detection."""

    bold: bool = False


@dataclass(frozen=True, slots=True)
class FillStyle:
    """Pattern-fill metadata needed by header detection."""

    pattern_type: str | None = None
    foreground: str | None = None
    background: str | None = None


@dataclass(frozen=True, slots=True)
class CellStyle:
    """One ``cellXfs`` entry."""

    num_fmt_id: int
    font_id: int
    fill_id: int
    is_date: bool


@dataclass(frozen=True, slots=True)
class StyleCatalog:
    """Immutable parsed workbook style metadata keyed by cell style index."""

    cell_xfs: tuple[CellStyle, ...]
    fonts: tuple[FontStyle, ...]
    fills: tuple[FillStyle, ...]
    custom_num_formats: Mapping[int, str]

    def is_date_style(self, style_idx: int) -> bool:
        """Return whether a valid cell style applies a date/time number format."""
        return 0 <= style_idx < len(self.cell_xfs) and self.cell_xfs[style_idx].is_date


DEFAULT_STYLE_CATALOG = StyleCatalog(
    cell_xfs=(CellStyle(0, 0, 0, False),),
    fonts=(FontStyle(),),
    fills=(FillStyle(),),
    custom_num_formats=MappingProxyType({}),
)


def custom_format_is_date(format_code: str) -> bool:
    """Apply the frozen token heuristic outside quotes and bracket blocks."""
    quoted = False
    bracket_depth = 0
    index = 0
    while index < len(format_code):
        character = format_code[index]
        if character == '"':
            quoted = not quoted
            index += 1
            continue
        if not quoted:
            if character == "[":
                bracket_depth += 1
                index += 1
                continue
            if character == "]" and bracket_depth:
                bracket_depth -= 1
                index += 1
                continue
            if character in {"\\", "_", "*"}:
                index += 2
                continue
            if bracket_depth == 0 and character.casefold() in {"y", "m", "d", "h", "s"}:
                return True
        index += 1
    return False


def parse_style_catalog(root: etree._Element) -> StyleCatalog:
    """Parse the style information required by cells and region heuristics."""
    custom_formats: dict[int, str] = {}
    fonts: list[FontStyle] = []
    fills: list[FillStyle] = []
    cell_xfs: list[CellStyle] = []

    for section in root:
        if not isinstance(section.tag, str):
            continue
        section_name = local_name(section.tag)
        if section_name == "numFmts":
            for num_fmt in children_by_local(section, "numFmt"):
                num_fmt_id = _required_int(num_fmt, "numFmtId")
                format_code = attr_by_local(num_fmt, "formatCode")
                if format_code is None:
                    raise ValueError("numFmt is missing formatCode")
                custom_formats[num_fmt_id] = format_code
        elif section_name == "fonts":
            for font in children_by_local(section, "font"):
                bold = child_by_local(font, "b")
                fonts.append(
                    FontStyle(
                        bold=bold is not None
                        and parse_bool(attr_by_local(bold, "val"), default=True)
                    )
                )
        elif section_name == "fills":
            for fill in children_by_local(section, "fill"):
                pattern = child_by_local(fill, "patternFill")
                if pattern is None:
                    fills.append(FillStyle())
                    continue
                foreground = child_by_local(pattern, "fgColor")
                background = child_by_local(pattern, "bgColor")
                fills.append(
                    FillStyle(
                        pattern_type=attr_by_local(pattern, "patternType"),
                        foreground=_color_value(foreground),
                        background=_color_value(background),
                    )
                )
        elif section_name == "cellXfs":
            for xf in children_by_local(section, "xf"):
                num_fmt_id = _optional_int(xf, "numFmtId", 0)
                font_id = _optional_int(xf, "fontId", 0)
                fill_id = _optional_int(xf, "fillId", 0)
                format_code = custom_formats.get(num_fmt_id)
                is_date = num_fmt_id in BUILTIN_DATE_FORMAT_IDS or (
                    format_code is not None and custom_format_is_date(format_code)
                )
                cell_xfs.append(CellStyle(num_fmt_id, font_id, fill_id, is_date))

    if not fonts:
        fonts.append(FontStyle())
    if not fills:
        fills.append(FillStyle())
    if not cell_xfs:
        cell_xfs.append(CellStyle(0, 0, 0, False))
    return StyleCatalog(
        cell_xfs=tuple(cell_xfs),
        fonts=tuple(fonts),
        fills=tuple(fills),
        custom_num_formats=MappingProxyType(custom_formats),
    )


def _required_int(element: etree._Element, name: str) -> int:
    value = attr_by_local(element, name)
    if value is None:
        raise ValueError(f"{local_name(element.tag)} is missing {name}")
    return int(value)


def _optional_int(element: etree._Element, name: str, default: int) -> int:
    value = attr_by_local(element, name)
    return default if value is None else int(value)


def _color_value(element: etree._Element | None) -> str | None:
    if element is None:
        return None
    for name in ("rgb", "indexed", "theme", "auto"):
        value = attr_by_local(element, name)
        if value is not None:
            return f"{name}:{value}"
    return None
