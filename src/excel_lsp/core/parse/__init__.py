"""OOXML package reading contracts."""

from excel_lsp.core.parse.parser import OOXMLParser
from excel_lsp.core.parse.styles import (
    BUILTIN_DATE_FORMAT_IDS,
    CellStyle,
    FillStyle,
    FontStyle,
    StyleCatalog,
    custom_format_is_date,
)

__all__ = (
    "BUILTIN_DATE_FORMAT_IDS",
    "CellStyle",
    "FillStyle",
    "FontStyle",
    "OOXMLParser",
    "StyleCatalog",
    "custom_format_is_date",
)
