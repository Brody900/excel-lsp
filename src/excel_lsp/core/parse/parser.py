"""Read-only, streaming OOXML package parser."""

# pyright: reportAttributeAccessIssue=false, reportMissingTypeStubs=false
# pyright: reportUnknownArgumentType=false, reportUnknownMemberType=false
# pyright: reportUnknownParameterType=false, reportUnknownVariableType=false

from __future__ import annotations

import hashlib
import math
import posixpath
import re
import time as time_module
import zipfile
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation
from pathlib import Path
from types import MappingProxyType
from typing import IO, cast

from lxml import etree
from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900, from_excel

from excel_lsp.core.errors import ErrorCode, ExcelLSPError
from excel_lsp.core.models import (
    CalculationMode,
    CalculationProperties,
    CellRecord,
    CellScalar,
    CellValueType,
    DataTableFormulaInfo,
    DataValidationInfo,
    DefinedName,
    DefinedNameKind,
    NameArea,
    PackageHashes,
    Rect,
    ReferenceMode,
    SheetDescriptor,
    SheetKind,
    SheetParseSummary,
    TableInfo,
    Visibility,
    WorkbookMetadata,
)
from excel_lsp.core.parse._xml import (
    attr_by_local,
    child_by_local,
    clear_element,
    local_name,
    parse_bool,
    parse_xml,
    text_content,
)
from excel_lsp.core.parse.coordinates import (
    contains,
    make_cell_ref,
    parse_cell_ref,
    parse_rect,
)
from excel_lsp.core.parse.styles import (
    DEFAULT_STYLE_CATALOG,
    StyleCatalog,
    parse_style_catalog,
)

_SUPPORTED_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xltx", ".xltm"})
_CFB_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")
_NUMERIC_LITERAL = re.compile(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?%?")
_ERROR_LITERAL = re.compile(r"#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA)", re.I)
_LAMBDA_PREFIX = re.compile(r"(?:_xlfn\.)?(?:_xlws\.)?LAMBDA\s*\(", re.I)


class _PackageCorrupt(Exception):
    """Internal marker for malformed package content."""


@dataclass(frozen=True, slots=True)
class _Relationship:
    rel_id: str
    rel_type: str
    target: str
    target_mode: str | None
    part: str | None


@dataclass(frozen=True, slots=True)
class _ArraySpan:
    ref: str
    rect: Rect


@dataclass(frozen=True, slots=True)
class _SharedFormula:
    formula: str
    origin: str
    rect: Rect


class OOXMLParser:
    """Own one read-only OOXML package and stream its worksheet cells.

    Metadata and selected part hashes are available only while the context is
    entered.  Worksheet XML is parsed exactly once per ``parse_sheet`` call and
    cells are handed to the supplied callback as soon as each element ends.
    """

    retry_delay_seconds = 0.5

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self._archive: zipfile.ZipFile | None = None
        self._members: dict[str, zipfile.ZipInfo] = {}
        self._metadata: WorkbookMetadata | None = None
        self._hashes: PackageHashes | None = None
        self._styles: StyleCatalog | None = None
        self._shared_strings: tuple[str, ...] = ()

    @property
    def metadata(self) -> WorkbookMetadata:
        """Return workbook metadata after the parser context is entered."""
        if self._metadata is None:
            raise RuntimeError("OOXMLParser must be entered before metadata is read")
        return self._metadata

    @property
    def hashes(self) -> PackageHashes:
        """Return whole-file and selected-part SHA-256 digests."""
        if self._hashes is None:
            raise RuntimeError("OOXMLParser must be entered before hashes are read")
        return self._hashes

    @property
    def styles(self) -> StyleCatalog:
        """Return immutable style metadata keyed by ``CellRecord.style_idx``."""
        if self._styles is None:
            raise RuntimeError("OOXMLParser must be entered before styles are read")
        return self._styles

    def __enter__(self) -> OOXMLParser:
        if self._archive is not None:
            raise RuntimeError("OOXMLParser context is already entered")
        self._validate_path()

        final_error: tuple[ErrorCode, str] | None = None
        for attempt in range(2):
            try:
                self._initialize_once()
                return self
            except ExcelLSPError:
                self.close()
                raise
            except FileNotFoundError:
                self.close()
                raise self._not_found_error() from None
            except OSError as error:
                self.close()
                if _is_lock_error(error):
                    final_error = (
                        ErrorCode.LOCKED,
                        "Workbook could not be opened because it is locked.",
                    )
                else:
                    final_error = (
                        ErrorCode.CORRUPT,
                        "Workbook package could not be read as a consistent OOXML file.",
                    )
            except (zipfile.BadZipFile, EOFError, etree.XMLSyntaxError, _PackageCorrupt):
                self.close()
                final_error = (
                    ErrorCode.CORRUPT,
                    "Workbook package is missing or contains malformed OOXML parts.",
                )
            if attempt == 0:
                time_module.sleep(self.retry_delay_seconds)

        assert final_error is not None
        code, message = final_error
        hint = (
            "Close Excel or wait for the current save to finish, then retry."
            if code is ErrorCode.LOCKED
            else "Wait for any in-progress save to finish, then reopen the workbook."
        )
        raise ExcelLSPError(code, message, hint=hint, details={"path": str(self.path)})

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: object | None,
    ) -> None:
        del exc_type, exc_value, traceback
        self.close()

    def close(self) -> None:
        """Close the package handle; safe to call repeatedly."""
        archive = self._archive
        self._archive = None
        if archive is not None:
            archive.close()

    def collect_cells(self, sheet: SheetDescriptor) -> tuple[CellRecord, ...]:
        """Collect a sheet stream for tests and the openpyxl oracle."""
        cells: list[CellRecord] = []
        self.parse_sheet(sheet, cells.append)
        return tuple(cells)

    def parse_sheet(
        self,
        sheet: SheetDescriptor,
        on_cell: Callable[[CellRecord], None],
    ) -> SheetParseSummary:
        """Stream one worksheet with one clearing ``lxml.iterparse`` pass."""
        self._require_open()
        if sheet not in self.metadata.sheets:
            raise ValueError("sheet descriptor does not belong to this workbook")
        part_hash = self.hashes.parts.get(sheet.xml_part)
        if part_hash is None:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                "Sheet part has no package hash.",
                details={"part": sheet.xml_part},
            )
        if sheet.kind != "worksheet":
            return SheetParseSummary(
                descriptor=sheet,
                part_hash=part_hash,
                max_row=0,
                max_col=0,
                cell_count=0,
            )

        shared_formulas: dict[int, _SharedFormula] = {}
        arrays: list[_ArraySpan] = []
        data_tables: list[DataTableFormulaInfo] = []
        merges: list[Rect] = []
        validations: list[DataValidationInfo] = []
        table_rel_ids: list[str] = []
        dimension_ref: str | None = None
        max_row = 0
        max_col = 0
        cell_count = 0
        current_row = 0
        previous_row = 0
        previous_col = 0

        try:
            with self._open_part(sheet.xml_part) as stream:
                context = etree.iterparse(
                    stream,
                    events=("start", "end"),
                    resolve_entities=False,
                    no_network=True,
                    recover=False,
                    huge_tree=True,
                )
                for event, element in context:
                    if not isinstance(element.tag, str):
                        if event == "end":
                            clear_element(element)
                        continue
                    name = local_name(element.tag)
                    if event == "start":
                        if name == "row":
                            row_text = attr_by_local(element, "r")
                            if row_text is None:
                                current_row = previous_row + 1
                            else:
                                current_row = _bounded_int(row_text, 1, 1_048_576, "row number")
                            if current_row <= previous_row:
                                raise _PackageCorrupt(
                                    "worksheet rows are duplicated or out of order"
                                )
                            previous_row = current_row
                            previous_col = 0
                        continue

                    if name == "c":
                        cell = self._parse_cell(
                            element,
                            current_row=current_row,
                            implicit_column=previous_col + 1,
                            shared_formulas=shared_formulas,
                            arrays=arrays,
                            data_tables=data_tables,
                        )
                        if cell is not None:
                            if current_row and cell.row != current_row:
                                raise _PackageCorrupt(
                                    f"cell {cell.ref} does not belong to its containing row"
                                )
                            if cell.col <= previous_col:
                                raise _PackageCorrupt(
                                    "worksheet cells are duplicated or out of order"
                                )
                            previous_col = cell.col
                            on_cell(cell)
                            cell_count += 1
                            max_row = max(max_row, cell.row)
                            max_col = max(max_col, cell.col)
                        else:
                            raw_ref = attr_by_local(element, "r")
                            if raw_ref is not None:
                                raw_row, raw_col = _checked_cell_ref(raw_ref)
                                if current_row and raw_row != current_row:
                                    raise _PackageCorrupt(
                                        f"cell {raw_ref} does not belong to its containing row"
                                    )
                                if raw_col <= previous_col:
                                    raise _PackageCorrupt(
                                        "worksheet cells are duplicated or out of order"
                                    )
                                previous_col = raw_col
                            else:
                                previous_col += 1
                        clear_element(element)
                        continue

                    if _has_ancestor(element, {"c", "dataValidation"}):
                        continue

                    if name == "dimension" and dimension_ref is None:
                        dimension_ref = attr_by_local(element, "ref")
                    elif name == "mergeCell":
                        merge_ref = attr_by_local(element, "ref")
                        if merge_ref is None:
                            raise _PackageCorrupt("mergeCell has no ref")
                        merges.append(_checked_rect(merge_ref))
                    elif name == "dataValidation":
                        validations.extend(self._parse_validation(element))
                    elif name == "tablePart":
                        rel_id = attr_by_local(element, "id")
                        if rel_id is None:
                            raise _PackageCorrupt("tablePart has no relationship id")
                        table_rel_ids.append(rel_id)
                    clear_element(element)
        except ExcelLSPError:
            raise
        except (zipfile.BadZipFile, EOFError, etree.XMLSyntaxError, _PackageCorrupt) as error:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                "Worksheet XML is malformed or incomplete.",
                details={"part": sheet.xml_part},
            ) from error

        tables = self._load_tables(sheet.xml_part, table_rel_ids)
        return SheetParseSummary(
            descriptor=sheet,
            part_hash=part_hash,
            max_row=max_row,
            max_col=max_col,
            cell_count=cell_count,
            dimension_ref=dimension_ref,
            merges=tuple(merges),
            validations=tuple(validations),
            tables=tables,
            array_formulas=tuple(span.rect for span in arrays),
            data_tables=tuple(data_tables),
        )

    def _validate_path(self) -> None:
        if not self.path.exists() or not self.path.is_file():
            raise self._not_found_error()
        if self.path.suffix.casefold() not in _SUPPORTED_SUFFIXES:
            raise ExcelLSPError(
                ErrorCode.UNSUPPORTED_FORMAT,
                "Only OOXML Excel workbooks are supported.",
                hint="Use an .xlsx, .xlsm, .xltx, or .xltm file.",
                details={"path": str(self.path)},
            )

    def _not_found_error(self) -> ExcelLSPError:
        return ExcelLSPError(
            ErrorCode.NOT_FOUND,
            "Workbook file was not found.",
            details={"path": str(self.path)},
        )

    def _initialize_once(self) -> None:
        whole_hash, signature = _hash_file(self.path)
        if signature == _CFB_SIGNATURE:
            raise ExcelLSPError(
                ErrorCode.ENCRYPTED,
                "Encrypted Office workbooks cannot be indexed.",
                hint="Open the workbook in Excel, remove its password, and save an OOXML copy.",
                details={"path": str(self.path)},
            )

        archive = zipfile.ZipFile(self.path, mode="r")
        self._archive = archive
        self._members = self._index_members(archive)
        if any(info.flag_bits & 0x1 for info in self._members.values()):
            raise ExcelLSPError(
                ErrorCode.ENCRYPTED,
                "Encrypted ZIP entries cannot be indexed.",
                hint="Remove workbook encryption in Excel and save a new copy.",
                details={"path": str(self.path)},
            )

        workbook_root = parse_xml(self._read_required("xl/workbook.xml"))
        workbook_rels = self._parse_relationships(
            "xl/_rels/workbook.xml.rels", source_part="xl/workbook.xml", required=True
        )
        content_types = self._parse_content_types()
        sheets = self._parse_sheets(workbook_root, workbook_rels, content_types)
        date1904 = self._parse_date_system(workbook_root)
        calculation = self._parse_calculation_properties(workbook_root)
        external_links = self._parse_external_links(workbook_root, workbook_rels)
        defined_names = self._parse_defined_names(workbook_root, sheets)

        if self._has_part("xl/styles.xml"):
            try:
                self._styles = parse_style_catalog(parse_xml(self._read_required("xl/styles.xml")))
            except ValueError as error:
                raise _PackageCorrupt("styles.xml contains invalid style metadata") from error
        else:
            self._styles = DEFAULT_STYLE_CATALOG
        self._shared_strings = self._parse_shared_strings()

        selected_parts = ["xl/workbook.xml", "xl/_rels/workbook.xml.rels"]
        if self._has_part("[Content_Types].xml"):
            selected_parts.append("[Content_Types].xml")
        selected_parts.extend(
            part for part in ("xl/sharedStrings.xml", "xl/styles.xml") if self._has_part(part)
        )
        selected_parts.extend(sheet.xml_part for sheet in sheets)
        selected_parts.extend(part for sheet in sheets for part in sheet.related_parts)
        part_hashes: dict[str, str] = {}
        for part in selected_parts:
            if part not in part_hashes:
                part_hashes[part] = self._hash_part(part)

        self._metadata = WorkbookMetadata(
            path=str(self.path.resolve()),
            date1904=date1904,
            sheets=sheets,
            defined_names=defined_names,
            calculation=calculation,
            external_links=MappingProxyType(external_links),
            has_vba=self._has_part("xl/vbaProject.bin"),
        )
        self._hashes = PackageHashes(
            whole_file=whole_hash,
            parts=MappingProxyType(part_hashes),
        )

    def _index_members(self, archive: zipfile.ZipFile) -> dict[str, zipfile.ZipInfo]:
        members: dict[str, zipfile.ZipInfo] = {}
        for info in archive.infolist():
            if info.is_dir():
                continue
            normalized = _normalize_part_name(info.filename)
            if normalized in members:
                raise _PackageCorrupt(f"duplicate package member: {normalized}")
            members[normalized] = info
        return members

    def _read_required(self, part: str) -> bytes:
        normalized = _normalize_part_name(part)
        info = self._members.get(normalized)
        archive = self._require_open()
        if info is None:
            raise _PackageCorrupt(f"required package part is missing: {normalized}")
        try:
            return archive.read(info)
        except RuntimeError as error:
            if "password" in str(error).casefold() or info.flag_bits & 0x1:
                raise ExcelLSPError(
                    ErrorCode.ENCRYPTED,
                    "Encrypted workbook parts cannot be indexed.",
                ) from error
            raise _PackageCorrupt(f"could not read package part: {normalized}") from error

    def _open_part(self, part: str) -> IO[bytes]:
        normalized = _normalize_part_name(part)
        info = self._members.get(normalized)
        if info is None:
            raise _PackageCorrupt(f"required package part is missing: {normalized}")
        try:
            return self._require_open().open(info, mode="r")
        except RuntimeError as error:
            if "password" in str(error).casefold() or info.flag_bits & 0x1:
                raise ExcelLSPError(ErrorCode.ENCRYPTED, "Workbook part is encrypted.") from error
            raise _PackageCorrupt(f"could not open package part: {normalized}") from error

    def _has_part(self, part: str) -> bool:
        return _normalize_part_name(part) in self._members

    def _hash_part(self, part: str) -> str:
        digest = hashlib.sha256()
        with self._open_part(part) as stream:
            while chunk := stream.read(1024 * 1024):
                digest.update(chunk)
        return digest.hexdigest()

    def _parse_content_types(self) -> dict[str, str]:
        if not self._has_part("[Content_Types].xml"):
            return {}
        root = parse_xml(self._read_required("[Content_Types].xml"))
        overrides: dict[str, str] = {}
        for element in root.iter():
            if not isinstance(element.tag, str) or local_name(element.tag) != "Override":
                continue
            part = attr_by_local(element, "PartName")
            content_type = attr_by_local(element, "ContentType")
            if part is not None and content_type is not None:
                overrides[_normalize_part_name(part)] = content_type
        return overrides

    def _parse_relationships(
        self,
        rels_part: str,
        *,
        source_part: str,
        required: bool,
    ) -> dict[str, _Relationship]:
        if not self._has_part(rels_part):
            if required:
                raise _PackageCorrupt(f"required relationships part is missing: {rels_part}")
            return {}
        root = parse_xml(self._read_required(rels_part))
        relationships: dict[str, _Relationship] = {}
        for element in root.iter():
            if not isinstance(element.tag, str) or local_name(element.tag) != "Relationship":
                continue
            rel_id = attr_by_local(element, "Id")
            rel_type = attr_by_local(element, "Type")
            target = attr_by_local(element, "Target")
            target_mode = attr_by_local(element, "TargetMode")
            if rel_id is None or rel_type is None or target is None:
                raise _PackageCorrupt("relationship is missing Id, Type, or Target")
            if rel_id in relationships:
                raise _PackageCorrupt(f"duplicate relationship id: {rel_id}")
            is_external = target_mode is not None and target_mode.casefold() == "external"
            part = None if is_external else _resolve_target(source_part, target)
            relationships[rel_id] = _Relationship(rel_id, rel_type, target, target_mode, part)
        return relationships

    def _parse_sheets(
        self,
        workbook_root: etree._Element,
        workbook_rels: dict[str, _Relationship],
        content_types: dict[str, str],
    ) -> tuple[SheetDescriptor, ...]:
        descriptors: list[SheetDescriptor] = []
        sheet_names: set[str] = set()
        sheet_ids: set[int] = set()
        relationship_ids: set[str] = set()
        for element in workbook_root.iter():
            if not isinstance(element.tag, str) or local_name(element.tag) != "sheet":
                continue
            name = attr_by_local(element, "name")
            sheet_id_text = attr_by_local(element, "sheetId")
            rel_id = attr_by_local(element, "id")
            if name is None or sheet_id_text is None or rel_id is None:
                raise _PackageCorrupt("workbook sheet is missing name, sheetId, or relationship")
            relationship = workbook_rels.get(rel_id)
            if relationship is None or relationship.part is None:
                raise _PackageCorrupt(f"sheet relationship is missing or external: {rel_id}")
            sheet_id = _bounded_int(sheet_id_text, 1, 2_147_483_647, "sheetId")
            normalized_name = name.casefold()
            if normalized_name in sheet_names:
                raise _PackageCorrupt(f"duplicate workbook sheet name: {name!r}")
            if sheet_id in sheet_ids:
                raise _PackageCorrupt(f"duplicate workbook sheetId: {sheet_id}")
            if rel_id in relationship_ids:
                raise _PackageCorrupt(f"duplicate workbook sheet relationship: {rel_id}")
            sheet_names.add(normalized_name)
            sheet_ids.add(sheet_id)
            relationship_ids.add(rel_id)
            kind = _classify_sheet(
                relationship.rel_type,
                relationship.part,
                content_types.get(relationship.part),
            )
            visibility = _parse_visibility(attr_by_local(element, "state"))
            related_parts = self._sheet_related_parts(relationship.part, kind)
            descriptors.append(
                SheetDescriptor(
                    order=len(descriptors),
                    name=name,
                    sheet_id=sheet_id,
                    rel_id=rel_id,
                    xml_part=relationship.part,
                    kind=kind,
                    visibility=visibility,
                    related_parts=related_parts,
                )
            )
        return tuple(descriptors)

    def _sheet_related_parts(self, sheet_part: str, kind: SheetKind) -> tuple[str, ...]:
        if kind != "worksheet":
            return ()
        rels_part = _relationships_part(sheet_part)
        if not self._has_part(rels_part):
            return ()
        relationships = self._parse_relationships(
            rels_part,
            source_part=sheet_part,
            required=False,
        )
        table_parts: set[str] = set()
        for relationship in relationships.values():
            rel_kind = relationship.rel_type.rsplit("/", 1)[-1].casefold()
            if rel_kind != "table":
                continue
            if relationship.part is None or not self._has_part(relationship.part):
                raise _PackageCorrupt("worksheet table relationship targets a missing part")
            table_parts.add(relationship.part)
        return (rels_part, *sorted(table_parts))

    def _parse_date_system(self, workbook_root: etree._Element) -> bool:
        for element in workbook_root.iter():
            if isinstance(element.tag, str) and local_name(element.tag) == "workbookPr":
                return parse_bool(attr_by_local(element, "date1904"))
        return False

    def _parse_calculation_properties(
        self,
        workbook_root: etree._Element,
    ) -> CalculationProperties:
        elements = [
            element
            for element in workbook_root.iter()
            if isinstance(element.tag, str) and local_name(element.tag) == "calcPr"
        ]
        if not elements:
            return CalculationProperties()
        if len(elements) > 1:
            raise _PackageCorrupt("workbook contains more than one calcPr element")
        element = elements[0]
        calc_mode_text = attr_by_local(element, "calcMode")
        if calc_mode_text not in {None, "manual", "auto", "autoNoTable"}:
            raise _PackageCorrupt(f"invalid calcMode: {calc_mode_text!r}")
        ref_mode_text = attr_by_local(element, "refMode")
        if ref_mode_text not in {None, "A1", "R1C1"}:
            raise _PackageCorrupt(f"invalid refMode: {ref_mode_text!r}")
        return CalculationProperties(
            calc_id=_optional_bounded_int(
                attr_by_local(element, "calcId"),
                0,
                0,
                4_294_967_295,
                "calcId",
            )
            if attr_by_local(element, "calcId") is not None
            else None,
            calc_mode=cast(CalculationMode | None, calc_mode_text),
            full_calc_on_load=_optional_xml_bool(
                attr_by_local(element, "fullCalcOnLoad"), "fullCalcOnLoad"
            ),
            ref_mode=cast(ReferenceMode | None, ref_mode_text),
            iterate=_optional_xml_bool(attr_by_local(element, "iterate"), "iterate"),
            iterate_count=_optional_bounded_int(
                attr_by_local(element, "iterateCount"),
                0,
                0,
                2_147_483_647,
                "iterateCount",
            )
            if attr_by_local(element, "iterateCount") is not None
            else None,
            iterate_delta=_optional_nonnegative_float(
                attr_by_local(element, "iterateDelta"), "iterateDelta"
            ),
            full_precision=_optional_xml_bool(
                attr_by_local(element, "fullPrecision"), "fullPrecision"
            ),
            calc_completed=_optional_xml_bool(
                attr_by_local(element, "calcCompleted"), "calcCompleted"
            ),
            calc_on_save=_optional_xml_bool(attr_by_local(element, "calcOnSave"), "calcOnSave"),
            concurrent_calc=_optional_xml_bool(
                attr_by_local(element, "concurrentCalc"), "concurrentCalc"
            ),
            concurrent_manual_count=_optional_bounded_int(
                attr_by_local(element, "concurrentManualCount"),
                0,
                0,
                2_147_483_647,
                "concurrentManualCount",
            )
            if attr_by_local(element, "concurrentManualCount") is not None
            else None,
            force_full_calc=_optional_xml_bool(
                attr_by_local(element, "forceFullCalc"), "forceFullCalc"
            ),
        )

    def _parse_defined_names(
        self,
        workbook_root: etree._Element,
        sheets: tuple[SheetDescriptor, ...],
    ) -> tuple[DefinedName, ...]:
        defined_names: list[DefinedName] = []
        for element in workbook_root.iter():
            if not isinstance(element.tag, str) or local_name(element.tag) != "definedName":
                continue
            name = attr_by_local(element, "name")
            if name is None:
                raise _PackageCorrupt("definedName has no name")
            scope_text = attr_by_local(element, "localSheetId")
            scope = (
                None
                if scope_text is None
                else _bounded_int(scope_text, 0, max(0, len(sheets) - 1), "localSheetId")
            )
            refers_to = "".join(element.itertext())
            kind, areas = _classify_defined_name(refers_to, scope, sheets)
            defined_names.append(
                DefinedName(
                    name=name,
                    refers_to=refers_to,
                    scope_sheet_order=scope,
                    kind=kind,
                    is_builtin=name.casefold().startswith("_xlnm."),
                    areas=areas,
                )
            )
        return tuple(defined_names)

    def _parse_external_links(
        self,
        workbook_root: etree._Element,
        workbook_rels: dict[str, _Relationship],
    ) -> dict[int, str]:
        links: dict[int, str] = {}
        link_index = 0
        for element in workbook_root.iter():
            if not isinstance(element.tag, str) or local_name(element.tag) != "externalReference":
                continue
            link_index += 1
            rel_id = attr_by_local(element, "id")
            relationship = None if rel_id is None else workbook_rels.get(rel_id)
            if (
                relationship is None
                or relationship.part is None
                or not self._has_part(relationship.part)
            ):
                continue
            target = self._external_link_target(relationship.part)
            if target is not None:
                links[link_index] = target
        return links

    def _external_link_target(self, link_part: str) -> str | None:
        root = parse_xml(self._read_required(link_part))
        external_book_rel_id: str | None = None
        for element in root.iter():
            if isinstance(element.tag, str) and local_name(element.tag) == "externalBook":
                external_book_rel_id = attr_by_local(element, "id")
                break
        relationships = self._parse_relationships(
            _relationships_part(link_part), source_part=link_part, required=False
        )
        if external_book_rel_id is not None:
            relationship = relationships.get(external_book_rel_id)
            if relationship is not None:
                return relationship.target
        for relationship in relationships.values():
            if (
                relationship.target_mode is not None
                and relationship.target_mode.casefold() == "external"
            ):
                return relationship.target
        return None

    def _parse_shared_strings(self) -> tuple[str, ...]:
        if not self._has_part("xl/sharedStrings.xml"):
            return ()
        values: list[str] = []
        with self._open_part("xl/sharedStrings.xml") as stream:
            context = etree.iterparse(
                stream,
                events=("end",),
                resolve_entities=False,
                no_network=True,
                recover=False,
                huge_tree=True,
            )
            for _, element in context:
                if not isinstance(element.tag, str):
                    clear_element(element)
                    continue
                name = local_name(element.tag)
                if name == "si":
                    values.append(text_content(element))
                    clear_element(element)
                elif not _has_ancestor(element, {"si"}):
                    clear_element(element)
        return tuple(values)

    def _parse_cell(
        self,
        element: etree._Element,
        *,
        current_row: int,
        implicit_column: int,
        shared_formulas: dict[int, _SharedFormula],
        arrays: list[_ArraySpan],
        data_tables: list[DataTableFormulaInfo],
    ) -> CellRecord | None:
        raw_ref = attr_by_local(element, "r")
        if raw_ref is None:
            row = current_row or 1
            col = implicit_column
            try:
                ref = make_cell_ref(row, col)
            except ValueError as error:
                raise _PackageCorrupt("cell without r has invalid implicit coordinates") from error
        else:
            row, col = _checked_cell_ref(raw_ref)
            ref = make_cell_ref(row, col)

        style_text = attr_by_local(element, "s")
        style_idx = 0 if style_text is None else _bounded_int(style_text, 0, 2_147_483_647, "style")
        type_code = (attr_by_local(element, "t") or "n").casefold()
        formula_element = child_by_local(element, "f")
        value_element = child_by_local(element, "v")
        inline_element = child_by_local(element, "is")
        has_value = value_element is not None or inline_element is not None
        if formula_element is None and not has_value:
            return None

        value, value_type = self._parse_cell_value(
            type_code,
            value_element,
            inline_element,
            style_idx,
        )
        formula: str | None = None
        formula_kind = None
        shared_index: int | None = None
        array_ref: str | None = None
        data_table: DataTableFormulaInfo | None = None

        if formula_element is not None:
            formula_type = (attr_by_local(formula_element, "t") or "normal").casefold()
            formula_body = formula_element.text or ""
            if formula_type == "shared":
                shared_text = attr_by_local(formula_element, "si")
                if shared_text is None:
                    raise _PackageCorrupt("shared formula has no si")
                shared_index = _bounded_int(shared_text, 0, 2_147_483_647, "shared formula si")
                formula_kind = "shared"
                if formula_body:
                    formula = _formula_with_equals(formula_body)
                    group_ref = attr_by_local(formula_element, "ref")
                    if group_ref is None:
                        raise _PackageCorrupt("shared formula master has no ref span")
                    group_rect = _checked_rect(group_ref)
                    if not contains(group_rect, row, col):
                        raise _PackageCorrupt(
                            f"shared formula master {ref} is outside its ref span"
                        )
                    if shared_index in shared_formulas:
                        raise _PackageCorrupt(
                            f"shared formula si {shared_index} has more than one master"
                        )
                    shared_formulas[shared_index] = _SharedFormula(formula, ref, group_rect)
                else:
                    master = shared_formulas.get(shared_index)
                    if master is None:
                        raise _PackageCorrupt(
                            f"shared formula follower {ref} appears before its master"
                        )
                    if not contains(master.rect, row, col):
                        raise _PackageCorrupt(
                            f"shared formula follower {ref} is outside its master ref span"
                        )
                    try:
                        formula = cast(
                            str,
                            Translator(master.formula, origin=master.origin).translate_formula(ref),
                        )
                    except (TranslatorError, ValueError) as error:
                        raise _PackageCorrupt(
                            f"could not translate shared formula at {ref}"
                        ) from error
            elif formula_type == "array":
                if not formula_body:
                    raise _PackageCorrupt("array formula master has no formula text")
                span_ref = attr_by_local(formula_element, "ref")
                if span_ref is None:
                    raise _PackageCorrupt("array formula has no ref span")
                span = _ArraySpan(span_ref, _checked_rect(span_ref))
                if not contains(span.rect, row, col):
                    raise _PackageCorrupt(f"array formula master {ref} is outside its ref span")
                if span not in arrays:
                    arrays.append(span)
                formula = _formula_with_equals(formula_body)
                formula_kind = "array"
                array_ref = span_ref
            elif formula_type == "datatable":
                span_ref = attr_by_local(formula_element, "ref")
                if span_ref is None:
                    raise _PackageCorrupt("data table formula has no ref span")
                span_rect = _checked_rect(span_ref)
                if not contains(span_rect, row, col):
                    raise _PackageCorrupt(
                        f"data table formula anchor {ref} is outside its ref span"
                    )
                input_cell_1 = attr_by_local(formula_element, "r1")
                input_cell_2 = attr_by_local(formula_element, "r2")
                if input_cell_1 is not None:
                    _checked_cell_ref(input_cell_1)
                if input_cell_2 is not None:
                    _checked_cell_ref(input_cell_2)
                data_table = DataTableFormulaInfo(
                    ref=span_ref,
                    rect=span_rect,
                    input_cell_1=input_cell_1,
                    input_cell_2=input_cell_2,
                    is_2d=_optional_xml_bool(attr_by_local(formula_element, "dt2D"), "dt2D")
                    or False,
                    row_oriented=_optional_xml_bool(attr_by_local(formula_element, "dtr"), "dtr")
                    or False,
                    calculate_always=_optional_xml_bool(attr_by_local(formula_element, "ca"), "ca")
                    or False,
                    deleted_row_input=_optional_xml_bool(
                        attr_by_local(formula_element, "del1"), "del1"
                    )
                    or False,
                    deleted_column_input=_optional_xml_bool(
                        attr_by_local(formula_element, "del2"), "del2"
                    )
                    or False,
                )
                data_tables.append(data_table)
                formula_kind = "dataTable"
                array_ref = span_ref
            else:
                if not formula_body:
                    raise _PackageCorrupt("formula cell has no formula text")
                formula = _formula_with_equals(formula_body)
                formula_kind = "normal"

        if formula_kind not in {"array", "dataTable"}:
            for span in reversed(arrays):
                if contains(span.rect, row, col):
                    formula_kind = "array"
                    array_ref = span.ref
                    break

        if formula_kind not in {"array", "dataTable"}:
            for candidate in reversed(data_tables):
                if contains(candidate.rect, row, col):
                    formula_kind = "dataTable"
                    array_ref = candidate.ref
                    data_table = candidate
                    break

        return CellRecord(
            ref=ref,
            row=row,
            col=col,
            value=value,
            value_type=value_type,
            formula=formula,
            style_idx=style_idx,
            formula_kind=formula_kind,
            shared_index=shared_index,
            array_ref=array_ref,
            data_table=data_table,
        )

    def _parse_cell_value(
        self,
        type_code: str,
        value_element: etree._Element | None,
        inline_element: etree._Element | None,
        style_idx: int,
    ) -> tuple[CellScalar, CellValueType]:
        raw = None if value_element is None else value_element.text
        if type_code == "inlinestr":
            if inline_element is None:
                return None, "blank"
            return text_content(inline_element), "string"
        if raw is None:
            if value_element is not None and type_code in {"str", "s", "e"}:
                raw = ""
            else:
                return None, "blank"
        if type_code == "s":
            try:
                index = int(raw)
            except ValueError as error:
                raise _PackageCorrupt("shared-string cell has a non-integer index") from error
            if not 0 <= index < len(self._shared_strings):
                raise _PackageCorrupt("shared-string index is outside sharedStrings.xml")
            return self._shared_strings[index], "string"
        if type_code == "str":
            return raw, "string"
        if type_code == "b":
            normalized = raw.strip().casefold()
            if normalized in {"1", "true"}:
                return True, "bool"
            if normalized in {"0", "false"}:
                return False, "bool"
            raise _PackageCorrupt("boolean cell has an invalid lexical value")
        if type_code == "e":
            return raw, "error"
        if type_code == "d":
            return _parse_iso_temporal(raw), "date"
        if type_code not in {"", "n"}:
            raise _PackageCorrupt(f"unsupported SpreadsheetML cell type: {type_code}")

        number = _parse_number(raw)
        if self.styles.is_date_style(style_idx):
            epoch = CALENDAR_MAC_1904 if self.metadata.date1904 else CALENDAR_WINDOWS_1900
            try:
                return cast(CellScalar, from_excel(number, epoch=epoch)), "date"
            except (OverflowError, ValueError) as error:
                raise _PackageCorrupt("date serial is outside the supported range") from error
        return number, "number"

    def _parse_validation(self, element: etree._Element) -> tuple[DataValidationInfo, ...]:
        sqref = attr_by_local(element, "sqref")
        if sqref is None:
            return ()
        formula1_element = child_by_local(element, "formula1")
        formula2_element = child_by_local(element, "formula2")
        formula1 = None if formula1_element is None else "".join(formula1_element.itertext())
        formula2 = None if formula2_element is None else "".join(formula2_element.itertext())
        result: list[DataValidationInfo] = []
        for area in sqref.split():
            result.append(
                DataValidationInfo(
                    rect=_checked_rect(area),
                    validation_type=attr_by_local(element, "type"),
                    operator=attr_by_local(element, "operator"),
                    formula1=formula1,
                    formula2=formula2,
                    allow_blank=parse_bool(attr_by_local(element, "allowBlank")),
                )
            )
        return tuple(result)

    def _load_tables(self, sheet_part: str, rel_ids: list[str]) -> tuple[TableInfo, ...]:
        if not rel_ids:
            return ()
        relationships = self._parse_relationships(
            _relationships_part(sheet_part), source_part=sheet_part, required=True
        )
        tables: list[TableInfo] = []
        try:
            for rel_id in rel_ids:
                relationship = relationships.get(rel_id)
                if relationship is None or relationship.part is None:
                    raise _PackageCorrupt(f"table relationship is missing: {rel_id}")
                root = parse_xml(self._read_required(relationship.part))
                name = attr_by_local(root, "name")
                display_name = attr_by_local(root, "displayName")
                ref = attr_by_local(root, "ref")
                if name is None or display_name is None or ref is None:
                    raise _PackageCorrupt("table is missing name, displayName, or ref")
                _checked_rect(ref)
                header_rows = _optional_bounded_int(
                    attr_by_local(root, "headerRowCount"), 1, 0, 1_048_576, "headerRowCount"
                )
                totals_rows = _optional_bounded_int(
                    attr_by_local(root, "totalsRowCount"), 0, 0, 1_048_576, "totalsRowCount"
                )
                column_names: list[str] = []
                for element in root.iter():
                    if isinstance(element.tag, str) and local_name(element.tag) == "tableColumn":
                        column_name = attr_by_local(element, "name")
                        if column_name is None:
                            raise _PackageCorrupt("tableColumn has no name")
                        column_names.append(column_name)
                tables.append(
                    TableInfo(
                        name=name,
                        display_name=display_name,
                        ref=ref,
                        header_rows=header_rows,
                        totals_rows=totals_rows,
                        columns=tuple(column_names),
                    )
                )
        except (zipfile.BadZipFile, EOFError, etree.XMLSyntaxError) as error:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                "Table metadata is malformed or incomplete.",
                details={"sheet_part": sheet_part},
            ) from error
        except _PackageCorrupt as error:
            raise ExcelLSPError(
                ErrorCode.CORRUPT,
                "Table metadata is malformed or incomplete.",
                details={"sheet_part": sheet_part},
            ) from error
        return tuple(tables)

    def _require_open(self) -> zipfile.ZipFile:
        if self._archive is None:
            raise RuntimeError("OOXMLParser context is not entered")
        return self._archive


def _hash_file(path: Path) -> tuple[str, bytes]:
    digest = hashlib.sha256()
    signature = b""
    with path.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            if not signature:
                signature = chunk[:8]
            digest.update(chunk)
    return digest.hexdigest(), signature


def _normalize_part_name(part: str) -> str:
    candidate = part.replace("\\", "/")
    if candidate.startswith("/"):
        candidate = candidate[1:]
    while candidate.startswith("./"):
        candidate = candidate[2:]
    normalized = posixpath.normpath(candidate)
    if normalized in {"", "."} or normalized == ".." or normalized.startswith("../"):
        raise _PackageCorrupt(f"invalid package part name: {part!r}")
    return normalized


def _resolve_target(source_part: str, target: str) -> str:
    if "://" in target or target.casefold().startswith(("mailto:", "file:")):
        raise _PackageCorrupt("internal relationship target is an external URI")
    if target.startswith("/"):
        return _normalize_part_name(target)
    return _normalize_part_name(posixpath.join(posixpath.dirname(source_part), target))


def _relationships_part(source_part: str) -> str:
    directory, filename = posixpath.split(source_part)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def _classify_sheet(
    relationship_type: str,
    part: str,
    content_type: str | None,
) -> SheetKind:
    value = " ".join((relationship_type, part, content_type or "")).casefold()
    if "chartsheet" in value:
        return "chartsheet"
    if "dialogsheet" in value or "/dialogsheets/" in value:
        return "dialog"
    if "macrosheet" in value or "/macrosheets/" in value:
        return "macro"
    return "worksheet"


def _parse_visibility(value: str | None) -> Visibility:
    if value is None or value.casefold() == "visible":
        return "visible"
    if value.casefold() == "hidden":
        return "hidden"
    if value.casefold() == "veryhidden":
        return "veryHidden"
    raise _PackageCorrupt(f"invalid sheet visibility: {value!r}")


def _classify_defined_name(
    refers_to: str,
    scope: int | None,
    sheets: tuple[SheetDescriptor, ...],
) -> tuple[DefinedNameKind, tuple[NameArea, ...]]:
    expression = refers_to.strip()
    if expression.startswith("="):
        expression = expression[1:]
    if _LAMBDA_PREFIX.match(expression):
        return "lambda", ()
    areas = _parse_name_areas(expression, scope, sheets)
    if areas:
        return ("range" if len(areas) == 1 else "multi_range"), areas
    if _is_constant_name(expression):
        return "constant", ()
    return "formula", ()


def _parse_name_areas(
    expression: str,
    scope: int | None,
    sheets: tuple[SheetDescriptor, ...],
) -> tuple[NameArea, ...]:
    if not expression:
        return ()
    tokens = _split_union(expression)
    sheet_by_name = {sheet.name.casefold(): sheet for sheet in sheets}
    areas: list[NameArea] = []
    for token in tokens:
        item = token.strip()
        delimiter = item.rfind("!")
        if delimiter < 0:
            if scope is None or not 0 <= scope < len(sheets):
                return ()
            sheet = sheets[scope]
            reference = item
        else:
            sheet_token = item[:delimiter].strip()
            reference = item[delimiter + 1 :].strip()
            if sheet_token.startswith("'") and sheet_token.endswith("'"):
                sheet_name = sheet_token[1:-1].replace("''", "'")
            else:
                sheet_name = sheet_token
            if "[" in sheet_name or "]" in sheet_name or ":" in sheet_name:
                return ()
            sheet = sheet_by_name.get(sheet_name.casefold())
            if sheet is None:
                return ()
        try:
            rect = parse_rect(reference)
        except ValueError:
            return ()
        areas.append(NameArea(sheet.name, rect))
    return tuple(areas)


def _split_union(expression: str) -> tuple[str, ...]:
    parts: list[str] = []
    start = 0
    quoted_sheet = False
    double_quoted = False
    bracket_depth = 0
    paren_depth = 0
    index = 0
    while index < len(expression):
        character = expression[index]
        if quoted_sheet:
            if character == "'":
                if index + 1 < len(expression) and expression[index + 1] == "'":
                    index += 2
                    continue
                quoted_sheet = False
        elif double_quoted:
            if character == '"':
                if index + 1 < len(expression) and expression[index + 1] == '"':
                    index += 2
                    continue
                double_quoted = False
        elif character == "'":
            quoted_sheet = True
        elif character == '"':
            double_quoted = True
        elif character == "[":
            bracket_depth += 1
        elif character == "]" and bracket_depth:
            bracket_depth -= 1
        elif character == "(":
            paren_depth += 1
        elif character == ")" and paren_depth:
            paren_depth -= 1
        elif character == "," and bracket_depth == 0 and paren_depth == 0:
            parts.append(expression[start:index])
            start = index + 1
        index += 1
    parts.append(expression[start:])
    return tuple(parts)


def _is_constant_name(expression: str) -> bool:
    if not expression:
        return True
    if _NUMERIC_LITERAL.fullmatch(expression):
        return True
    if expression.casefold() in {"true", "false"}:
        return True
    if _ERROR_LITERAL.fullmatch(expression):
        return True
    if expression.startswith('"') and expression.endswith('"'):
        return True
    return expression.startswith("{") and expression.endswith("}")


def _checked_cell_ref(ref: str) -> tuple[int, int]:
    try:
        return parse_cell_ref(ref)
    except ValueError as error:
        raise _PackageCorrupt(f"invalid cell reference: {ref!r}") from error


def _checked_rect(ref: str) -> Rect:
    try:
        return parse_rect(ref)
    except ValueError as error:
        raise _PackageCorrupt(f"invalid rectangle: {ref!r}") from error


def _bounded_int(value: str, minimum: int, maximum: int, field: str) -> int:
    try:
        parsed = int(value)
    except ValueError as error:
        raise _PackageCorrupt(f"{field} is not an integer") from error
    if not minimum <= parsed <= maximum:
        raise _PackageCorrupt(f"{field} is outside its allowed bounds")
    return parsed


def _optional_bounded_int(
    value: str | None,
    default: int,
    minimum: int,
    maximum: int,
    field: str,
) -> int:
    return default if value is None else _bounded_int(value, minimum, maximum, field)


def _optional_xml_bool(value: str | None, field: str) -> bool | None:
    if value is None:
        return None
    normalized = value.strip().casefold()
    if normalized in {"1", "true", "on"}:
        return True
    if normalized in {"0", "false", "off"}:
        return False
    raise _PackageCorrupt(f"{field} is not a valid boolean")


def _optional_nonnegative_float(value: str | None, field: str) -> float | None:
    if value is None:
        return None
    try:
        result = float(value)
    except ValueError as error:
        raise _PackageCorrupt(f"{field} is not a number") from error
    if not math.isfinite(result) or result < 0:
        raise _PackageCorrupt(f"{field} is outside its allowed bounds")
    return result


def _formula_with_equals(formula: str) -> str:
    return formula if formula.startswith("=") else f"={formula}"


def _parse_number(value: str) -> int | float:
    try:
        parsed = Decimal(value)
    except InvalidOperation as error:
        raise _PackageCorrupt("numeric cell has an invalid value") from error
    if not parsed.is_finite():
        raise _PackageCorrupt("numeric cell is not finite")
    if parsed == parsed.to_integral_value():
        return int(parsed)
    result = float(parsed)
    if not math.isfinite(result):
        raise _PackageCorrupt("numeric cell is outside the float range")
    return result


def _parse_iso_temporal(value: str) -> date | datetime | time:
    try:
        if "T" in value or " " in value:
            return datetime.fromisoformat(value)
        if ":" in value:
            return time.fromisoformat(value)
        return date.fromisoformat(value)
    except ValueError as error:
        raise _PackageCorrupt("ISO date cell has an invalid value") from error


def _has_ancestor(element: etree._Element, names: set[str]) -> bool:
    ancestor = element.getparent()
    while ancestor is not None:
        if isinstance(ancestor.tag, str) and local_name(ancestor.tag) in names:
            return True
        ancestor = ancestor.getparent()
    return False


def _is_lock_error(error: OSError) -> bool:
    return isinstance(error, PermissionError) or getattr(error, "winerror", None) in {32, 33}
