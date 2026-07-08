"""RNMC workbook preview helpers.

This module mirrors the legacy VBA import detection rules, but it does not
write catalog rows. It is used to preview what an RNMC zip import would find.
"""

from __future__ import annotations

import re
import sqlite3
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from io import BytesIO
from numbers import Real
from pathlib import PurePosixPath
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from app.services.rnmc_zip import EXCEL_SUFFIXES
from core.catalog import CatalogRow
from core.normalize import NormCode, NormUnit
from core.storage.catalog import (
    ROW_LOG_STATUS_REJECTED,
    STATUS_DUPLICATE_NAME as DB_STATUS_DUPLICATE_NAME,
    STATUS_FAILED,
    STATUS_NO_DATA,
    STATUS_SKIPPED,
    STATUS_SUCCESS,
    CatalogRowStorageItem,
    filename_is_final_for_preview,
    imported_file_exists_for_region,
    normalize_import_filename,
    record_imported_file,
    replace_catalog_rows_for_file,
    replace_import_row_logs,
)

HEADER_NAME_WORKS = "\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435 \u0440\u0430\u0431\u043e\u0442"
HEADER_NAME_SHORT = "\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435"
HEADER_UNIT = "\u0415\u0434.\u0438\u0437\u043c."
HEADER_UNIT_NO_DOT = "\u0415\u0434.\u0438\u0437\u043c"
HEADER_UNIT_LONG = "\u0415\u0434\u0438\u043d\u0438\u0446\u0430 \u0438\u0437\u043c\u0435\u0440\u0435\u043d\u0438\u044f"
HEADER_QTY_SHORT = "\u041a\u043e\u043b-\u0432\u043e"
HEADER_QTY_LONG = "\u041a\u043e\u043b\u0438\u0447\u0435\u0441\u0442\u0432\u043e"
TASK_LABEL_FULL = "\u2116 \u0437\u0430\u0434\u0430\u0447\u0438 1\u0424"
TASK_LABEL_SHORT = "\u2116 \u0437\u0430\u0434\u0430\u0447\u0438"

CODE_HEADER_PATTERNS = (
    "\u0413\u042d\u0421\u041d/\u0424\u0415\u0420/\u041f\u0435\u0440\u0435\u0447\u0435\u043d\u044c",
    "\u041f\u0435\u0440\u0435\u0447\u0435\u043d\u044c \u0413\u042d\u0421\u041d",
    "\u041f\u0435\u0440\u0435\u0447\u0435\u043d\u044c",
    "\u041e\u0431\u043e\u0441\u043d\u043e\u0432\u0430\u043d\u0438\u0435",
    "\u0428\u0438\u0444\u0440",
    "\u041a\u043e\u0434",
)
PRICE_HEADER_PATTERNS = (
    "\u0426\u0435\u043d\u0430 \u0435\u0434\u0438\u043d\u0438\u0446\u044b \u0440\u0430\u0431\u043e\u0442",
    "\u0426\u0435\u043d\u0430 \u0435\u0434\u0438\u043d\u0438\u0446\u044b",
    "\u0426\u0435\u043d\u0430 \u0437\u0430 \u0435\u0434",
    "\u0421\u0442\u043e\u0438\u043c\u043e\u0441\u0442\u044c \u0435\u0434\u0438\u043d\u0438\u0446\u044b",
    "\u0421\u0442\u043e\u0438\u043c\u043e\u0441\u0442\u044c \u0437\u0430 \u0435\u0434",
    "\u0411\u0430\u0437\u043e\u0432\u0430\u044f \u0446\u0435\u043d\u0430",
    "\u0426\u0435\u043d\u0430 \u0437\u0430 1",
    "\u0426\u0435\u043d\u0430",
)
ADDED_DATE_HEADER_PATTERNS = (
    "CatalogAddedDate",
    "\u0414\u0430\u0442\u0430 \u0434\u043e\u0431\u0430\u0432\u043b\u0435\u043d\u0438\u044f",
    "\u0414\u0430\u0442\u0430 \u0434\u043e\u0431",
)

LSR_QUARTER_LABEL_PATTERNS = (
    "\u0413\u043e\u0434 \u041a\u0432\u0430\u0440\u0442\u0430\u043b \u041b\u0421\u0420",
    "\u0413\u043e\u0434/\u043a\u0432\u0430\u0440\u0442\u0430\u043b \u041b\u0421\u0420",
    "\u0413\u043e\u0434 \u0438 \u043a\u0432\u0430\u0440\u0442\u0430\u043b \u041b\u0421\u0420",
    "\u041a\u0432\u0430\u0440\u0442\u0430\u043b \u041b\u0421\u0420",
    "\u041b\u0421\u0420",
)
PLANNED_START_LABEL_PATTERNS = (
    "\u041f\u043b\u0430\u043d\u0438\u0440\u0443\u043c\u044b\u0439 \u0441\u0440\u043e\u043a \u043d\u0430\u0447\u0430\u043b\u0430 \u0440\u0430\u0431\u043e\u0442",
    "\u041f\u043b\u0430\u043d\u0438\u0440\u0443\u0435\u043c\u044b\u0439 \u0441\u0440\u043e\u043a \u043d\u0430\u0447\u0430\u043b\u0430 \u0440\u0430\u0431\u043e\u0442",
    "\u041f\u043b\u0430\u043d\u043e\u0432\u044b\u0439 \u0441\u0440\u043e\u043a \u043d\u0430\u0447\u0430\u043b\u0430 \u0440\u0430\u0431\u043e\u0442",
    "\u0414\u0430\u0442\u0430 \u043d\u0430\u0447\u0430\u043b\u0430 \u0440\u0430\u0431\u043e\u0442",
    "\u043d\u0430\u0447\u0430\u043b\u043e \u0440\u0430\u0431\u043e\u0442",
)
PLANNED_FINISH_LABEL_PATTERNS = (
    "\u041f\u043b\u0430\u043d\u0438\u0440\u0443\u0435\u043c\u044b\u0439 \u0441\u0440\u043e\u043a \u043e\u043a\u043e\u043d\u0447\u0430\u043d\u0438\u044f \u0440\u0430\u0431\u043e\u0442",
    "\u041f\u043b\u0430\u043d\u0438\u0440\u0443\u043c\u044b\u0439 \u0441\u0440\u043e\u043a \u043e\u043a\u043e\u043d\u0447\u0430\u043d\u0438\u044f \u0440\u0430\u0431\u043e\u0442",
    "\u041f\u043b\u0430\u043d\u043e\u0432\u044b\u0439 \u0441\u0440\u043e\u043a \u043e\u043a\u043e\u043d\u0447\u0430\u043d\u0438\u044f \u0440\u0430\u0431\u043e\u0442",
    "\u0414\u0430\u0442\u0430 \u043e\u043a\u043e\u043d\u0447\u0430\u043d\u0438\u044f \u0440\u0430\u0431\u043e\u0442",
    "\u043e\u043a\u043e\u043d\u0447\u0430\u043d\u0438\u0435 \u0440\u0430\u0431\u043e\u0442",
)

STATUS_PREVIEW_OK = "preview_ok"
STATUS_NO_TABLE = "no_table"
STATUS_NO_ROWS = "no_rows"
STATUS_SKIPPED_PROCESSED = "skipped_processed"
STATUS_DUPLICATE_NAME = "duplicate_name"
STATUS_UNSUPPORTED_FORMAT = "unsupported_format"
STATUS_PARSE_ERROR = "parse_error"

SUPPORTED_PREVIEW_SUFFIXES = frozenset({".xlsx", ".xlsm"})
EXCEL_DATE_BASE = date(1899, 12, 30)
VALUE_WITH_VAT_DIVISOR = 1.2


@dataclass(frozen=True)
class RnmcWorkbookMetadata:
    lsr_quarter: str = ""
    planned_start: str = ""
    planned_finish: str = ""


@dataclass(frozen=True)
class RnmcRowSample:
    row_number: int
    work_name: str
    unit: str
    quantity: str


@dataclass(frozen=True)
class RnmcWorkbookPreview:
    archive_path: str
    filename: str
    region_folder: str
    status: str
    reason: str
    sheet_name: str
    header_row: int
    task_number: str
    lsr_quarter: str
    planned_start: str
    planned_finish: str
    rows_ok: int
    rows_rejected: int
    sample_rows: list[RnmcRowSample]


@dataclass(frozen=True)
class RnmcZipRowPreviewResult:
    entries: list[RnmcWorkbookPreview]
    ignored_files: int

    @property
    def total_excel_files(self) -> int:
        return len(self.entries)

    @property
    def preview_ok_count(self) -> int:
        return _count_status(self.entries, STATUS_PREVIEW_OK)

    @property
    def skipped_processed_count(self) -> int:
        return _count_status(self.entries, STATUS_SKIPPED_PROCESSED)

    @property
    def duplicate_name_count(self) -> int:
        return _count_status(self.entries, STATUS_DUPLICATE_NAME)

    @property
    def no_table_count(self) -> int:
        return _count_status(self.entries, STATUS_NO_TABLE)

    @property
    def no_rows_count(self) -> int:
        return _count_status(self.entries, STATUS_NO_ROWS)

    @property
    def unsupported_format_count(self) -> int:
        return _count_status(self.entries, STATUS_UNSUPPORTED_FORMAT)

    @property
    def parse_error_count(self) -> int:
        return _count_status(self.entries, STATUS_PARSE_ERROR)

    @property
    def rows_ok_total(self) -> int:
        return sum(entry.rows_ok for entry in self.entries)

    @property
    def rows_rejected_total(self) -> int:
        return sum(entry.rows_rejected for entry in self.entries)


@dataclass(frozen=True)
class RnmcRejectedRow:
    row_number: int
    reason: str


@dataclass(frozen=True)
class RnmcCatalogRowCandidate:
    row_number: int
    catalog_row: CatalogRow


@dataclass(frozen=True)
class RnmcZipCatalogImportEntry:
    archive_path: str
    filename: str
    region_folder: str
    status: str
    reason: str
    sheet_name: str
    header_row: int
    task_number: str
    lsr_quarter: str
    planned_start: str
    planned_finish: str
    rows_imported: int
    rows_rejected: int


@dataclass(frozen=True)
class RnmcZipCatalogImportResult:
    entries: list[RnmcZipCatalogImportEntry]
    ignored_files: int

    @property
    def total_excel_files(self) -> int:
        return len(self.entries)

    @property
    def success_count(self) -> int:
        return _count_import_status(self.entries, STATUS_SUCCESS)

    @property
    def no_data_count(self) -> int:
        return _count_import_status(self.entries, STATUS_NO_DATA)

    @property
    def failed_count(self) -> int:
        return _count_import_status(self.entries, STATUS_FAILED)

    @property
    def skipped_count(self) -> int:
        return _count_import_status(self.entries, STATUS_SKIPPED)

    @property
    def duplicate_name_count(self) -> int:
        return _count_import_status(self.entries, DB_STATUS_DUPLICATE_NAME)

    @property
    def rows_imported_total(self) -> int:
        return sum(entry.rows_imported for entry in self.entries)

    @property
    def rows_rejected_total(self) -> int:
        return sum(entry.rows_rejected for entry in self.entries)


def analyze_rnmc_zip_row_preview(
    connection: sqlite3.Connection,
    zip_path: str,
    *,
    region_override: str = "",
) -> RnmcZipRowPreviewResult:
    """Preview RNMC workbook rows inside a zip archive without database writes."""
    manual_region = _text(region_override)
    entries: list[RnmcWorkbookPreview] = []
    ignored_files = 0
    seen_keys: set[str] = set()

    try:
        with ZipFile(zip_path) as archive:
            for info in archive.infolist():
                if info.is_dir():
                    continue
                path = _normalize_archive_path(info.filename)
                if not _is_supported_excel_path(path):
                    ignored_files += 1
                    continue
                filename = PurePosixPath(path).name
                key = normalize_import_filename(filename)
                region = manual_region or _region_from_archive_path(path)

                if key in seen_keys:
                    entries.append(
                        _empty_preview(
                            path,
                            filename,
                            region,
                            STATUS_DUPLICATE_NAME,
                            "Duplicate filename inside uploaded zip",
                        )
                    )
                elif filename_is_final_for_preview(connection, filename):
                    entries.append(
                        _empty_preview(
                            path,
                            filename,
                            region,
                            STATUS_SKIPPED_PROCESSED,
                            "Filename already has a final imported_files status",
                        )
                    )
                elif PurePosixPath(path).suffix.casefold() not in SUPPORTED_PREVIEW_SUFFIXES:
                    entries.append(
                        _empty_preview(
                            path,
                            filename,
                            region,
                            STATUS_UNSUPPORTED_FORMAT,
                            "Preview supports .xlsx and .xlsm; legacy .xls will need a separate reader",
                        )
                    )
                else:
                    try:
                        data = archive.read(info)
                        entries.append(_preview_workbook_bytes(data, path, filename, region))
                    except Exception as exc:  # pragma: no cover - defensive UI boundary
                        entries.append(
                            _empty_preview(
                                path,
                                filename,
                                region,
                                STATUS_PARSE_ERROR,
                                str(exc),
                            )
                        )
                seen_keys.add(key)
    except BadZipFile as exc:
        raise ValueError("uploaded file is not a valid zip archive") from exc

    return RnmcZipRowPreviewResult(entries=entries, ignored_files=ignored_files)


def import_rnmc_zip_catalog_rows(
    connection: sqlite3.Connection,
    zip_path: str,
    *,
    region_override: str = "",
) -> RnmcZipCatalogImportResult:
    """Import valid RNMC rows from a zip archive into catalog_items."""
    manual_region = _text(region_override)
    entries: list[RnmcZipCatalogImportEntry] = []
    ignored_files = 0
    seen_keys: set[str] = set()

    try:
        with ZipFile(zip_path) as archive:
            for info in archive.infolist():
                if info.is_dir():
                    continue
                path = _normalize_archive_path(info.filename)
                if not _is_supported_excel_path(path):
                    ignored_files += 1
                    continue
                filename = PurePosixPath(path).name
                key = normalize_import_filename(filename)
                region = manual_region or _region_from_archive_path(path)

                if key in seen_keys:
                    entry = _record_non_imported_file(
                        connection,
                        path,
                        filename,
                        region,
                        DB_STATUS_DUPLICATE_NAME,
                        "Duplicate filename inside uploaded zip",
                    )
                elif filename_is_final_for_preview(connection, filename):
                    if imported_file_exists_for_region(
                        connection,
                        region_folder=region,
                        filename=filename,
                    ):
                        entry = _import_entry(
                            path,
                            filename,
                            region,
                            STATUS_SKIPPED,
                            "Filename already has a final imported_files status",
                        )
                    else:
                        entry = _record_non_imported_file(
                            connection,
                            path,
                            filename,
                            region,
                            STATUS_SKIPPED,
                            "Skipped because filename already exists in imported_files",
                        )
                elif PurePosixPath(path).suffix.casefold() not in SUPPORTED_PREVIEW_SUFFIXES:
                    entry = _record_non_imported_file(
                        connection,
                        path,
                        filename,
                        region,
                        STATUS_FAILED,
                        "Import supports .xlsx and .xlsm; legacy .xls will need a separate reader",
                    )
                else:
                    try:
                        data = archive.read(info)
                        entry = _import_workbook_bytes(
                            connection,
                            data,
                            path,
                            filename,
                            region,
                        )
                    except Exception as exc:  # pragma: no cover - defensive UI boundary
                        entry = _record_non_imported_file(
                            connection,
                            path,
                            filename,
                            region,
                            STATUS_FAILED,
                            str(exc),
                        )
                seen_keys.add(key)
                entries.append(entry)
    except BadZipFile as exc:
        raise ValueError("uploaded file is not a valid zip archive") from exc

    connection.commit()
    return RnmcZipCatalogImportResult(entries=entries, ignored_files=ignored_files)


def _import_workbook_bytes(
    connection: sqlite3.Connection,
    data: bytes,
    archive_path: str,
    filename: str,
    region_folder: str,
) -> RnmcZipCatalogImportEntry:
    workbook = load_workbook(BytesIO(data), read_only=True, data_only=True)
    try:
        metadata = _extract_workbook_metadata(workbook)
        task_number = _extract_task_number(workbook)
        for sheet in workbook.worksheets:
            header = _find_header_row(sheet)
            if header is None:
                continue
            rows, rejected_rows = _extract_catalog_row_candidates(
                sheet,
                header,
                task_number,
                region_folder,
                filename,
            )
            rejected = len(rejected_rows)
            if not rows:
                file_id = record_imported_file(
                    connection,
                    region_folder=region_folder,
                    filename=filename,
                    status=STATUS_NO_DATA,
                    rows_ok=0,
                    rows_rejected=rejected,
                    failure_reason="Header found, but no valid catalog rows",
                    task_number=task_number,
                    lsr_quarter=metadata.lsr_quarter,
                    planned_start=metadata.planned_start,
                    planned_finish=metadata.planned_finish,
                )
                replace_import_row_logs(
                    connection,
                    file_id,
                    [
                        (row.row_number, ROW_LOG_STATUS_REJECTED, row.reason)
                        for row in rejected_rows
                    ],
                )
                return _import_entry(
                    archive_path,
                    filename,
                    region_folder,
                    STATUS_NO_DATA,
                    "Header found, but no valid catalog rows",
                    sheet_name=str(sheet.title),
                    header_row=header.row_number,
                    task_number=task_number,
                    lsr_quarter=metadata.lsr_quarter,
                    planned_start=metadata.planned_start,
                    planned_finish=metadata.planned_finish,
                    rows_rejected=rejected,
                )

            result = replace_catalog_rows_for_file(
                connection,
                [
                    CatalogRowStorageItem(
                        catalog_row=row.catalog_row,
                        source_region_folder=region_folder,
                        source_filename=filename,
                        source_row_number=row.row_number,
                    )
                    for row in rows
                ],
                region_folder=region_folder,
                filename=filename,
            )
            file_id = record_imported_file(
                connection,
                region_folder=region_folder,
                filename=filename,
                status=STATUS_SUCCESS,
                rows_ok=result.rows_imported,
                rows_rejected=rejected + result.rows_skipped,
                failure_reason="",
                task_number=task_number,
                lsr_quarter=metadata.lsr_quarter,
                planned_start=metadata.planned_start,
                planned_finish=metadata.planned_finish,
            )
            replace_import_row_logs(
                connection,
                file_id,
                [
                    (row.row_number, ROW_LOG_STATUS_REJECTED, row.reason)
                    for row in rejected_rows
                ],
            )
            return _import_entry(
                archive_path,
                filename,
                region_folder,
                STATUS_SUCCESS,
                "Catalog rows imported",
                sheet_name=str(sheet.title),
                header_row=header.row_number,
                task_number=task_number,
                lsr_quarter=metadata.lsr_quarter,
                planned_start=metadata.planned_start,
                planned_finish=metadata.planned_finish,
                rows_imported=result.rows_imported,
                rows_rejected=rejected + result.rows_skipped,
            )

        file_id = record_imported_file(
            connection,
            region_folder=region_folder,
            filename=filename,
            status=STATUS_NO_DATA,
            rows_ok=0,
            rows_rejected=0,
            failure_reason="Required headers were not found",
            task_number=task_number,
            lsr_quarter=metadata.lsr_quarter,
            planned_start=metadata.planned_start,
            planned_finish=metadata.planned_finish,
        )
        replace_import_row_logs(connection, file_id, [])
        return _import_entry(
            archive_path,
            filename,
            region_folder,
            STATUS_NO_DATA,
            "Required headers were not found",
            task_number=task_number,
            lsr_quarter=metadata.lsr_quarter,
            planned_start=metadata.planned_start,
            planned_finish=metadata.planned_finish,
        )
    finally:
        workbook.close()


def _extract_catalog_row_candidates(
    sheet: Worksheet,
    header: _HeaderMatch,
    task_number: str,
    region_folder: str,
    filename: str,
) -> tuple[list[RnmcCatalogRowCandidate], list[RnmcRejectedRow]]:
    num_col = _find_numbering_col(header.header_map) or 1
    code_col = _find_pattern_col(header.header_map, CODE_HEADER_PATTERNS)
    value_columns = _detect_value_columns(header.header_map)
    date_col = _find_pattern_col(header.header_map, ADDED_DATE_HEADER_PATTERNS)
    max_row = int(sheet.max_row or 0)
    rows: list[RnmcCatalogRowCandidate] = []
    rejected: list[RnmcRejectedRow] = []
    started = False
    blank_streak = 0

    for row_number in range(header.row_number + 1, max_row + 1):
        num_value = sheet.cell(row_number, num_col).value
        name_value = sheet.cell(row_number, header.name_col).value
        unit_value = sheet.cell(row_number, header.unit_col).value
        qty_value = sheet.cell(row_number, header.qty_col).value

        if not started:
            if _is_blank(num_value) and _is_blank(unit_value) and _is_blank(qty_value):
                continue
            started = True

        is_end_blank = (
            _is_blank(num_value)
            and _is_blank(name_value)
            and _is_blank(unit_value)
            and _is_blank(qty_value)
        )
        if is_end_blank:
            blank_streak += 1
        else:
            blank_streak = 0
        if blank_streak >= 3:
            break
        if is_end_blank:
            continue

        if _is_blank(unit_value) and _is_blank(qty_value):
            rejected.append(RnmcRejectedRow(row_number, "missing_unit_and_quantity"))
            continue

        price_value = _cell_value(sheet, row_number, value_columns.unit_price.column)
        parsed_price = _parse_positive_number(price_value)
        if parsed_price is not None:
            parsed_price = parsed_price / value_columns.unit_price.divisor

        total_price = _parse_optional_value(
            _cell_value(sheet, row_number, value_columns.total_price.column),
            divisor=value_columns.total_price.divisor,
        )
        catalog_row = CatalogRow(
            task_id=task_number,
            price=parsed_price if parsed_price is not None else price_value,
            code=_cell_value(sheet, row_number, code_col),
            unit=unit_value,
            work_name=name_value,
            region=region_folder,
            added_date=_cell_value(sheet, row_number, date_col),
            total_price=total_price,
            labor_unit=_parse_optional_value(
                _cell_value(sheet, row_number, value_columns.labor_unit_col)
            ),
            labor_total=_parse_optional_value(
                _cell_value(sheet, row_number, value_columns.labor_total_col)
            ),
            machine_labor_unit=_parse_optional_value(
                _cell_value(sheet, row_number, value_columns.machine_labor_unit_col)
            ),
            machine_labor_total=_parse_optional_value(
                _cell_value(sheet, row_number, value_columns.machine_labor_total_col)
            ),
        )
        issue = _catalog_row_issue(catalog_row)
        if issue != "":
            rejected.append(RnmcRejectedRow(row_number, issue))
            continue
        rows.append(RnmcCatalogRowCandidate(row_number=row_number, catalog_row=catalog_row))

    return rows, rejected


def _record_non_imported_file(
    connection: sqlite3.Connection,
    archive_path: str,
    filename: str,
    region_folder: str,
    status: str,
    reason: str,
) -> RnmcZipCatalogImportEntry:
    file_id = record_imported_file(
        connection,
        region_folder=region_folder,
        filename=filename,
        status=status,
        rows_ok=0,
        rows_rejected=0,
        failure_reason=reason,
    )
    replace_import_row_logs(connection, file_id, [])
    return _import_entry(archive_path, filename, region_folder, status, reason)


def _import_entry(
    archive_path: str,
    filename: str,
    region_folder: str,
    status: str,
    reason: str,
    *,
    sheet_name: str = "",
    header_row: int = 0,
    task_number: str = "",
    lsr_quarter: str = "",
    planned_start: str = "",
    planned_finish: str = "",
    rows_imported: int = 0,
    rows_rejected: int = 0,
) -> RnmcZipCatalogImportEntry:
    return RnmcZipCatalogImportEntry(
        archive_path=archive_path,
        filename=filename,
        region_folder=region_folder,
        status=status,
        reason=reason,
        sheet_name=sheet_name,
        header_row=header_row,
        task_number=task_number,
        lsr_quarter=lsr_quarter,
        planned_start=planned_start,
        planned_finish=planned_finish,
        rows_imported=rows_imported,
        rows_rejected=rows_rejected,
    )


def _preview_workbook_bytes(
    data: bytes,
    archive_path: str,
    filename: str,
    region_folder: str,
) -> RnmcWorkbookPreview:
    workbook = load_workbook(BytesIO(data), read_only=True, data_only=True)
    try:
        metadata = _extract_workbook_metadata(workbook)
        task_number = _extract_task_number(workbook)
        for sheet in workbook.worksheets:
            header = _find_header_row(sheet)
            if header is None:
                continue
            rows_ok, rows_rejected, samples = _preview_table_rows(sheet, header)
            status = STATUS_PREVIEW_OK if rows_ok > 0 else STATUS_NO_ROWS
            reason = "Rows found" if rows_ok > 0 else "Header found, but no accepted rows"
            return RnmcWorkbookPreview(
                archive_path=archive_path,
                filename=filename,
                region_folder=region_folder,
                status=status,
                reason=reason,
                sheet_name=str(sheet.title),
                header_row=header.row_number,
                task_number=task_number,
                lsr_quarter=metadata.lsr_quarter,
                planned_start=metadata.planned_start,
                planned_finish=metadata.planned_finish,
                rows_ok=rows_ok,
                rows_rejected=rows_rejected,
                sample_rows=samples,
            )

        return _empty_preview(
            archive_path,
            filename,
            region_folder,
            STATUS_NO_TABLE,
            "Required headers were not found",
            task_number=task_number,
            lsr_quarter=metadata.lsr_quarter,
            planned_start=metadata.planned_start,
            planned_finish=metadata.planned_finish,
        )
    finally:
        workbook.close()


@dataclass(frozen=True)
class _HeaderMatch:
    row_number: int
    name_col: int
    unit_col: int
    qty_col: int
    header_map: dict[str, int]


@dataclass(frozen=True)
class _ValueColumn:
    column: int
    divisor: float = 1.0


@dataclass(frozen=True)
class _HeaderValueColumns:
    unit_price: _ValueColumn
    total_price: _ValueColumn
    labor_unit_col: int
    labor_total_col: int
    machine_labor_unit_col: int
    machine_labor_total_col: int


def _find_header_row(sheet: Worksheet) -> _HeaderMatch | None:
    max_row = min(int(sheet.max_row or 0), 400)
    max_col = min(int(sheet.max_column or 0), 150)
    if max_row <= 0 or max_col <= 0:
        return None

    rows = sheet.iter_rows(
        min_row=1,
        max_row=max_row,
        min_col=1,
        max_col=max_col,
        values_only=True,
    )
    for row_number, row in enumerate(rows, start=1):
        name_col = 0
        unit_col = 0
        qty_col = 0
        header_map: dict[str, int] = {}
        for index, value in enumerate(row, start=1):
            for normalized in _header_keys(value):
                header_map.setdefault(normalized, index)
                if _is_name_header(normalized):
                    name_col = index
                if _is_unit_header(normalized):
                    unit_col = index
                if _is_qty_header(normalized):
                    qty_col = index
        if name_col and unit_col and qty_col:
            return _HeaderMatch(
                row_number=row_number,
                name_col=name_col,
                unit_col=unit_col,
                qty_col=qty_col,
                header_map=header_map,
            )
    return None


def _preview_table_rows(
    sheet: Worksheet,
    header: _HeaderMatch,
    *,
    sample_limit: int = 5,
) -> tuple[int, int, list[RnmcRowSample]]:
    num_col = _find_numbering_col(header.header_map) or 1
    max_row = int(sheet.max_row or 0)
    rows_ok = 0
    rows_rejected = 0
    samples: list[RnmcRowSample] = []
    started = False
    blank_streak = 0

    for row_number in range(header.row_number + 1, max_row + 1):
        num_value = sheet.cell(row_number, num_col).value
        name_value = sheet.cell(row_number, header.name_col).value
        unit_value = sheet.cell(row_number, header.unit_col).value
        qty_value = sheet.cell(row_number, header.qty_col).value

        if not started:
            if _is_blank(num_value) and _is_blank(unit_value) and _is_blank(qty_value):
                continue
            started = True

        is_end_blank = (
            _is_blank(num_value)
            and _is_blank(name_value)
            and _is_blank(unit_value)
            and _is_blank(qty_value)
        )
        if is_end_blank:
            blank_streak += 1
        else:
            blank_streak = 0
        if blank_streak >= 3:
            break
        if is_end_blank:
            continue

        if _is_blank(unit_value) and _is_blank(qty_value):
            rows_rejected += 1
            continue

        rows_ok += 1
        if len(samples) < sample_limit:
            samples.append(
                RnmcRowSample(
                    row_number=row_number,
                    work_name=_text(name_value),
                    unit=_text(unit_value),
                    quantity=_text(qty_value),
                )
            )
    return rows_ok, rows_rejected, samples


def _extract_workbook_metadata(workbook) -> RnmcWorkbookMetadata:
    lsr_quarter = ""
    planned_start = ""
    planned_finish = ""
    for sheet in workbook.worksheets:
        if lsr_quarter == "":
            lsr_quarter = _find_metadata_value(
                sheet,
                LSR_QUARTER_LABEL_PATTERNS,
                value_kind="quarter",
            )
        if planned_start == "":
            planned_start = _find_metadata_value(
                sheet,
                PLANNED_START_LABEL_PATTERNS,
                value_kind="date",
            )
        if planned_finish == "":
            planned_finish = _find_metadata_value(
                sheet,
                PLANNED_FINISH_LABEL_PATTERNS,
                value_kind="date",
            )

        context = _find_metadata_context(sheet)
        if lsr_quarter == "":
            lsr_quarter = context.lsr_quarter
        if planned_start == "":
            planned_start = context.planned_start
        if planned_finish == "":
            planned_finish = context.planned_finish
        if lsr_quarter and planned_start and planned_finish:
            break
    return RnmcWorkbookMetadata(
        lsr_quarter=lsr_quarter,
        planned_start=planned_start,
        planned_finish=planned_finish,
    )


def _find_metadata_value(
    sheet: Worksheet,
    label_patterns: tuple[str, ...],
    *,
    value_kind: str,
) -> str:
    normalized_patterns = tuple(_normalize_header(pattern) for pattern in label_patterns)
    max_row = min(int(sheet.max_row or 0), 120)
    max_col = min(int(sheet.max_column or 0), 60)
    for row_number in range(1, max_row + 1):
        for col_number in range(1, max_col + 1):
            value = sheet.cell(row_number, col_number).value
            normalized = _normalize_header(value)
            if normalized == "" or not _metadata_label_matches(normalized, normalized_patterns):
                continue

            inline = _metadata_inline_value(value)
            if inline:
                formatted = _format_metadata_value(inline, value_kind=value_kind)
                if formatted:
                    return formatted

            for candidate in _metadata_neighbor_values(sheet, row_number, col_number, max_row, max_col):
                formatted = _format_metadata_value(candidate, value_kind=value_kind)
                if formatted:
                    return formatted
    return ""


def _metadata_label_matches(normalized: str, normalized_patterns: tuple[str, ...]) -> bool:
    for pattern in normalized_patterns:
        if pattern == "":
            continue
        if pattern in normalized:
            return True
        if len(normalized) >= 8 and normalized in pattern:
            return True
    return False


def _metadata_inline_value(value: object) -> str:
    text = _text(value)
    if text == "":
        return ""
    for separator in (":", "—", "-", "="):
        if separator in text:
            tail = text.split(separator, 1)[1].strip()
            if tail:
                return tail
    return ""


def _metadata_neighbor_values(
    sheet: Worksheet,
    row_number: int,
    col_number: int,
    max_row: int,
    max_col: int,
) -> list[object]:
    values: list[object] = []
    for offset in range(1, 5):
        column = col_number + offset
        if column <= max_col:
            values.append(sheet.cell(row_number, column).value)
    if row_number + 1 <= max_row:
        values.append(sheet.cell(row_number + 1, col_number).value)
        for offset in range(1, 3):
            column = col_number + offset
            if column <= max_col:
                values.append(sheet.cell(row_number + 1, column).value)
    return values


def _format_metadata_value(value: object, *, value_kind: str) -> str:
    if value_kind == "date":
        return _format_date_metadata(value)
    if value_kind == "quarter":
        return _format_quarter_metadata(value)
    return _text(value)


def _format_date_metadata(value: object) -> str:
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    serial_date = _excel_serial_date(value)
    if serial_date:
        return serial_date

    text = _text(value)
    if text == "":
        return ""
    normalized = text.replace("\u00a0", " ").strip()
    for pattern in (
        r"(?P<day>\d{1,2})[./-](?P<month>\d{1,2})[./-](?P<year>\d{4})",
        r"(?P<year>\d{4})[./-](?P<month>\d{1,2})[./-](?P<day>\d{1,2})",
    ):
        match = re.search(pattern, normalized)
        if not match:
            continue
        try:
            parsed = date(
                int(match.group("year")),
                int(match.group("month")),
                int(match.group("day")),
            )
        except ValueError:
            continue
        return parsed.isoformat()
    start, finish = _parse_month_period_context(normalized)
    return start or finish


def _format_quarter_metadata(value: object) -> str:
    text = _text(value)
    if text == "":
        return ""
    parsed = _parse_quarter_context(text)
    if parsed:
        return parsed
    normalized = text.replace("\u00a0", " ").strip()
    return normalized


def _find_metadata_context(sheet: Worksheet) -> RnmcWorkbookMetadata:
    lsr_quarter = ""
    planned_start = ""
    planned_finish = ""
    max_row = min(int(sheet.max_row or 0), 180)
    max_col = min(int(sheet.max_column or 0), 80)
    for row_number in range(1, max_row + 1):
        values = [sheet.cell(row_number, col_number).value for col_number in range(1, max_col + 1)]
        text = " ".join(_text(value) for value in values if _text(value))
        if text == "":
            continue
        if lsr_quarter == "":
            lsr_quarter = _parse_quarter_context(text)
        if planned_start == "" or planned_finish == "":
            start, finish = _parse_month_period_context(text)
            planned_start = planned_start or start
            planned_finish = planned_finish or finish
        if lsr_quarter and planned_start and planned_finish:
            break
    return RnmcWorkbookMetadata(
        lsr_quarter=lsr_quarter,
        planned_start=planned_start,
        planned_finish=planned_finish,
    )


def _parse_quarter_context(value: object) -> str:
    text = _text(value)
    if text == "":
        return ""
    folded = text.replace("\u00a0", " ").casefold()
    year_pattern = "(?P<year>(?:20|19)?\\d{2})"
    roman_pattern = "(?P<roman>iv|iii|ii|i)"
    number_pattern = "(?P<number>[1-4])"
    quarter_word = "(?:\\u043a\\u0432\\.?|\\u043a\\u0432\\u0430\\u0440\\u0442\\u0430\\u043b\\w*)"
    patterns = (
        roman_pattern + "\\s*" + quarter_word + "[^0-9]{0,20}" + year_pattern,
        number_pattern + "\\s*" + quarter_word + "[^0-9]{0,20}" + year_pattern,
        quarter_word + "\\s*" + number_pattern + "[^0-9]{0,20}" + year_pattern,
    )
    for pattern in patterns:
        match = re.search(pattern, folded)
        if match is None:
            continue
        year = _normalize_year(match.group("year"))
        quarter = match.groupdict().get("number") or _roman_quarter(match.groupdict().get("roman", ""))
        if quarter:
            return f"{year} Q{quarter}"
    return ""


def _roman_quarter(value: str) -> str:
    return {"i": "1", "ii": "2", "iii": "3", "iv": "4"}.get(value.casefold(), "")


def _normalize_year(value: str) -> str:
    text = _text(value)
    if text == "":
        return ""
    if len(text) == 2:
        return f"20{text}"
    return text


def _parse_month_period_context(value: object) -> tuple[str, str]:
    text = _text(value)
    if text == "":
        return "", ""
    folded = text.replace("\u00a0", " ").casefold()
    tokens = _month_tokens(folded)
    if not tokens:
        return "", ""
    years = [_normalize_year(item) for item in re.findall(r"(?:20|19)?\d{2}", folded)]
    if not _looks_like_period_text(folded) and not years:
        return "", ""
    if len(tokens) == 1:
        month, year = tokens[0]
        year = year or (years[0] if years else "")
        if year:
            value = _month_date(year, month)
            return value, value
        return "", ""
    start_month, start_year = tokens[0]
    finish_month, finish_year = tokens[1]
    if start_year == "" and finish_year:
        start_year = finish_year
    if finish_year == "" and start_year:
        finish_year = start_year
    if start_year == "" and finish_year == "" and years:
        start_year = years[0]
        finish_year = years[-1]
    if start_year and finish_year:
        return _month_date(start_year, start_month), _month_date(finish_year, finish_month)
    return "", ""


def _looks_like_period_text(value: str) -> bool:
    return any(
        token in value
        for token in (
            "\u043f\u0435\u0440\u0438\u043e\u0434",
            "\u043d\u0430\u0447\u0430\u043b",
            "\u043e\u043a\u043e\u043d\u0447",
            "\u0440\u0430\u0431\u043e\u0442",
        )
    ) or " \u043f\u043e " in value or "-" in value or "\u2013" in value or "\u2014" in value


def _month_tokens(value: str) -> list[tuple[int, str]]:
    month_pattern = (
        "(?P<month>"
        "\\u044f\\u043d\\u0432\\u0430\\u0440\\w*|"
        "\\u0444\\u0435\\u0432\\u0440\\u0430\\u043b\\w*|"
        "\\u043c\\u0430\\u0440\\u0442\\w*|"
        "\\u0430\\u043f\\u0440\\u0435\\u043b\\w*|"
        "\\u043c\\u0430(?:\\u0439|\\u044f|\\u0435|\\u044e|\\u0435\\u043c)?|"
        "\\u0438\\u044e\\u043d\\w*|"
        "\\u0438\\u044e\\u043b\\w*|"
        "\\u0430\\u0432\\u0433\\u0443\\u0441\\u0442\\w*|"
        "\\u0441\\u0435\\u043d\\u0442\\u044f\\u0431\\u0440\\w*|"
        "\\u043e\\u043a\\u0442\\u044f\\u0431\\u0440\\w*|"
        "\\u043d\\u043e\\u044f\\u0431\\u0440\\w*|"
        "\\u0434\\u0435\\u043a\\u0430\\u0431\\u0440\\w*)"
        "\\s*(?P<year>(?:20|19)?\\d{2})?"
    )
    tokens: list[tuple[int, str]] = []
    for match in re.finditer(month_pattern, value):
        month = _month_number(match.group("month"))
        if month:
            tokens.append((month, _normalize_year(match.group("year") or "")))
    return tokens


def _month_number(value: str) -> int:
    folded = value.casefold()
    prefixes = (
        ("\u044f\u043d\u0432\u0430\u0440", 1),
        ("\u0444\u0435\u0432\u0440\u0430\u043b", 2),
        ("\u043c\u0430\u0440\u0442", 3),
        ("\u0430\u043f\u0440\u0435\u043b", 4),
        ("\u043c\u0430", 5),
        ("\u0438\u044e\u043d", 6),
        ("\u0438\u044e\u043b", 7),
        ("\u0430\u0432\u0433\u0443\u0441\u0442", 8),
        ("\u0441\u0435\u043d\u0442\u044f\u0431\u0440", 9),
        ("\u043e\u043a\u0442\u044f\u0431\u0440", 10),
        ("\u043d\u043e\u044f\u0431\u0440", 11),
        ("\u0434\u0435\u043a\u0430\u0431\u0440", 12),
    )
    for prefix, month in prefixes:
        if folded.startswith(prefix):
            return month
    return 0


def _month_date(year: str, month: int) -> str:
    return date(int(year), int(month), 1).isoformat()


def _excel_serial_date(value: object) -> str:
    if not isinstance(value, Real) or isinstance(value, bool):
        return ""
    serial = float(value)
    if serial < 20000 or serial > 80000:
        return ""
    return (EXCEL_DATE_BASE + timedelta(days=int(serial))).isoformat()


def _extract_task_number(workbook) -> str:
    full_label = TASK_LABEL_FULL.casefold()
    short_label = TASK_LABEL_SHORT.casefold()
    for sheet in workbook.worksheets:
        max_row = min(int(sheet.max_row or 0), 50)
        max_col = min(int(sheet.max_column or 0), 20)
        for row_number in range(1, max_row + 1):
            for col_number in range(1, max_col + 1):
                text = _text(sheet.cell(row_number, col_number).value)
                if text == "":
                    continue
                folded = text.casefold()
                for label in (full_label, short_label):
                    position = folded.find(label)
                    if position < 0:
                        continue
                    tail = _cleanup_task_tail(text[position + len(label) :])
                    if tail:
                        return tail
                    neighbor = _neighbor_task_value(sheet, row_number, col_number)
                    if neighbor:
                        return neighbor
    return ""


def _neighbor_task_value(sheet: Worksheet, row_number: int, col_number: int) -> str:
    for offset in range(1, 4):
        value = _text(sheet.cell(row_number, col_number + offset).value)
        if value:
            return value
    return ""


def _cleanup_task_tail(value: str) -> str:
    text = value.replace(":", "").replace("#", "")
    return " ".join(text.split())


def _find_numbering_col(header_map: dict[str, int]) -> int:
    number_sign = chr(8470).casefold()
    for key, column in header_map.items():
        lowered = key.casefold()
        if (
            number_sign in lowered
            or lowered.startswith("no")
            or "pp" in lowered
            or "p/p" in lowered
            or "p-p" in lowered
        ):
            return int(column)
    return 0




def _detect_value_columns(header_map: dict[str, int]) -> _HeaderValueColumns:
    unit_price_candidates: list[_ValueColumn] = []
    total_price_candidates: list[_ValueColumn] = []
    labor_unit_col = 0
    labor_total_col = 0
    machine_labor_unit_col = 0
    machine_labor_total_col = 0

    for key, column in header_map.items():
        if _is_average_value_header(key):
            continue
        if machine_labor_unit_col == 0 and _is_machine_labor_unit_header(key):
            machine_labor_unit_col = int(column)
            continue
        if machine_labor_total_col == 0 and _is_machine_labor_total_header(key):
            machine_labor_total_col = int(column)
            continue
        if labor_unit_col == 0 and _is_labor_unit_header(key):
            labor_unit_col = int(column)
            continue
        if labor_total_col == 0 and _is_labor_total_header(key):
            labor_total_col = int(column)
            continue
        if _is_unit_price_header(key):
            unit_price_candidates.append(_ValueColumn(int(column), _vat_divisor(key)))
            continue
        if _is_total_price_header(key):
            total_price_candidates.append(_ValueColumn(int(column), _vat_divisor(key)))

    return _HeaderValueColumns(
        unit_price=_prefer_without_vat(unit_price_candidates),
        total_price=_prefer_without_vat(total_price_candidates),
        labor_unit_col=labor_unit_col,
        labor_total_col=labor_total_col,
        machine_labor_unit_col=machine_labor_unit_col,
        machine_labor_total_col=machine_labor_total_col,
    )


def _prefer_without_vat(candidates: list[_ValueColumn]) -> _ValueColumn:
    if not candidates:
        return _ValueColumn(0, 1.0)
    for candidate in candidates:
        if candidate.divisor == 1.0:
            return candidate
    return candidates[0]


def _vat_divisor(normalized_header: str) -> float:
    without_vat = "\u0431\u0435\u0437\u043d\u0434\u0441"
    with_vat = "\u0441\u043d\u0434\u0441"
    if with_vat in normalized_header and without_vat not in normalized_header:
        return VALUE_WITH_VAT_DIVISOR
    return 1.0


def _is_average_value_header(normalized_header: str) -> bool:
    return any(
        token in normalized_header
        for token in (
            "\u0441\u0440\u0435\u0434",
            "\u0441\u0440\u0437\u043d\u0430\u0447",
            "\u0441\u0440.\u0437\u043d\u0430\u0447",
        )
    )


def _is_unit_price_header(normalized_header: str) -> bool:
    if _is_total_price_header(normalized_header):
        return False
    price = "\u0446\u0435\u043d\u0430"
    cost = "\u0441\u0442\u043e\u0438\u043c\u043e\u0441\u0442"
    if normalized_header == price:
        return True
    if "\u0431\u0430\u0437\u043e\u0432\u0430\u044f\u0446\u0435\u043d\u0430" in normalized_header:
        return True
    unit_markers = (
        "\u0435\u0434\u0438\u043d\u0438\u0446",
        "\u0437\u0430\u0435\u0434",
        "\u0437\u04301",
        "\u0435\u0434.",
    )
    return (price in normalized_header or cost in normalized_header) and any(
        marker in normalized_header for marker in unit_markers
    )


def _is_total_price_header(normalized_header: str) -> bool:
    return (
        "\u0438\u0442\u043e\u0433\u043e" in normalized_header
        and "\u0441\u0442\u043e\u0438\u043c\u043e\u0441\u0442" in normalized_header
    )


def _is_labor_unit_header(normalized_header: str) -> bool:
    return (
        "\u0442\u0437\u043d\u0430\u0435\u0434" in normalized_header
        or "\u0442\u0437\u0440\u043d\u0430\u0435\u0434" in normalized_header
    )


def _is_labor_total_header(normalized_header: str) -> bool:
    return (
        "\u0442\u0437\u0432\u0441\u0435\u0433\u043e" in normalized_header
        or "\u0442\u0437\u0440\u0432\u0441\u0435\u0433\u043e" in normalized_header
    )


def _is_machine_labor_unit_header(normalized_header: str) -> bool:
    return "\u0442\u0437\u043c\u043d\u0430\u0435\u0434" in normalized_header


def _is_machine_labor_total_header(normalized_header: str) -> bool:
    return "\u0442\u0437\u043c\u0432\u0441\u0435\u0433\u043e" in normalized_header


def _find_pattern_col(header_map: dict[str, int], patterns: tuple[str, ...]) -> int:
    normalized_patterns = tuple(_normalize_header(pattern) for pattern in patterns)
    for pattern in normalized_patterns:
        for key, column in header_map.items():
            if key == pattern or pattern in key:
                return int(column)
    return 0


def _cell_value(sheet: Worksheet, row_number: int, col_number: int) -> object:
    if col_number <= 0:
        return None
    return sheet.cell(row_number, col_number).value


def _catalog_row_issue(row: CatalogRow) -> str:
    if _text(row.task_id) == "":
        return "missing_task_number"
    if NormCode(row.code) == "":
        return "missing_or_invalid_code"
    if NormUnit(row.unit) == "":
        return "missing_or_invalid_unit"
    if _parse_positive_number(row.price) is None:
        return "missing_or_invalid_price"
    return ""


def _parse_positive_number(value: object) -> float | None:
    number = _parse_number(value)
    if number is None or number <= 0:
        return None
    return number


def _parse_optional_value(value: object, *, divisor: float = 1.0) -> float | None:
    number = _parse_number(value)
    if number is None:
        return None
    return number / divisor


def _parse_number(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Real):
        return float(value)
    if isinstance(value, (date, datetime)):
        return None
    text = str(value).strip()
    if text == "":
        return None
    text = text.replace(" ", " ").replace(" ", "")
    text = "".join(char for char in text if char.isdigit() or char in {".", ",", "-"})
    if text in {"", ".", ",", "-"}:
        return None
    if "," in text and "." in text:
        comma_pos = text.rfind(",")
        dot_pos = text.rfind(".")
        decimal_sep = "," if comma_pos > dot_pos else "."
        thousands_sep = "." if decimal_sep == "," else ","
        text = text.replace(thousands_sep, "")
        text = text.replace(decimal_sep, ".")
    else:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _count_import_status(entries: list[RnmcZipCatalogImportEntry], status: str) -> int:
    return sum(1 for entry in entries if entry.status == status)


def _is_name_header(value: str) -> bool:
    name_works = _normalize_header(HEADER_NAME_WORKS)
    name_short = _normalize_header(HEADER_NAME_SHORT)
    return value in {name_works, name_short} or value.startswith(name_short)


def _is_unit_header(value: str) -> bool:
    units = {
        _normalize_header(HEADER_UNIT),
        _normalize_header(HEADER_UNIT_NO_DOT),
        _normalize_header(HEADER_UNIT_LONG),
    }
    return any(value.startswith(unit) or unit in value for unit in units if unit)


def _is_qty_header(value: str) -> bool:
    return _normalize_header(HEADER_QTY_SHORT) in value or _normalize_header(HEADER_QTY_LONG) in value


def _header_keys(value: object) -> list[str]:
    text = _text(value)
    if text == "":
        return []
    keys: list[str] = []
    for part in (text, text.splitlines()[0] if text.splitlines() else ""):
        normalized = _normalize_header(part)
        if normalized and normalized not in keys:
            keys.append(normalized)
    return keys


def _normalize_header(value: object) -> str:
    text = _text(value).casefold()
    text = text.replace("\u00a0", " ").replace("\r", " ").replace("\n", " ")
    return "".join(text.split())


def _normalize_archive_path(value: str) -> str:
    path = str(value).replace("\\", "/").strip("/")
    return "/".join(part for part in path.split("/") if part not in {"", "."})


def _is_supported_excel_path(path: str) -> bool:
    pure = PurePosixPath(path)
    name = pure.name
    if name.startswith("~$"):
        return False
    return pure.suffix.casefold() in EXCEL_SUFFIXES


def _region_from_archive_path(path: str) -> str:
    parent = PurePosixPath(path).parent
    if str(parent) in {"", "."}:
        return ""
    return parent.name.strip()


def _empty_preview(
    archive_path: str,
    filename: str,
    region_folder: str,
    status: str,
    reason: str,
    *,
    task_number: str = "",
    lsr_quarter: str = "",
    planned_start: str = "",
    planned_finish: str = "",
) -> RnmcWorkbookPreview:
    return RnmcWorkbookPreview(
        archive_path=archive_path,
        filename=filename,
        region_folder=region_folder,
        status=status,
        reason=reason,
        sheet_name="",
        header_row=0,
        task_number=task_number,
        lsr_quarter=lsr_quarter,
        planned_start=planned_start,
        planned_finish=planned_finish,
        rows_ok=0,
        rows_rejected=0,
        sample_rows=[],
    )


def _count_status(entries: list[RnmcWorkbookPreview], status: str) -> int:
    return sum(1 for entry in entries if entry.status == status)


def _is_blank(value: object) -> bool:
    return _text(value) == ""


def _text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()
