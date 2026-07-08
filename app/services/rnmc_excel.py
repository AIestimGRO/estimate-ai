"""RNMC workbook preview helpers.

This module mirrors the legacy VBA import detection rules, but it does not
write catalog rows. It is used to preview what an RNMC zip import would find.
"""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from io import BytesIO
from pathlib import PurePosixPath
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from app.services.rnmc_zip import EXCEL_SUFFIXES
from core.storage.catalog import filename_is_final_for_preview, normalize_import_filename

HEADER_NAME_WORKS = "\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435 \u0440\u0430\u0431\u043e\u0442"
HEADER_NAME_SHORT = "\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435"
HEADER_UNIT = "\u0415\u0434.\u0438\u0437\u043c."
HEADER_QTY_SHORT = "\u041a\u043e\u043b-\u0432\u043e"
HEADER_QTY_LONG = "\u041a\u043e\u043b\u0438\u0447\u0435\u0441\u0442\u0432\u043e"
TASK_LABEL_FULL = "\u2116 \u0437\u0430\u0434\u0430\u0447\u0438 1\u0424"
TASK_LABEL_SHORT = "\u2116 \u0437\u0430\u0434\u0430\u0447\u0438"

STATUS_PREVIEW_OK = "preview_ok"
STATUS_NO_TABLE = "no_table"
STATUS_NO_ROWS = "no_rows"
STATUS_SKIPPED_PROCESSED = "skipped_processed"
STATUS_DUPLICATE_NAME = "duplicate_name"
STATUS_UNSUPPORTED_FORMAT = "unsupported_format"
STATUS_PARSE_ERROR = "parse_error"

SUPPORTED_PREVIEW_SUFFIXES = frozenset({".xlsx", ".xlsm"})


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


def _preview_workbook_bytes(
    data: bytes,
    archive_path: str,
    filename: str,
    region_folder: str,
) -> RnmcWorkbookPreview:
    workbook = load_workbook(BytesIO(data), read_only=True, data_only=True)
    try:
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
            normalized = _normalize_header(value)
            if normalized == "":
                continue
            header_map[normalized] = index
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


def _is_name_header(value: str) -> bool:
    name_works = _normalize_header(HEADER_NAME_WORKS)
    name_short = _normalize_header(HEADER_NAME_SHORT)
    return value in {name_works, name_short} or value.startswith(name_short)


def _is_unit_header(value: str) -> bool:
    unit = _normalize_header(HEADER_UNIT)
    return value.startswith(unit) or unit in value


def _is_qty_header(value: str) -> bool:
    return _normalize_header(HEADER_QTY_SHORT) in value or _normalize_header(HEADER_QTY_LONG) in value


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
