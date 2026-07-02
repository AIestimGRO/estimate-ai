"""Excel readers for catalog and estimate workbooks."""

from dataclasses import dataclass
from numbers import Real
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from core.catalog import CatalogRow
from core.layout import DEFAULT_MAX_BLANK_RUN, data_row_numbers
from core.matching import EstimateRow
from core.normalize import NormCode, NormUnit


CATALOG_SHEET_PART = "\u041a\u0430\u0442"
ESTIMATE_SHEET_PART = "\u041e\u0421"
GESN_MARKER = "\u0413\u042d\u0421\u041d"
FER_MARKER = "\u0424\u0415\u0420"
PERECHEN_MARKER = "\u041f\u0435\u0440"


@dataclass(frozen=True)
class Settings:
    """Column settings mirroring Module1.TSettings defaults."""

    col_search: int = 14
    col_avg: int = 7
    col_kr: int = 14
    col_f: int = 6
    col_section: int = 15
    col_analog_start: int = 16
    col_smeta_work_name: int = 3
    col_smeta_unit: int = 4
    cat_task_col: int = 2
    cat_price_col: int = 7
    cat_code_col: int = 14
    cat_region_col: int = 16
    cat_work_name_col: int = 3
    cat_unit_col: int = 4
    cat_added_date_col: int = 17


def read_catalog_rows(
    workbook_path: str | Path,
    settings: Settings | None = None,
) -> list[CatalogRow]:
    """Read raw catalog rows from the catalog sheet."""
    active_settings = Settings() if settings is None else settings
    workbook = load_workbook(workbook_path, data_only=False)

    try:
        worksheet = _find_sheet_by_part(workbook.worksheets, CATALOG_SHEET_PART)
        rows: list[CatalogRow] = []

        for row_number in range(4, worksheet.max_row + 1):
            rows.append(
                CatalogRow(
                    task_id=_cell_value(worksheet, row_number, active_settings.cat_task_col),
                    price=_cell_value(worksheet, row_number, active_settings.cat_price_col),
                    code=_cell_value(worksheet, row_number, active_settings.cat_code_col),
                    unit=_cell_value(worksheet, row_number, active_settings.cat_unit_col),
                    work_name=_cell_value(
                        worksheet,
                        row_number,
                        active_settings.cat_work_name_col,
                    ),
                    region=_cell_value(worksheet, row_number, active_settings.cat_region_col),
                    added_date=_cell_value(
                        worksheet,
                        row_number,
                        active_settings.cat_added_date_col,
                    ),
                )
            )

        return rows
    finally:
        workbook.close()


def read_estimate_rows(
    workbook_path: str | Path,
    settings: Settings | None = None,
) -> list[EstimateRow]:
    """Read raw working estimate rows from the estimate sheet."""
    return [row for _, row in read_estimate_rows_with_positions(workbook_path, settings)]


def read_estimate_rows_with_positions(
    workbook_path: str | Path,
    settings: Settings | None = None,
) -> list[tuple[int, EstimateRow]]:
    """Read working estimate rows tagged with their physical worksheet row.

    The physical row number is needed by the Excel writer to write results
    back onto the exact source row; matching itself stays position-agnostic.
    """
    active_settings = Settings() if settings is None else settings
    workbook = load_workbook(workbook_path, data_only=False)

    try:
        worksheet = find_estimate_sheet(workbook, active_settings)
        return _collect_estimate_rows(worksheet, active_settings)
    finally:
        workbook.close()


def _collect_estimate_rows(
    worksheet: Worksheet,
    settings: Settings,
) -> list[tuple[int, EstimateRow]]:
    header_row = _detect_estimate_header_row(worksheet, settings)
    if header_row == 0:
        return []

    rows: list[tuple[int, EstimateRow]] = []
    for row_number in range(header_row + 2, worksheet.max_row + 1):
        code = _cell_value(worksheet, row_number, settings.col_search)
        unit = _cell_value(worksheet, row_number, settings.col_smeta_unit)
        base_price = _cell_value(worksheet, row_number, settings.col_f)

        if not _is_working_estimate_row(code, unit, base_price):
            continue

        rows.append(
            (
                row_number,
                EstimateRow(
                    code=code,
                    unit=unit,
                    work_name=_cell_value(
                        worksheet,
                        row_number,
                        settings.col_smeta_work_name,
                    ),
                    base_price=base_price,
                ),
            )
        )

    return rows


def find_estimate_sheet(
    workbook: Any,
    settings: Settings | None = None,
) -> Worksheet:
    """Locate the estimate worksheet (name contains the estimate marker)."""
    return _find_sheet_by_part(workbook.worksheets, ESTIMATE_SHEET_PART)


def read_estimate_rows_from_worksheet(
    worksheet: Worksheet,
    settings: Settings | None = None,
) -> list[tuple[int, EstimateRow]]:
    """Template read of an already-open worksheet (fixed Settings columns)."""
    active_settings = Settings() if settings is None else settings
    return _collect_estimate_rows(worksheet, active_settings)


def read_estimate_rows_by_columns(
    worksheet: Worksheet,
    *,
    header_row: int,
    code_column: int,
    unit_column: int,
    work_name_column: int,
    base_price_column: int,
    max_blank_run: int = DEFAULT_MAX_BLANK_RUN,
) -> list[tuple[int, EstimateRow]]:
    """Read estimate rows using explicitly resolved columns (flexible path).

    Body rows are scanned tolerating isolated blank rows (data_row_numbers,
    R5); each row is still validated as a working estimate row.
    """
    key_columns = [code_column, unit_column, base_price_column]
    rows: list[tuple[int, EstimateRow]] = []

    for row_number in data_row_numbers(
        worksheet,
        header_row + 1,
        key_columns,
        max_blank_run=max_blank_run,
    ):
        code = _cell_value(worksheet, row_number, code_column)
        unit = _cell_value(worksheet, row_number, unit_column)
        base_price = _cell_value(worksheet, row_number, base_price_column)

        if not _is_working_estimate_row(code, unit, base_price):
            continue

        rows.append(
            (
                row_number,
                EstimateRow(
                    code=code,
                    unit=unit,
                    work_name=_cell_value(worksheet, row_number, work_name_column),
                    base_price=base_price,
                ),
            )
        )

    return rows


def _find_sheet_by_part(worksheets: list[Worksheet], part: str) -> Worksheet:
    part_key = part.casefold()
    for worksheet in worksheets:
        if part_key in worksheet.title.casefold():
            return worksheet
    return worksheets[0]


def _detect_estimate_header_row(worksheet: Worksheet, settings: Settings) -> int:
    markers = tuple(marker.casefold() for marker in (GESN_MARKER, FER_MARKER, PERECHEN_MARKER))

    for row_number in range(1, 51):
        value = _cell_value(worksheet, row_number, settings.col_search)
        value_key = "" if value is None else str(value).casefold()
        if any(marker in value_key for marker in markers):
            return row_number

    return 0


def _is_working_estimate_row(code: object, unit: object, base_price: object) -> bool:
    if NormCode(code) == "":
        return False
    if NormUnit(unit) == "":
        return False

    parsed_base_price = _parse_number(base_price)
    return parsed_base_price is not None and parsed_base_price > 0


def _parse_number(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Real):
        return float(value)

    text = str(value).strip()
    if text == "":
        return None

    try:
        return float(text)
    except ValueError:
        return None


def _cell_value(worksheet: Worksheet, row: int, column: int) -> Any:
    return worksheet.cell(row=row, column=column).value
