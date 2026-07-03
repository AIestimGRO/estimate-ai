"""Excel readers for catalog and estimate workbooks."""

from dataclasses import dataclass
from numbers import Real
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from core.catalog import CatalogRow
from core.layout import (
    CATALOG_FIELD_ADDED_DATE,
    CATALOG_FIELD_CODE,
    CATALOG_FIELD_PRICE,
    CATALOG_FIELD_REGION,
    CATALOG_FIELD_TASK_ID,
    CATALOG_FIELD_UNIT,
    CATALOG_FIELD_WORK_NAME,
    DEFAULT_MAX_BLANK_RUN,
    LayoutResult,
    data_row_numbers,
    load_catalog_layout_config,
    resolve_catalog_layout,
)
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
    price_spread_limit: float = 3.0


def catalog_default_columns(settings: Settings) -> dict[str, int]:
    """Template column map used as the catalog resolver's fallback defaults."""
    return {
        CATALOG_FIELD_TASK_ID: settings.cat_task_col,
        CATALOG_FIELD_WORK_NAME: settings.cat_work_name_col,
        CATALOG_FIELD_UNIT: settings.cat_unit_col,
        CATALOG_FIELD_PRICE: settings.cat_price_col,
        CATALOG_FIELD_CODE: settings.cat_code_col,
        CATALOG_FIELD_REGION: settings.cat_region_col,
        CATALOG_FIELD_ADDED_DATE: settings.cat_added_date_col,
    }


def _catalog_column(
    layout: LayoutResult,
    field: str,
    defaults: dict[str, int],
) -> int:
    column = layout.column(field)
    if column is not None:
        return column
    return defaults[field]


def read_catalog_rows(
    workbook_path: str | Path,
    settings: Settings | None = None,
) -> list[CatalogRow]:
    """Read raw catalog rows from the catalog sheet.

    Uses ``data_only=True`` so formula-driven price/date cells resolve to the
    values Excel cached on save (real files store prices as formulas), instead
    of returning the formula text (OPEN_ITEMS #2).
    """
    return [row for _, row in read_catalog_rows_with_positions(workbook_path, settings)]


def read_catalog_rows_with_positions(
    workbook_path: str | Path,
    settings: Settings | None = None,
) -> list[tuple[int, CatalogRow]]:
    """Read catalog rows with their 1-based Excel row numbers."""
    active_settings = Settings() if settings is None else settings
    workbook = load_workbook(workbook_path, data_only=True)

    try:
        worksheet = _find_sheet_by_part(workbook.worksheets, CATALOG_SHEET_PART)
        data_limit = int(getattr(worksheet, "max_row", 0) or 0)
        defaults = catalog_default_columns(active_settings)
        catalog_config = load_catalog_layout_config()
        layout = (
            resolve_catalog_layout(worksheet, catalog_config, default_columns=defaults)
            if catalog_config is not None
            else None
        )

        if layout is not None and layout.ok and layout.header_row > 0:
            data_start = layout.header_row + catalog_config.data_start_offset
            task_col = _catalog_column(layout, CATALOG_FIELD_TASK_ID, defaults)
            price_col = _catalog_column(layout, CATALOG_FIELD_PRICE, defaults)
            code_col = _catalog_column(layout, CATALOG_FIELD_CODE, defaults)
            unit_col = _catalog_column(layout, CATALOG_FIELD_UNIT, defaults)
            work_name_col = _catalog_column(layout, CATALOG_FIELD_WORK_NAME, defaults)
            region_col = _catalog_column(layout, CATALOG_FIELD_REGION, defaults)
            added_date_col = _catalog_column(layout, CATALOG_FIELD_ADDED_DATE, defaults)
        else:
            data_start = 4
            task_col = active_settings.cat_task_col
            price_col = active_settings.cat_price_col
            code_col = active_settings.cat_code_col
            unit_col = active_settings.cat_unit_col
            work_name_col = active_settings.cat_work_name_col
            region_col = active_settings.cat_region_col
            added_date_col = active_settings.cat_added_date_col

        rows: list[tuple[int, CatalogRow]] = []

        for row_number in range(data_start, data_limit + 1):
            rows.append(
                (
                    row_number,
                    CatalogRow(
                        task_id=_cell_value(worksheet, row_number, task_col),
                        price=_cell_value(worksheet, row_number, price_col),
                        code=_cell_value(worksheet, row_number, code_col),
                        unit=_cell_value(worksheet, row_number, unit_col),
                        work_name=_cell_value(
                            worksheet,
                            row_number,
                            work_name_col,
                        ),
                        region=_cell_value(worksheet, row_number, region_col),
                        added_date=_cell_value(
                            worksheet,
                            row_number,
                            added_date_col,
                        ),
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
    Values are read with ``data_only=True`` so formula prices resolve to their
    cached numbers (OPEN_ITEMS #2).
    """
    active_settings = Settings() if settings is None else settings
    workbook = load_workbook(workbook_path, data_only=True)

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

        # Skip the column-enumeration row under the header (e.g. "1 2 3 ...").
        # Real codes always contain letters/dashes, so a bare integer in the
        # code cell is a header artefact, not a work item (R4-lite).
        if str(code).strip().isdigit():
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
