"""Write a matching run result back into a copy of the estimate workbook.

Pure Excel output primitive (symmetric to core/excel_io.py): given a source
estimate file, the structured run result, and the physical worksheet row for
each result row, it writes analog columns, the average-price formula, the
section code, the `/KR` code suffix, cell colouring, and a risk-check log
sheet into a *copy* of the workbook. The source file is never modified.

Ports the output side-effects of ProcessSmeta (Module4), DOMAIN_RULES.md
section 6, and uses the average-column placement rule from core/layout.py
(R12). This module makes no matching or pricing decisions of its own.
"""

from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.services.run_matching import EstimateRowResult, MatchingRunResult
from core.excel_io import Settings, find_estimate_sheet
from core.layout import resolve_average_placement
from core.risk import REASON_RATIO_EXCEEDED

RISK_LOG_SHEET = "Price_Check_Log"

DUP_FILL = PatternFill(start_color="FFD9D9D9", end_color="FFD9D9D9", fill_type="solid")
PROBLEM_FILL = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")

_RISK_LOG_HEADER = (
    "estimate_row",
    "code",
    "unit",
    "reason",
    "min_price",
    "max_price",
    "ratio",
    "recommended_price",
)


@dataclass(frozen=True)
class WriteReport:
    """Summary of what the writer produced."""

    output_path: Path
    written_rows: int
    inserted_average_column: bool
    average_column: int
    analog_start_column: int
    risk_log_rows: int


def write_run_result(
    source_path: str | Path,
    output_path: str | Path,
    result: MatchingRunResult,
    row_numbers: list[int],
    *,
    settings: Settings | None = None,
    regional_coefficient: float = 1.0,
) -> WriteReport:
    """Write `result` into a copy of the estimate workbook at `output_path`."""
    if len(row_numbers) != len(result.rows):
        raise ValueError("row_numbers length must match result.rows length")

    active_settings = Settings() if settings is None else settings
    workbook = load_workbook(source_path, data_only=False)

    try:
        worksheet = find_estimate_sheet(workbook, active_settings)

        placement = _plan_average_column(worksheet, active_settings, row_numbers)
        columns = _plan_output_columns(active_settings, placement)
        if placement.needs_insert:
            worksheet.insert_cols(placement.column)

        written_rows = 0
        for row_number, row in zip(row_numbers, result.rows):
            if _write_row(worksheet, row_number, row, columns, regional_coefficient):
                written_rows += 1

        risk_log_rows = _write_risk_log(workbook, result, row_numbers, regional_coefficient)

        destination = Path(output_path)
        destination.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(destination)
    finally:
        workbook.close()

    return WriteReport(
        output_path=destination,
        written_rows=written_rows,
        inserted_average_column=placement.needs_insert,
        average_column=columns.average,
        analog_start_column=columns.analog_start,
        risk_log_rows=risk_log_rows,
    )


@dataclass(frozen=True)
class _OutputColumns:
    base_price: int
    average: int
    code_kr: int
    section: int
    analog_start: int


def _plan_average_column(
    worksheet: Worksheet,
    settings: Settings,
    row_numbers: list[int],
):
    neighbour = settings.col_f + 1
    occupied: set[int] = set()
    for row_number in row_numbers:
        if worksheet.cell(row=row_number, column=neighbour).value not in (None, ""):
            occupied.add(neighbour)
            break
    return resolve_average_placement(settings.col_f, occupied)


def _plan_output_columns(settings: Settings, placement) -> _OutputColumns:
    def shifted(column: int) -> int:
        if placement.needs_insert and column >= placement.column:
            return column + 1
        return column

    return _OutputColumns(
        base_price=settings.col_f,
        average=placement.column,
        code_kr=shifted(settings.col_kr),
        section=shifted(settings.col_section),
        analog_start=shifted(settings.col_analog_start),
    )


def _write_row(
    worksheet: Worksheet,
    row_number: int,
    row: EstimateRowResult,
    columns: _OutputColumns,
    coefficient: float,
) -> bool:
    if row.section_code:
        worksheet.cell(row=row_number, column=columns.section).value = row.section_code

    if not row.has_analogs:
        return False

    analog_count = len(row.analogs)
    for offset, analog in enumerate(row.analogs):
        column = columns.analog_start + offset
        cell = worksheet.cell(row=row_number, column=column)
        cell.value = analog.entry.price * coefficient

    _apply_colours(worksheet, row_number, row, columns.analog_start)

    if row.kr_code is not None:
        worksheet.cell(row=row_number, column=columns.code_kr).value = row.kr_code

    worksheet.cell(row=row_number, column=columns.average).value = _average_formula(
        row_number,
        columns.base_price,
        columns.analog_start,
        analog_count,
    )
    return True


def _apply_colours(
    worksheet: Worksheet,
    row_number: int,
    row: EstimateRowResult,
    analog_start: int,
) -> None:
    flagged_ids = {id(entry) for entry in row.risk_result.flagged_entries}
    colour_all = (
        row.risk_result.is_flagged
        and row.risk_result.reason == REASON_RATIO_EXCEEDED
    )

    for offset, analog in enumerate(row.analogs):
        cell = worksheet.cell(row=row_number, column=analog_start + offset)
        if analog.price_position > 1:
            cell.fill = DUP_FILL
        if colour_all or id(analog.entry) in flagged_ids:
            cell.fill = PROBLEM_FILL


def _average_formula(
    row_number: int,
    base_column: int,
    analog_start: int,
    analog_count: int,
) -> str:
    base = f"{get_column_letter(base_column)}{row_number}"
    first = f"{get_column_letter(analog_start)}{row_number}"
    last = f"{get_column_letter(analog_start + analog_count - 1)}{row_number}"
    return f"=MAX({base}, IFERROR(AVERAGE({base}, {first}:{last}), {base}))"


def _write_risk_log(
    workbook,
    result: MatchingRunResult,
    row_numbers: list[int],
    coefficient: float,
) -> int:
    if RISK_LOG_SHEET in workbook.sheetnames:
        del workbook[RISK_LOG_SHEET]

    log_rows = [
        (row_number, row)
        for row_number, row in zip(row_numbers, result.rows)
        if row.risk_result.is_flagged
    ]
    if not log_rows:
        return 0

    sheet = workbook.create_sheet(title=RISK_LOG_SHEET)
    sheet.append(list(_RISK_LOG_HEADER))

    for row_number, row in log_rows:
        risk = row.risk_result
        min_price, max_price = _log_min_max(row, coefficient)
        sheet.append(
            [
                row_number,
                _text(row.estimate_row.code),
                _text(row.estimate_row.unit),
                risk.reason,
                min_price,
                max_price,
                risk.ratio or None,
                row.recommended_price,
            ]
        )

    return len(log_rows)


def _log_min_max(
    row: EstimateRowResult,
    coefficient: float,
) -> tuple[float | None, float | None]:
    risk = row.risk_result
    entries = risk.flagged_entries

    if risk.min_entry is not None:
        min_price = risk.min_entry.price
    elif entries:
        min_price = min(entry.price for entry in entries)
    else:
        min_price = None

    if risk.max_entry is not None:
        max_price = risk.max_entry.price
    elif entries:
        max_price = max(entry.price for entry in entries)
    else:
        max_price = None

    scaled_min = None if min_price is None else min_price * coefficient
    scaled_max = None if max_price is None else max_price * coefficient
    return scaled_min, scaled_max


def _text(value: object) -> str:
    return "" if value is None else str(value)
