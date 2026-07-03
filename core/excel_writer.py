"""Write a matching run result back into a copy of the estimate workbook.

Pure Excel output primitive (symmetric to core/excel_io.py): given a source
estimate file, the structured run result, the physical worksheet row for each
result row, and an explicit column plan, it writes analog columns, the
average-price formula, the section code, the `/KR` code suffix, cell
colouring, and a risk-check log sheet into a *copy* of the workbook. The
source file is never modified.

Ports the output side-effects of ProcessSmeta (Module4), DOMAIN_RULES.md
section 6, and uses the average-column placement rule from core/layout.py
(R12). It makes no matching or pricing decisions of its own. The column plan
is supplied by the caller so both the fixed template layout and a detected
layout (Step 4c) can be written correctly.
"""

from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.services.run_matching import EstimateRowResult, MatchingRunResult
from core.excel_io import find_estimate_sheet
from core.exclusions import TaskColorEntry, is_task_marked
from core.layout import resolve_average_placement
from core.risk import REASON_RATIO_EXCEEDED

RISK_LOG_SHEET = "Price_Check_Log"

# Module4_updated.bas RGB fills
HEADER_FILL = PatternFill(start_color="FFD9E1F2", end_color="FFD9E1F2", fill_type="solid")
TASK_FILL = PatternFill(start_color="FFDDEBF7", end_color="FFDDEBF7", fill_type="solid")
DUP_FILL = PatternFill(start_color="FFD9D9D9", end_color="FFD9D9D9", fill_type="solid")
PROBLEM_FILL = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")

PRICE_NUMBER_FORMAT = "#,##0.00"
HEADER_FONT = Font(bold=True, size=9)
REGION_FONT = Font(italic=True, size=9)
AVG_FONT = Font(bold=True)

# Header labels for columns inserted when absent (synonyms live in layout.json).
KR_HEADER_LABEL = "/\u041a\u0420"
SECTION_HEADER_LABEL = "\u041a\u043e\u0434 \u0440\u0430\u0437\u0434\u0435\u043b\u0430"

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
class WriterColumns:
    """Where the writer should place each output on the worksheet."""

    base_price: int
    code: int
    code_kr: int
    section: int | None
    analog_start: int | None
    header_row: int = 0


@dataclass(frozen=True)
class WriteReport:
    """Summary of what the writer produced."""

    output_path: Path
    written_rows: int
    inserted_average_column: bool
    average_column: int
    analog_start_column: int
    analog_column_count: int
    risk_log_rows: int


@dataclass(frozen=True)
class AnalogColumnDef:
    """One global analog output column (task header + region sub-header)."""

    column: int
    task_id: str
    price_position: int
    region: str


@dataclass(frozen=True)
class GlobalAnalogPlan:
    """Stable (task_id, price_position) -> column map for the whole run."""

    by_key: dict[tuple[str, int], int]
    columns: tuple[AnalogColumnDef, ...]
    last_column: int


def write_run_result(
    source_path: str | Path,
    output_path: str | Path,
    result: MatchingRunResult,
    row_numbers: list[int],
    *,
    columns: WriterColumns,
    regional_coefficient: float = 1.0,
    sheet_title: str | None = None,
    task_color_entries: list[TaskColorEntry] | None = None,
) -> WriteReport:
    """Write `result` into a copy of the estimate workbook at `output_path`."""
    if len(row_numbers) != len(result.rows):
        raise ValueError("row_numbers length must match result.rows length")

    task_colors = [] if task_color_entries is None else task_color_entries
    workbook = load_workbook(source_path, data_only=False)

    try:
        if sheet_title is not None:
            worksheet = workbook[sheet_title]
        else:
            worksheet = find_estimate_sheet(workbook)

        placement = _plan_average_column(worksheet, columns.base_price, row_numbers)
        if placement.needs_insert:
            worksheet.insert_cols(placement.column)

        layout_columns = _ensure_kr_and_section_columns(
            worksheet,
            columns,
            placement,
            row_numbers,
        )
        analog_start_base = (
            max(worksheet.max_column + 1, placement.column + 1)
            if layout_columns.analog_start is None
            else layout_columns.analog_start
        )
        output_columns = _plan_output_columns(layout_columns, placement, analog_start_base)

        header_row = _effective_header_row(layout_columns, row_numbers)
        analog_plan = _build_global_analog_plan(result, output_columns.analog_start)
        last_data_row = max(row_numbers) if row_numbers else output_columns.analog_start
        if analog_plan.columns:
            _clear_analog_block(
                worksheet,
                header_row or output_columns.analog_start,
                last_data_row,
                output_columns.analog_start,
                analog_plan.last_column,
            )
            _write_analog_headers(
                worksheet,
                header_row,
                analog_plan,
                last_data_row,
                task_colors,
            )

        written_rows = 0
        for row_number, row in zip(row_numbers, result.rows):
            if _write_row(
                worksheet,
                row_number,
                row,
                output_columns,
                analog_plan,
                regional_coefficient,
                task_colors,
            ):
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
        average_column=output_columns.average,
        analog_start_column=output_columns.analog_start,
        analog_column_count=len(analog_plan.columns),
        risk_log_rows=risk_log_rows,
    )


@dataclass(frozen=True)
class _OutputColumns:
    base_price: int
    average: int
    code_kr: int
    section: int | None
    analog_start: int


def _plan_average_column(
    worksheet: Worksheet,
    base_price_column: int,
    row_numbers: list[int],
):
    neighbour = base_price_column + 1
    occupied: set[int] = set()
    for row_number in row_numbers:
        if worksheet.cell(row=row_number, column=neighbour).value not in (None, ""):
            occupied.add(neighbour)
            break
    return resolve_average_placement(base_price_column, occupied)


def _plan_output_columns(columns: WriterColumns, placement, analog_start_base: int) -> _OutputColumns:
    return _OutputColumns(
        base_price=columns.base_price,
        average=placement.column,
        code_kr=columns.code_kr,
        section=columns.section,
        analog_start=analog_start_base,
    )


def _effective_header_row(columns: WriterColumns, row_numbers: list[int]) -> int:
    if columns.header_row > 0:
        return columns.header_row
    if not row_numbers:
        return 0
    return min(row_numbers) - 2


def _shifted_column(column: int, placement) -> int:
    if placement.needs_insert and column >= placement.column:
        return column + 1
    return column


def _ensure_kr_and_section_columns(
    worksheet: Worksheet,
    columns: WriterColumns,
    placement,
    row_numbers: list[int],
) -> WriterColumns:
    """Ensure dedicated `/KR` and section columns exist (R11, DOMAIN_RULES.md section 6).

    The GESN code column keeps the plain code; rows with analogs receive
    ``code + /KR`` in the adjacent `/KR` column. Section codes go into their
    own column. Missing columns are inserted after the average-column shift.
    """
    header_row = _effective_header_row(columns, row_numbers)
    code_col = _shifted_column(columns.code, placement)

    kr_col = _find_kr_column(worksheet, header_row, code_col)
    section_col = (
        _shifted_column(columns.section, placement)
        if columns.section is not None
        else _find_section_column(worksheet, header_row, code_col)
    )

    pending: list[tuple[int, str]] = []

    if kr_col is None:
        kr_col = code_col + 1
        pending.append((kr_col, KR_HEADER_LABEL))
        if section_col is not None and section_col >= kr_col:
            section_col += 1

    if section_col is None or section_col <= kr_col:
        section_col = kr_col + 1
        if not any(insert_at == section_col for insert_at, _ in pending):
            pending.append((section_col, SECTION_HEADER_LABEL))

    for insert_at, label in sorted(pending, key=lambda item: item[0], reverse=True):
        worksheet.insert_cols(insert_at)
        if header_row > 0:
            header_cell = worksheet.cell(row=header_row, column=insert_at)
            header_cell.value = label
            header_cell.font = HEADER_FONT
            header_cell.fill = HEADER_FILL

    analog_start = section_col + 1
    return WriterColumns(
        base_price=columns.base_price,
        code=code_col,
        code_kr=kr_col,
        section=section_col,
        analog_start=analog_start,
        header_row=columns.header_row,
    )


def _find_kr_column(
    worksheet: Worksheet,
    header_row: int,
    code_column: int,
) -> int | None:
    if header_row <= 0:
        return None

    for column in range(code_column + 1, code_column + 4):
        if _is_kr_header(worksheet.cell(row=header_row, column=column).value):
            return column
    return None


def _find_section_column(
    worksheet: Worksheet,
    header_row: int,
    code_column: int,
) -> int | None:
    if header_row <= 0:
        return None

    for column in range(code_column + 1, code_column + 6):
        if _is_section_header(worksheet.cell(row=header_row, column=column).value):
            return column
    return None


def _is_kr_header(value: object) -> bool:
    text = _normalize_header_text(value)
    return text == "\u043a\u0440" or text.startswith("/")


def _is_section_header(value: object) -> bool:
    text = _normalize_header_text(value)
    return "\u043a\u043e\u0434 \u0440\u0430\u0437\u0434\u0435\u043b\u0430" in text


def _build_global_analog_plan(
    result: MatchingRunResult,
    analog_start: int,
) -> GlobalAnalogPlan:
    """Assign one worksheet column per (task_id, price_position) pair (Module4 step 2)."""
    task_order: list[str] = []
    seen_tasks: set[str] = set()
    task_max_pi: dict[str, int] = {}
    region_by_key: dict[tuple[str, int], str] = {}

    for row in result.rows:
        for analog in row.analogs:
            task_id = analog.task_id
            if task_id not in seen_tasks:
                seen_tasks.add(task_id)
                task_order.append(task_id)
            task_max_pi[task_id] = max(task_max_pi.get(task_id, 0), analog.price_position)
            region_by_key[(task_id, analog.price_position)] = analog.entry.region

    by_key: dict[tuple[str, int], int] = {}
    column_defs: list[AnalogColumnDef] = []
    next_column = analog_start

    for task_id in task_order:
        for price_position in range(1, task_max_pi[task_id] + 1):
            key = (task_id, price_position)
            by_key[key] = next_column
            column_defs.append(
                AnalogColumnDef(
                    column=next_column,
                    task_id=task_id,
                    price_position=price_position,
                    region=region_by_key.get(key, ""),
                )
            )
            next_column += 1

    last_column = next_column - 1 if column_defs else analog_start - 1
    return GlobalAnalogPlan(by_key=by_key, columns=tuple(column_defs), last_column=last_column)


def _clear_analog_block(
    worksheet: Worksheet,
    header_row: int,
    last_row: int,
    first_col: int,
    last_col: int,
) -> None:
    start_row = header_row if header_row > 0 else 1
    for row in range(start_row, last_row + 1):
        for column in range(first_col, last_col + 1):
            cell = worksheet.cell(row=row, column=column)
            cell.value = None
            cell.fill = PatternFill()


def _write_analog_headers(
    worksheet: Worksheet,
    header_row: int,
    plan: GlobalAnalogPlan,
    last_data_row: int,
    task_colors: list[TaskColorEntry],
) -> None:
    if header_row <= 0 or not plan.columns:
        return

    region_row = header_row + 1
    for column_def in plan.columns:
        column = column_def.column
        task_cell = worksheet.cell(row=header_row, column=column)
        task_cell.value = column_def.task_id
        task_cell.font = HEADER_FONT
        task_cell.fill = HEADER_FILL

        region_cell = worksheet.cell(row=region_row, column=column)
        region_cell.value = column_def.region
        region_cell.font = REGION_FONT
        region_cell.fill = HEADER_FILL

        if is_task_marked(task_colors, column_def.task_id):
            for row in range(header_row, last_data_row + 1):
                worksheet.cell(row=row, column=column).fill = TASK_FILL
            task_cell.fill = HEADER_FILL
            region_cell.fill = HEADER_FILL


def _write_row(
    worksheet: Worksheet,
    row_number: int,
    row: EstimateRowResult,
    columns: _OutputColumns,
    plan: GlobalAnalogPlan,
    coefficient: float,
    task_colors: list[TaskColorEntry],
) -> bool:
    if columns.section is not None and row.section_code:
        section_cell = worksheet.cell(row=row_number, column=columns.section)
        section_cell.number_format = "@"
        section_cell.value = row.section_code

    _write_average_formula(worksheet, row_number, columns, plan)

    if not row.has_analogs:
        return False

    for analog in row.analogs:
        column = plan.by_key[(analog.task_id, analog.price_position)]
        price_cell = worksheet.cell(row=row_number, column=column)
        price_cell.value = round(analog.entry.price * coefficient, 2)
        price_cell.number_format = PRICE_NUMBER_FORMAT
        _apply_cell_colour(price_cell, row, analog, task_colors)

    if row.kr_code is not None:
        kr_cell = worksheet.cell(row=row_number, column=columns.code_kr)
        kr_cell.number_format = "@"
        kr_cell.value = row.kr_code

    return True


def _write_average_formula(
    worksheet: Worksheet,
    row_number: int,
    columns: _OutputColumns,
    plan: GlobalAnalogPlan,
) -> None:
    avg_cell = worksheet.cell(row=row_number, column=columns.average)
    avg_cell.value = _average_formula(
        row_number,
        columns.base_price,
        columns.analog_start,
        plan.last_column,
    )
    avg_cell.number_format = PRICE_NUMBER_FORMAT
    avg_cell.font = AVG_FONT


def _apply_cell_colour(
    cell,
    row: EstimateRowResult,
    analog,
    task_colors: list[TaskColorEntry],
) -> None:
    if is_task_marked(task_colors, analog.task_id):
        cell.fill = TASK_FILL
        return

    flagged_ids = {id(entry) for entry in row.risk_result.flagged_entries}
    colour_all = (
        row.risk_result.is_flagged
        and row.risk_result.reason == REASON_RATIO_EXCEEDED
    )

    if colour_all or id(analog.entry) in flagged_ids:
        cell.fill = PROBLEM_FILL
    elif analog.price_position > 1:
        cell.fill = DUP_FILL


def _average_formula(
    row_number: int,
    base_column: int,
    analog_start: int,
    last_analog_column: int,
) -> str:
    base = f"{get_column_letter(base_column)}{row_number}"
    if last_analog_column < analog_start:
        return f"={base}"

    first = f"{get_column_letter(analog_start)}{row_number}"
    last = f"{get_column_letter(last_analog_column)}{row_number}"
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


def resolve_kr_column(
    worksheet: Worksheet,
    header_row: int,
    code_column: int,
    settings_code_kr: int,
    settings_code_search: int,
) -> int:
    """Pick the `/KR` destination column; never reuse the plain code column."""
    found = _find_kr_column(worksheet, header_row, code_column)
    if found is not None:
        return found

    if settings_code_kr != settings_code_search:
        return settings_code_kr
    return code_column + 1


def _normalize_header_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).lower().strip()
    for char in ("\r", "\n", "\t", ".", ",", ";", ":"):
        text = text.replace(char, " ")
    while "  " in text:
        text = text.replace("  ", " ")
    return text.strip()
