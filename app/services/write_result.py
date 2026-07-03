"""Application service: run matching and write the result into a copy.

Ties the tested pieces together end to end: read catalog + estimate flexibly
(Step 4c: sheet selection, template-or-detected columns, clear errors),
resolve the regional coefficient by label (R16), run matching, then write the
structured result into a `WA` copy of the estimate workbook. One function
controls the row positions, the column plan, and the run result, so they stay
aligned for the writer.
"""

from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from app.services.read_estimate import METHOD_TEMPLATE, EstimateData, load_estimate
from app.services.run_matching import MatchingRunResult, run_matching
from core.excel_io import Settings, read_catalog_rows
from core.excel_writer import WriteReport, WriterColumns, resolve_kr_column, write_run_result
from core.exclusions import NameExclusionRule, TaskColorEntry
from core.layout import FIELD_SECTION, LayoutConfig, load_layout_config, resolve_regional_coefficient
from core.risk import DEFAULT_PRICE_SPREAD_LIMIT, GesnException

WA_SUFFIX = " WA"


@dataclass(frozen=True)
class RunAndWriteResult:
    """End-to-end outcome of a matching run plus its Excel export."""

    result: MatchingRunResult
    write_report: WriteReport
    regional_coefficient: float
    coefficient_method: str
    output_path: Path
    sheet_title: str
    read_method: str


def run_and_write(
    catalog_path: str | Path,
    estimate_path: str | Path,
    output_path: str | Path | None = None,
    *,
    settings: Settings | None = None,
    selected_sheet_title: str | None = None,
    name_exclusion_rules: list[NameExclusionRule] | None = None,
    task_color_entries: list[TaskColorEntry] | None = None,
    gesn_exceptions: dict[str, GesnException] | None = None,
    demontazh_filter_enabled: bool = True,
    price_spread_limit: float = DEFAULT_PRICE_SPREAD_LIMIT,
    regional_coefficient: float | None = None,
    layout_config: LayoutConfig | None = None,
) -> RunAndWriteResult:
    """Run matching over the files and write the result into a `WA` copy."""
    active_settings = Settings() if settings is None else settings
    config = load_layout_config() if layout_config is None else layout_config

    catalog_rows = read_catalog_rows(catalog_path, active_settings)
    estimate = load_estimate(
        estimate_path,
        settings=active_settings,
        config=config,
        selected_sheet_title=selected_sheet_title,
    )
    row_numbers = [row_number for row_number, _ in estimate.positioned_rows]
    estimate_rows = [estimate_row for _, estimate_row in estimate.positioned_rows]

    coefficient, coefficient_method = _resolve_coefficient(
        estimate_path,
        estimate.sheet_title,
        regional_coefficient,
        config,
    )

    result = run_matching(
        catalog_rows,
        estimate_rows,
        name_exclusion_rules=name_exclusion_rules,
        gesn_exceptions=gesn_exceptions,
        demontazh_filter_enabled=demontazh_filter_enabled,
        price_spread_limit=price_spread_limit,
        regional_coefficient=coefficient,
    )

    destination = _resolve_output_path(estimate_path, output_path)
    writer_columns = _writer_columns(
        estimate,
        active_settings,
        estimate_path,
        estimate.sheet_title,
    )
    write_report = write_run_result(
        estimate_path,
        destination,
        result,
        row_numbers,
        columns=writer_columns,
        regional_coefficient=coefficient,
        sheet_title=estimate.sheet_title,
        task_color_entries=task_color_entries,
    )

    return RunAndWriteResult(
        result=result,
        write_report=write_report,
        regional_coefficient=coefficient,
        coefficient_method=coefficient_method,
        output_path=destination,
        sheet_title=estimate.sheet_title,
        read_method=estimate.method,
    )


def _writer_columns(
    estimate: EstimateData,
    settings: Settings,
    estimate_path: str | Path,
    sheet_title: str,
) -> WriterColumns:
    if estimate.method == METHOD_TEMPLATE:
        return WriterColumns(
            base_price=settings.col_f,
            code=settings.col_search,
            code_kr=settings.col_kr,
            section=settings.col_section,
            analog_start=settings.col_analog_start,
            header_row=estimate.header_row,
        )

    section = estimate.layout.column(FIELD_SECTION)
    if section is None:
        section = settings.col_section
        analog_start = settings.col_analog_start
    else:
        analog_start = section + 1

    workbook = load_workbook(estimate_path, read_only=True, data_only=True)
    try:
        worksheet = workbook[sheet_title]
        code_kr = resolve_kr_column(
            worksheet,
            estimate.header_row,
            estimate.code_column,
            settings.col_kr,
            settings.col_search,
        )
    finally:
        workbook.close()

    return WriterColumns(
        base_price=estimate.base_price_column,
        code=estimate.code_column,
        code_kr=code_kr,
        section=section,
        analog_start=analog_start,
        header_row=estimate.header_row,
    )


def _resolve_coefficient(
    estimate_path: str | Path,
    sheet_title: str,
    explicit: float | None,
    config: LayoutConfig,
) -> tuple[float, str]:
    if explicit is not None:
        return explicit, "explicit"

    workbook = load_workbook(estimate_path, data_only=True)
    try:
        worksheet = workbook[sheet_title]
        resolution = resolve_regional_coefficient(worksheet, config)
    finally:
        workbook.close()

    return resolution.value, resolution.method


def _resolve_output_path(
    estimate_path: str | Path,
    output_path: str | Path | None,
) -> Path:
    if output_path is not None:
        return Path(output_path)

    source = Path(estimate_path)
    return source.with_name(f"{source.stem}{WA_SUFFIX}{source.suffix}")
