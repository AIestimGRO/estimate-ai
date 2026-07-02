"""Application service: run matching and write the result into a copy.

Ties the tested pieces together end to end: read catalog + estimate (with
physical row positions), resolve the regional coefficient by label (R16),
run matching, then write the structured result into a `WA` copy of the
estimate workbook. One function controls both the row positions and the run
result, so they stay perfectly aligned for the writer.
"""

from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from app.services.run_matching import MatchingRunResult, run_matching
from core.excel_io import (
    Settings,
    find_estimate_sheet,
    read_catalog_rows,
    read_estimate_rows_with_positions,
)
from core.excel_writer import WriteReport, write_run_result
from core.exclusions import NameExclusionRule
from core.layout import LayoutConfig, load_layout_config, resolve_regional_coefficient
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


def run_and_write(
    catalog_path: str | Path,
    estimate_path: str | Path,
    output_path: str | Path | None = None,
    *,
    settings: Settings | None = None,
    name_exclusion_rules: list[NameExclusionRule] | None = None,
    gesn_exceptions: dict[str, GesnException] | None = None,
    demontazh_filter_enabled: bool = True,
    price_spread_limit: float = DEFAULT_PRICE_SPREAD_LIMIT,
    regional_coefficient: float | None = None,
    layout_config: LayoutConfig | None = None,
) -> RunAndWriteResult:
    """Run matching over the files and write the result into a `WA` copy."""
    active_settings = Settings() if settings is None else settings

    catalog_rows = read_catalog_rows(catalog_path, active_settings)
    positioned = read_estimate_rows_with_positions(estimate_path, active_settings)
    row_numbers = [row_number for row_number, _ in positioned]
    estimate_rows = [estimate_row for _, estimate_row in positioned]

    coefficient, coefficient_method = _resolve_coefficient(
        estimate_path,
        active_settings,
        regional_coefficient,
        layout_config,
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
    write_report = write_run_result(
        estimate_path,
        destination,
        result,
        row_numbers,
        settings=active_settings,
        regional_coefficient=coefficient,
    )

    return RunAndWriteResult(
        result=result,
        write_report=write_report,
        regional_coefficient=coefficient,
        coefficient_method=coefficient_method,
        output_path=destination,
    )


def _resolve_coefficient(
    estimate_path: str | Path,
    settings: Settings,
    explicit: float | None,
    layout_config: LayoutConfig | None,
) -> tuple[float, str]:
    if explicit is not None:
        return explicit, "explicit"

    config = load_layout_config() if layout_config is None else layout_config
    workbook = load_workbook(estimate_path, data_only=False)
    try:
        worksheet = find_estimate_sheet(workbook, settings)
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
