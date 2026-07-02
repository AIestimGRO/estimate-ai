from pathlib import Path

from openpyxl import Workbook, load_workbook

from app.services.run_matching import KR_END
from app.services.write_result import run_and_write
from core.excel_writer import PROBLEM_FILL, RISK_LOG_SHEET
from core.risk import REASON_RATIO_EXCEEDED

METER = "\u043c"
CODE = "\u0413\u042d\u0421\u041d01-01-001-01"
INSTALLATION = "\u043c\u043e\u043d\u0442\u0430\u0436"
REGION_LABEL = "\u0420\u0435\u0433\u0438\u043e\u043d"
COEFFICIENT_LABEL = "\u041a\u043e\u044d\u0444\u0444\u0438\u0446\u0438\u0435\u043d\u0442"
REGION_NAME = "\u041c\u043e\u0441\u043a\u0432\u0430"
ESTIMATE_TITLE = "\u041e\u0421 estimate"
CATALOG_TITLE = "\u041a\u0430\u0442\u0430\u043b\u043e\u0433"

RED_RGB = "FFFFC7CE"
GREY_RGB = "FFD9D9D9"


def _make_catalog_file(path: Path, prices: list[tuple[str, float]]) -> Path:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = CATALOG_TITLE
    for offset, (task_id, price) in enumerate(prices):
        row = 4 + offset
        worksheet.cell(row=row, column=2).value = task_id
        worksheet.cell(row=row, column=3).value = INSTALLATION
        worksheet.cell(row=row, column=4).value = METER
        worksheet.cell(row=row, column=7).value = price
        worksheet.cell(row=row, column=14).value = CODE
    workbook.save(path)
    workbook.close()
    return path


def _make_estimate_file(
    path: Path,
    *,
    base_price: float = 50.0,
    neighbour_value: object = None,
    with_coefficient: float | None = None,
) -> Path:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = ESTIMATE_TITLE
    worksheet.cell(row=7, column=14).value = CODE
    worksheet.cell(row=9, column=3).value = INSTALLATION
    worksheet.cell(row=9, column=4).value = METER
    worksheet.cell(row=9, column=6).value = base_price
    worksheet.cell(row=9, column=14).value = CODE
    if neighbour_value is not None:
        worksheet.cell(row=9, column=7).value = neighbour_value
    if with_coefficient is not None:
        worksheet.cell(row=1, column=1).value = REGION_LABEL
        worksheet.cell(row=1, column=2).value = REGION_NAME
        worksheet.cell(row=2, column=1).value = COEFFICIENT_LABEL
        worksheet.cell(row=2, column=2).value = with_coefficient
    workbook.save(path)
    workbook.close()
    return path


def test_writes_analogs_formula_kr_and_section(tmp_path: Path) -> None:
    catalog = _make_catalog_file(tmp_path / "catalog.xlsx", [("task-1", 100), ("task-2", 120)])
    estimate = _make_estimate_file(tmp_path / "estimate.xlsx")
    output = tmp_path / "out.xlsx"

    outcome = run_and_write(catalog, estimate, output)

    assert outcome.write_report.written_rows == 1
    workbook = load_workbook(output, data_only=False)
    try:
        sheet = workbook[ESTIMATE_TITLE]
        assert sheet.cell(row=9, column=16).value == 100
        assert sheet.cell(row=9, column=17).value == 120
        assert sheet.cell(row=9, column=7).value == (
            "=MAX(F9, IFERROR(AVERAGE(F9, P9:Q9), F9))"
        )
        assert str(sheet.cell(row=9, column=14).value).endswith(KR_END)
        assert sheet.cell(row=9, column=15).value == "01"
        assert RISK_LOG_SHEET not in workbook.sheetnames
    finally:
        workbook.close()


def test_source_file_is_not_modified(tmp_path: Path) -> None:
    catalog = _make_catalog_file(tmp_path / "catalog.xlsx", [("task-1", 100), ("task-2", 120)])
    estimate = _make_estimate_file(tmp_path / "estimate.xlsx")

    run_and_write(catalog, estimate, tmp_path / "out.xlsx")

    workbook = load_workbook(estimate, data_only=False)
    try:
        sheet = workbook[ESTIMATE_TITLE]
        assert sheet.cell(row=9, column=16).value is None
        assert sheet.cell(row=9, column=7).value is None
    finally:
        workbook.close()


def test_default_output_name_gets_wa_suffix(tmp_path: Path) -> None:
    catalog = _make_catalog_file(tmp_path / "catalog.xlsx", [("task-1", 100)])
    estimate = _make_estimate_file(tmp_path / "estimate.xlsx")

    outcome = run_and_write(catalog, estimate)

    assert outcome.output_path.name == "estimate WA.xlsx"
    assert outcome.output_path.exists()


def test_ratio_risk_writes_log_and_colours_red(tmp_path: Path) -> None:
    catalog = _make_catalog_file(tmp_path / "catalog.xlsx", [("task-1", 100), ("task-2", 300)])
    estimate = _make_estimate_file(tmp_path / "estimate.xlsx")
    output = tmp_path / "out.xlsx"

    outcome = run_and_write(catalog, estimate, output)

    assert outcome.write_report.risk_log_rows == 1
    workbook = load_workbook(output, data_only=False)
    try:
        assert RISK_LOG_SHEET in workbook.sheetnames
        log = workbook[RISK_LOG_SHEET]
        assert log.cell(row=2, column=4).value == REASON_RATIO_EXCEEDED
        assert log.cell(row=2, column=5).value == 100
        assert log.cell(row=2, column=6).value == 300

        sheet = workbook[ESTIMATE_TITLE]
        assert sheet.cell(row=9, column=16).fill.start_color.rgb == RED_RGB
        assert sheet.cell(row=9, column=17).fill.start_color.rgb == RED_RGB
    finally:
        workbook.close()


def test_second_price_within_task_is_greyed(tmp_path: Path) -> None:
    catalog = _make_catalog_file(tmp_path / "catalog.xlsx", [("task-1", 100), ("task-1", 150)])
    estimate = _make_estimate_file(tmp_path / "estimate.xlsx")
    output = tmp_path / "out.xlsx"

    run_and_write(catalog, estimate, output)

    workbook = load_workbook(output, data_only=False)
    try:
        sheet = workbook[ESTIMATE_TITLE]
        assert sheet.cell(row=9, column=17).fill.start_color.rgb == GREY_RGB
    finally:
        workbook.close()


def test_regional_coefficient_by_label_scales_analogs(tmp_path: Path) -> None:
    catalog = _make_catalog_file(tmp_path / "catalog.xlsx", [("task-1", 100), ("task-2", 120)])
    estimate = _make_estimate_file(tmp_path / "estimate.xlsx", with_coefficient=1.5)
    output = tmp_path / "out.xlsx"

    outcome = run_and_write(catalog, estimate, output)

    assert outcome.regional_coefficient == 1.5
    assert outcome.coefficient_method == "labeled_region"
    workbook = load_workbook(output, data_only=False)
    try:
        sheet = workbook[ESTIMATE_TITLE]
        assert sheet.cell(row=9, column=16).value == 150
        assert sheet.cell(row=9, column=17).value == 180
    finally:
        workbook.close()


def test_average_column_inserted_when_neighbour_occupied(tmp_path: Path) -> None:
    catalog = _make_catalog_file(tmp_path / "catalog.xlsx", [("task-1", 100), ("task-2", 120)])
    estimate = _make_estimate_file(tmp_path / "estimate.xlsx", neighbour_value=999)
    output = tmp_path / "out.xlsx"

    outcome = run_and_write(catalog, estimate, output)

    assert outcome.write_report.inserted_average_column is True
    assert outcome.write_report.average_column == 7
    assert outcome.write_report.analog_start_column == 17
    workbook = load_workbook(output, data_only=False)
    try:
        sheet = workbook[ESTIMATE_TITLE]
        assert sheet.cell(row=9, column=17).value == 100
        assert sheet.cell(row=9, column=18).value == 120
        assert sheet.cell(row=9, column=7).value == (
            "=MAX(F9, IFERROR(AVERAGE(F9, Q9:R9), F9))"
        )
        assert str(sheet.cell(row=9, column=15).value).endswith(KR_END)
        assert sheet.cell(row=9, column=16).value == "01"
    finally:
        workbook.close()
