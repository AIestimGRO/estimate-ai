from pathlib import Path

from openpyxl import Workbook, load_workbook

from app.services.run_matching import KR_END
from app.services.write_result import run_and_write
from core.excel_writer import PROBLEM_FILL, TASK_FILL
from core.exclusions import TaskColorEntry
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
BLUE_RGB = "FFDDEBF7"


def _make_catalog_file(
    path: Path,
    prices: list[tuple[str, float]],
    *,
    region: str = "",
) -> Path:
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
        if region:
            worksheet.cell(row=row, column=16).value = region
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
        assert sheet.cell(row=9, column=17).value == 100
        assert sheet.cell(row=9, column=18).value == 120
        assert sheet.cell(row=7, column=17).value == "task-1"
        assert sheet.cell(row=7, column=18).value == "task-2"
        assert sheet.cell(row=9, column=7).value == (
            "=MAX(F9, IFERROR(AVERAGE(F9, Q9:R9), F9))"
        )
        assert sheet.cell(row=9, column=14).value == CODE
        assert str(sheet.cell(row=9, column=15).value).endswith(KR_END)
        assert sheet.cell(row=9, column=16).value == "01"
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


def test_ratio_risk_colours_red_without_price_check_log_sheet(tmp_path: Path) -> None:
    catalog = _make_catalog_file(tmp_path / "catalog.xlsx", [("task-1", 100), ("task-2", 300)])
    estimate = _make_estimate_file(tmp_path / "estimate.xlsx")
    output = tmp_path / "out.xlsx"
    db_path = tmp_path / "risks.db"

    outcome = run_and_write(catalog, estimate, output, database_path=db_path)

    assert outcome.result.flagged_row_count == 1
    workbook = load_workbook(output, data_only=False)
    try:
        assert "Price_Check_Log" not in workbook.sheetnames

        sheet = workbook[ESTIMATE_TITLE]
        assert sheet.cell(row=9, column=17).fill.start_color.rgb == RED_RGB
        assert sheet.cell(row=9, column=18).fill.start_color.rgb == RED_RGB
    finally:
        workbook.close()

    from core.storage import connect, init_database, list_price_risks

    connection = connect(db_path)
    try:
        init_database(connection)
        risks = list_price_risks(connection, status="open")
    finally:
        connection.close()

    assert len(risks) == 1
    assert risks[0].reason == REASON_RATIO_EXCEEDED
    assert risks[0].min_price == 100
    assert risks[0].max_price == 300


def test_second_price_within_task_is_greyed(tmp_path: Path) -> None:
    catalog = _make_catalog_file(tmp_path / "catalog.xlsx", [("task-1", 100), ("task-1", 150)])
    estimate = _make_estimate_file(tmp_path / "estimate.xlsx")
    output = tmp_path / "out.xlsx"

    run_and_write(catalog, estimate, output)

    workbook = load_workbook(output, data_only=False)
    try:
        sheet = workbook[ESTIMATE_TITLE]
        assert sheet.cell(row=9, column=18).fill.start_color.rgb == GREY_RGB
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
        assert sheet.cell(row=9, column=17).value == 150
        assert sheet.cell(row=9, column=18).value == 180
    finally:
        workbook.close()


def test_average_column_inserted_when_neighbour_occupied(tmp_path: Path) -> None:
    catalog = _make_catalog_file(tmp_path / "catalog.xlsx", [("task-1", 100), ("task-2", 120)])
    estimate = _make_estimate_file(tmp_path / "estimate.xlsx", neighbour_value=999)
    output = tmp_path / "out.xlsx"

    outcome = run_and_write(catalog, estimate, output)

    assert outcome.write_report.inserted_average_column is True
    assert outcome.write_report.average_column == 7
    assert outcome.write_report.analog_start_column == 18
    workbook = load_workbook(output, data_only=False)
    try:
        sheet = workbook[ESTIMATE_TITLE]
        assert sheet.cell(row=9, column=18).value == 100
        assert sheet.cell(row=9, column=19).value == 120
        assert sheet.cell(row=9, column=7).value == (
            "=MAX(F9, IFERROR(AVERAGE(F9, R9:S9), F9))"
        )
        assert sheet.cell(row=9, column=15).value == CODE
        assert str(sheet.cell(row=9, column=16).value).endswith(KR_END)
        assert sheet.cell(row=9, column=17).value == "01"
    finally:
        workbook.close()


def test_analog_headers_include_task_number_and_region(tmp_path: Path) -> None:
    catalog = _make_catalog_file(
        tmp_path / "catalog.xlsx",
        [("task-1", 100), ("task-2", 120)],
        region=REGION_NAME,
    )
    estimate = _make_estimate_file(tmp_path / "estimate.xlsx")
    output = tmp_path / "out.xlsx"

    run_and_write(catalog, estimate, output)

    workbook = load_workbook(output, data_only=False)
    try:
        sheet = workbook[ESTIMATE_TITLE]
        assert sheet.cell(row=7, column=17).value == "task-1"
        assert sheet.cell(row=8, column=17).value == REGION_NAME
        assert sheet.cell(row=7, column=18).value == "task-2"
        assert sheet.cell(row=8, column=18).value == REGION_NAME
    finally:
        workbook.close()


def test_task_color_list_tints_analog_column_blue(tmp_path: Path) -> None:
    catalog = _make_catalog_file(tmp_path / "catalog.xlsx", [("task-1", 100), ("task-2", 120)])
    estimate = _make_estimate_file(tmp_path / "estimate.xlsx")
    output = tmp_path / "out.xlsx"
    colors = [TaskColorEntry(enabled=True, task_number="task-1")]

    run_and_write(catalog, estimate, output, task_color_entries=colors)

    workbook = load_workbook(output, data_only=False)
    try:
        sheet = workbook[ESTIMATE_TITLE]
        assert sheet.cell(row=9, column=17).fill.start_color.rgb == BLUE_RGB
        assert sheet.cell(row=9, column=18).fill.start_color.rgb != BLUE_RGB
    finally:
        workbook.close()


def test_kr_and_section_use_dedicated_columns_with_real_headers(tmp_path: Path) -> None:
    code_header = "\u041f\u0435\u0440\u0435\u0447\u0435\u043d\u044c \u0413\u042d\u0421\u041d/\u0424\u0415\u0420/\u0422\u0415\u0420/\u041a\u0420"
    kr_header = "/\u041a\u0420"
    section_header = "\u041a\u043e\u0434 \u0440\u0430\u0437\u0434\u0435\u043b\u0430"
    unit_header = "\u0415\u0434.\u0438\u0437\u043c."
    base_header = "\u0426\u0435\u043d\u0430 \u0435\u0434\u0438\u043d\u0438\u0446\u044b \u0440\u0430\u0431\u043e\u0442, \u0440\u0443\u0431. \u0431\u0435\u0437 \u041d\u0414\u0421"

    catalog = _make_catalog_file(
        tmp_path / "catalog.xlsx",
        [("5818383", 100), ("4768644", 120)],
        region="\u042f\u043a\u0443\u0442\u0438\u044f",
    )
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = ESTIMATE_TITLE
    header_row = 26
    sheet.cell(row=header_row, column=4, value=unit_header)
    sheet.cell(row=header_row, column=6, value=base_header)
    sheet.cell(row=header_row, column=14, value=code_header)
    sheet.cell(row=header_row, column=15, value=kr_header)
    sheet.cell(row=header_row, column=16, value=section_header)
    data_row = 28
    sheet.cell(row=data_row, column=3, value=INSTALLATION)
    sheet.cell(row=data_row, column=4, value=METER)
    sheet.cell(row=data_row, column=6, value=50.0)
    sheet.cell(row=data_row, column=14, value=CODE)
    estimate = tmp_path / "estimate.xlsx"
    workbook.save(estimate)
    workbook.close()

    run_and_write(catalog, estimate, tmp_path / "out.xlsx")

    workbook = load_workbook(tmp_path / "out.xlsx", data_only=False)
    try:
        sheet = workbook[ESTIMATE_TITLE]
        assert sheet.cell(row=data_row, column=14).value == CODE
        assert str(sheet.cell(row=data_row, column=15).value).endswith(KR_END)
        assert sheet.cell(row=data_row, column=16).value == "01"
        assert sheet.cell(row=header_row, column=17).value == "5818383"
        assert sheet.cell(row=header_row + 1, column=17).value == "\u042f\u043a\u0443\u0442\u0438\u044f"
        assert sheet.cell(row=header_row, column=18).value == "4768644"
    finally:
        workbook.close()


def test_kr_reuses_free_column_without_inserting(tmp_path: Path) -> None:
    """eV-grup layout with blank col 15: /KR goes into 15, section stays at 16."""
    code_header = "\u041f\u0435\u0440\u0435\u0447\u0435\u043d\u044c \u0413\u042d\u0421\u041d/\u0424\u0415\u0420/\u0422\u0415\u0420/\u041a\u0420"
    section_header = "\u041a\u043e\u0434 \u0440\u0430\u0437\u0434\u0435\u043b\u0430"
    unit_header = "\u0415\u0434.\u0438\u0437\u043c."
    base_header = "\u0426\u0435\u043d\u0430 \u0435\u0434\u0438\u043d\u0438\u0446\u044b \u0440\u0430\u0431\u043e\u0442, \u0440\u0443\u0431. \u0431\u0435\u0437 \u041d\u0414\u0421"

    catalog = _make_catalog_file(
        tmp_path / "catalog.xlsx",
        [("5818383", 100), ("4768644", 120)],
        region="\u042f\u043a\u0443\u0442\u0438\u044f",
    )
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = ESTIMATE_TITLE
    header_row = 26
    sheet.cell(row=header_row, column=4, value=unit_header)
    sheet.cell(row=header_row, column=6, value=base_header)
    sheet.cell(row=header_row, column=14, value=code_header)
    sheet.cell(row=header_row, column=16, value=section_header)
    data_row = 28
    sheet.cell(row=data_row, column=3, value=INSTALLATION)
    sheet.cell(row=data_row, column=4, value=METER)
    sheet.cell(row=data_row, column=6, value=50.0)
    sheet.cell(row=data_row, column=14, value=CODE)
    estimate = tmp_path / "estimate.xlsx"
    workbook.save(estimate)
    workbook.close()

    outcome = run_and_write(catalog, estimate, tmp_path / "out.xlsx")

    assert outcome.write_report.analog_start_column == 17
    workbook = load_workbook(tmp_path / "out.xlsx", data_only=False)
    try:
        sheet = workbook[ESTIMATE_TITLE]
        assert sheet.cell(row=header_row, column=15).value == "/\u041a\u0420"
        assert sheet.cell(row=header_row, column=16).value == section_header
        assert sheet.cell(row=header_row, column=17).value == "5818383"
        assert str(sheet.cell(row=data_row, column=15).value).endswith(KR_END)
        assert sheet.cell(row=data_row, column=16).value == "01"
        assert sheet.cell(row=data_row, column=17).value == 100
    finally:
        workbook.close()
