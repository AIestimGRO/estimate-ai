from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

from core.catalog import _catalog_date_serial
from core.excel_io import read_catalog_rows, read_estimate_rows


CATALOG_SHEET = "\u041a\u0430\u0442\u0430\u043b\u043e\u0433 2026"
ESTIMATE_SHEET = "\u041e\u0421 estimate"
GESN_HEADER = "\u041a\u043e\u0434 \u0413\u042d\u0421\u041d/\u0424\u0415\u0420/\u041f\u0435\u0440\u0435\u0447\u0435\u043d\u044c"
METER = "\u043c"
WORK_NAME = "\u0440\u0430\u0431\u043e\u0442\u0430"
REGION = "\u0440\u0435\u0433\u0438\u043e\u043d"


def fixture_path(tmp_path: Path, name: str) -> Path:
    fixtures_dir = tmp_path / "fixtures"
    fixtures_dir.mkdir()
    return fixtures_dir / name


def save_workbook(workbook: Workbook, path: Path) -> Path:
    workbook.save(path)
    workbook.close()
    return path


def make_catalog_workbook(sheet_name: str = CATALOG_SHEET) -> Workbook:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name

    worksheet.cell(row=3, column=2).value = "task"
    worksheet.cell(row=3, column=3).value = "work name"
    worksheet.cell(row=3, column=4).value = "unit"
    worksheet.cell(row=3, column=7).value = "price"
    worksheet.cell(row=3, column=14).value = "code"
    worksheet.cell(row=3, column=16).value = "region"
    worksheet.cell(row=3, column=17).value = "added date"

    worksheet.cell(row=4, column=2).value = "task-1"
    worksheet.cell(row=4, column=3).value = WORK_NAME
    worksheet.cell(row=4, column=4).value = METER
    worksheet.cell(row=4, column=7).value = 123.45
    worksheet.cell(row=4, column=14).value = "gesn01-01-001-01"
    worksheet.cell(row=4, column=16).value = REGION

    return workbook


def test_read_catalog_rows_finds_catalog_sheet_and_reads_from_row_4(tmp_path: Path) -> None:
    workbook = Workbook()
    workbook.active.title = "Other"
    worksheet = workbook.create_sheet(CATALOG_SHEET)
    worksheet.cell(row=3, column=2).value = "task"
    worksheet.cell(row=4, column=2).value = "task-1"
    worksheet.cell(row=4, column=3).value = WORK_NAME
    worksheet.cell(row=4, column=4).value = METER
    worksheet.cell(row=4, column=7).value = 123.45
    worksheet.cell(row=4, column=14).value = "gesn01-01-001-01"
    worksheet.cell(row=4, column=16).value = REGION

    rows = read_catalog_rows(save_workbook(workbook, fixture_path(tmp_path, "catalog.xlsx")))

    assert len(rows) == 1
    assert rows[0].task_id == "task-1"
    assert rows[0].work_name == WORK_NAME
    assert rows[0].unit == METER
    assert rows[0].price == 123.45
    assert rows[0].code == "gesn01-01-001-01"
    assert rows[0].region == REGION


def test_read_catalog_rows_falls_back_to_first_sheet(tmp_path: Path) -> None:
    workbook = make_catalog_workbook(sheet_name="First")

    rows = read_catalog_rows(save_workbook(workbook, fixture_path(tmp_path, "fallback.xlsx")))

    assert len(rows) == 1
    assert rows[0].task_id == "task-1"


def test_read_estimate_rows_detects_header_row_dynamically(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = ESTIMATE_SHEET
    worksheet.cell(row=7, column=14).value = GESN_HEADER
    worksheet.cell(row=9, column=3).value = WORK_NAME
    worksheet.cell(row=9, column=4).value = METER
    worksheet.cell(row=9, column=6).value = 100
    worksheet.cell(row=9, column=14).value = "gesn01-01-001-01"
    worksheet.cell(row=10, column=3).value = "ignored invalid row"
    worksheet.cell(row=10, column=4).value = METER
    worksheet.cell(row=10, column=6).value = 0
    worksheet.cell(row=10, column=14).value = "gesn01-01-001-02"

    rows = read_estimate_rows(save_workbook(workbook, fixture_path(tmp_path, "estimate.xlsx")))

    assert len(rows) == 1
    assert rows[0].code == "gesn01-01-001-01"
    assert rows[0].unit == METER
    assert rows[0].work_name == WORK_NAME
    assert rows[0].base_price == 100


def test_catalog_excel_date_reads_back_as_datetime_and_converts_to_serial(
    tmp_path: Path,
) -> None:
    workbook = make_catalog_workbook()
    worksheet = workbook[CATALOG_SHEET]
    worksheet.cell(row=4, column=17).value = datetime(2026, 6, 30)
    worksheet.cell(row=4, column=17).number_format = "dd.mm.yyyy"

    rows = read_catalog_rows(save_workbook(workbook, fixture_path(tmp_path, "date.xlsx")))

    assert type(rows[0].added_date) is datetime
    assert _catalog_date_serial(rows[0].added_date) > 0


def test_blank_catalog_added_date_resolves_to_serial_zero_downstream(
    tmp_path: Path,
) -> None:
    workbook = make_catalog_workbook()

    rows = read_catalog_rows(save_workbook(workbook, fixture_path(tmp_path, "blank-date.xlsx")))

    assert rows[0].added_date is None
    assert _catalog_date_serial(rows[0].added_date) == 0
