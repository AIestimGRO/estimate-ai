"""Tests for the RNMC import control-center admin workflow."""

from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from fastapi.testclient import TestClient
from openpyxl import Workbook

from app.services.rnmc_excel import import_rnmc_zip_catalog_rows
from app.web.app import create_app
from core.storage import connect, init_database
from core.storage.catalog import (
    STATUS_NO_DATA,
    STATUS_PENDING,
    STATUS_SUCCESS,
    allow_import_retry,
    get_imported_file,
    list_catalog_items_for_imported_file,
    list_import_row_logs,
    list_imported_files,
    update_imported_file_metadata,
)


def test_import_control_center_stores_rejected_rows_and_manual_metadata(
    tmp_path: Path,
) -> None:
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(zip_path, {"Kazan/control.xlsx": _workbook_bytes("TASK-10")})

    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        result = import_rnmc_zip_catalog_rows(connection, str(zip_path))
        records = list_imported_files(connection)
        record = next(item for item in records if item.filename == "control.xlsx")

        assert result.success_count == 1
        assert record.status == STATUS_SUCCESS
        assert record.rows_ok == 2
        assert record.rows_rejected == 2

        catalog_rows = list_catalog_items_for_imported_file(connection, record.id)
        row_logs = list_import_row_logs(connection, record.id)
        assert [row.source_row_number for row in catalog_rows] == [4, 5]
        assert [row.row_number for row in row_logs] == [6, 7]
        assert {row.reason for row in row_logs} == {
            "missing_or_invalid_price",
            "missing_unit_and_quantity",
        }

        assert update_imported_file_metadata(
            connection,
            record.id,
            region_folder="Manual Region",
            task_number="TASK-EDITED",
            lsr_quarter="2026 Q3",
            planned_start="2026-07-01",
            planned_finish="2026-09-30",
        )
        updated = get_imported_file(connection, record.id)
    finally:
        connection.close()

    assert updated is not None
    assert updated.region_folder == "Manual Region"
    assert updated.task_number == "TASK-EDITED"
    assert updated.lsr_quarter == "2026 Q3"
    assert updated.planned_start == "2026-07-01"
    assert updated.planned_finish == "2026-09-30"


def test_allow_retry_for_no_data_makes_file_pending(tmp_path: Path) -> None:
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(zip_path, {"Tula/no-data.xlsx": _workbook_without_valid_rows("TASK-99")})

    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        import_rnmc_zip_catalog_rows(connection, str(zip_path))
        record = next(item for item in list_imported_files(connection) if item.filename == "no-data.xlsx")
        assert record.status == STATUS_NO_DATA

        assert allow_import_retry(connection, record.id)
        updated = get_imported_file(connection, record.id)
    finally:
        connection.close()

    assert updated is not None
    assert updated.status == STATUS_PENDING
    assert "Retry allowed" in updated.failure_reason


def test_admin_import_detail_filters_edits_and_allows_retry(tmp_path: Path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(zip_path, {"Tula/no-data.xlsx": _workbook_without_valid_rows("TASK-99")})

    connection = connect(db_path)
    try:
        init_database(connection)
        import_rnmc_zip_catalog_rows(connection, str(zip_path))
        record = next(item for item in list_imported_files(connection) if item.filename == "no-data.xlsx")
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        filtered = client.get("/admin/imports?status=no_data")
        assert filtered.status_code == 200
        assert "no-data.xlsx" in filtered.text
        assert "Фильтры журнала" in filtered.text

        detail = client.get(f"/admin/imports/{record.id}")
        assert detail.status_code == 200
        assert "Детали импорта" in detail.text
        assert "Ручная правка данных файла" in detail.text
        assert "Разрешить повторную обработку" in detail.text

        update_response = client.post(
            "/admin/imports/update",
            data={
                "import_id": str(record.id),
                "region_folder": "Manual Region",
                "task_number": "TASK-WEB",
                "lsr_quarter": "2026 Q4",
                "planned_start": "2026-10-01",
                "planned_finish": "2026-12-31",
            },
        )
        assert update_response.status_code == 200
        assert "Данные импорта обновлены" in update_response.text
        assert "Manual Region" in update_response.text
        assert "TASK-WEB" in update_response.text

        retry_response = client.post(
            "/admin/imports/allow-retry",
            data={"import_id": str(record.id)},
        )
        assert retry_response.status_code == 200
        assert "Повторная обработка разрешена" in retry_response.text
        assert "pending" in retry_response.text


def _write_zip(path: Path, files: dict[str, bytes]) -> None:
    with ZipFile(path, "w") as archive:
        for name, data in files.items():
            archive.writestr(name, data)


def _workbook_bytes(task_number: str) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = f"№ задачи 1Ф: {task_number}"
    sheet.append([])
    sheet.append([
        "№ п/п",
        "Наименование работ",
        "Ед.изм.",
        "Кол-во",
        "ГЭСН/ФЕР/Перечень",
        "Цена",
    ])
    sheet.append([1, "Work A", "м", 10, "GESN01-01-001-01", 100.5])
    sheet.append([2, "Work B", "шт", 5, "GESN02-01-001-01", "1 234,56"])
    sheet.append([3, "Bad price", "м", 1, "GESN03-01-001-01", ""])
    sheet.append([4, "Work without unit and qty", "", "", "GESN04-01-001-01", 10])
    sheet.append([None, None, None, None, None, None])
    sheet.append([None, None, None, None, None, None])
    sheet.append([None, None, None, None, None, None])
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def _workbook_without_valid_rows(task_number: str) -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = f"№ задачи 1Ф: {task_number}"
    sheet.append([])
    sheet.append(["№ п/п", "Наименование работ", "Ед.изм.", "Кол-во", "Цена"])
    sheet.append([1, "No code", "м", 1, 100])
    sheet.append([None, None, None, None, None])
    sheet.append([None, None, None, None, None])
    sheet.append([None, None, None, None, None])
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()
