"""Tests for RNMC zip workbook row previews."""

from pathlib import Path
from zipfile import ZipFile

from fastapi.testclient import TestClient
from openpyxl import Workbook

from app.services.rnmc_excel import (
    STATUS_DUPLICATE_NAME,
    STATUS_PREVIEW_OK,
    STATUS_SKIPPED_PROCESSED,
    analyze_rnmc_zip_row_preview,
)
from app.web.app import create_app
from core.storage import connect, init_database
from core.storage.catalog import STATUS_PENDING, import_legacy_file_log, record_imported_file


def test_rnmc_zip_row_preview_counts_rows_and_keeps_pending_previewable(tmp_path: Path) -> None:
    file_log = tmp_path / "File_Log.xlsx"
    _write_file_log(file_log, [["Moscow", "old.xlsx", 3, "", "", ""]])
    valid = _workbook_bytes(task_number="TASK-42")
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(
        zip_path,
        {
            "Moscow/old.xlsx": valid,
            "Kazan/new.xlsx": valid,
            "Samara/new.xlsx": valid,
            "Tula/pending.xlsx": valid,
            "notes/readme.txt": b"ignored",
        },
    )

    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        import_legacy_file_log(connection, file_log)
        record_imported_file(
            connection,
            region_folder="Tula",
            filename="pending.xlsx",
            status=STATUS_PENDING,
        )
        connection.commit()
        result = analyze_rnmc_zip_row_preview(connection, str(zip_path))
    finally:
        connection.close()

    by_path = {entry.archive_path: entry for entry in result.entries}
    assert result.total_excel_files == 4
    assert result.preview_ok_count == 2
    assert result.rows_ok_total == 4
    assert result.rows_rejected_total == 2
    assert result.skipped_processed_count == 1
    assert result.duplicate_name_count == 1
    assert result.ignored_files == 1

    assert by_path["Moscow/old.xlsx"].status == STATUS_SKIPPED_PROCESSED
    assert by_path["Kazan/new.xlsx"].status == STATUS_PREVIEW_OK
    assert by_path["Kazan/new.xlsx"].region_folder == "Kazan"
    assert by_path["Kazan/new.xlsx"].sheet_name == "Data"
    assert by_path["Kazan/new.xlsx"].header_row == 3
    assert by_path["Kazan/new.xlsx"].task_number == "TASK-42"
    assert by_path["Kazan/new.xlsx"].rows_ok == 2
    assert by_path["Kazan/new.xlsx"].rows_rejected == 1
    assert by_path["Kazan/new.xlsx"].sample_rows[0].work_name == "Work A"
    assert by_path["Samara/new.xlsx"].status == STATUS_DUPLICATE_NAME
    assert by_path["Tula/pending.xlsx"].status == STATUS_PREVIEW_OK


def test_admin_can_preview_rnmc_zip_rows(tmp_path: Path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(zip_path, {"Folder/admin-new.xlsx": _workbook_bytes(task_number="ADMIN-7")})

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        with zip_path.open("rb") as handle:
            response = client.post(
                "/admin/imports/rnmc-row-preview",
                data={"region_override": "Manual Region"},
                files={"rnmc_zip": ("rnmc.zip", handle, "application/zip")},
            )

    assert response.status_code == 200
    assert "Предпросмотр строк РНМЦ" in response.text
    assert "admin-new.xlsx" in response.text
    assert "Manual Region" in response.text
    assert "ADMIN-7" in response.text
    assert "preview_ok" in response.text
    assert "OK-строк найдено" in response.text


def _write_zip(path: Path, files: dict[str, bytes]) -> None:
    with ZipFile(path, "w") as archive:
        for name, data in files.items():
            archive.writestr(name, data)


def _write_file_log(path: Path, rows: list[list[object]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "FileLog"
    sheet.append(
        [
            "Folder",
            "File",
            "Status",
            "Год Квартал ЛСР",
            "Планирумый срок начала работ",
            "Планируемый срок окончания работ",
        ]
    )
    for row in rows:
        sheet.append(row)
    workbook.save(path)


def _workbook_bytes(*, task_number: str) -> bytes:
    from io import BytesIO

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = f"№ задачи 1Ф: {task_number}"
    sheet.append([])
    sheet.append(["№ п/п", "Наименование работ", "Ед.изм.", "Кол-во"])
    sheet.append([1, "Work A", "м", 10])
    sheet.append([2, "Work B", "шт", 5])
    sheet.append([3, "Work without unit and qty", "", ""])
    sheet.append([None, None, None, None])
    sheet.append([None, None, None, None])
    sheet.append([None, None, None, None])
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()
