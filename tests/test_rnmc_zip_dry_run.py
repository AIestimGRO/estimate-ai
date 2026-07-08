"""Tests for RNMC zip dry-run planning."""

from pathlib import Path
from zipfile import ZipFile

from fastapi.testclient import TestClient
from openpyxl import Workbook

from app.services.rnmc_zip import (
    STATUS_DUPLICATE_NAME,
    STATUS_SKIPPED_PROCESSED,
    STATUS_WILL_PROCESS,
    analyze_rnmc_zip_dry_run,
)
from app.web.app import create_app
from core.storage import connect, init_database
from core.storage.catalog import import_legacy_file_log


def test_rnmc_zip_dry_run_marks_processed_and_duplicate_names(tmp_path: Path) -> None:
    db_path = tmp_path / "estimate_ai.db"
    file_log = tmp_path / "File_Log.xlsx"
    _write_file_log(file_log, [["Moscow", "old.xlsx", 3, "", "", ""]])
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(
        zip_path,
        {
            "Moscow/old.xlsx": b"old",
            "Kazan/new.xlsx": b"new",
            "Tula/new.xlsx": b"duplicate",
            "notes/readme.txt": b"ignore",
        },
    )

    connection = connect(db_path)
    try:
        init_database(connection)
        import_legacy_file_log(connection, file_log)
        result = analyze_rnmc_zip_dry_run(connection, str(zip_path))
    finally:
        connection.close()

    by_path = {entry.archive_path: entry for entry in result.entries}
    assert result.total_excel_files == 3
    assert result.will_process_count == 1
    assert result.skipped_processed_count == 1
    assert result.duplicate_name_count == 1
    assert result.ignored_files == 1
    assert by_path["Moscow/old.xlsx"].status == STATUS_SKIPPED_PROCESSED
    assert by_path["Kazan/new.xlsx"].region_folder == "Kazan"
    assert by_path["Kazan/new.xlsx"].status == STATUS_WILL_PROCESS
    assert by_path["Tula/new.xlsx"].status == STATUS_DUPLICATE_NAME


def test_rnmc_zip_dry_run_can_override_region(tmp_path: Path) -> None:
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(zip_path, {"AnyFolder/new.xlsx": b"new"})
    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        result = analyze_rnmc_zip_dry_run(
            connection,
            str(zip_path),
            region_override="Manual Region",
        )
    finally:
        connection.close()

    assert result.entries[0].region_folder == "Manual Region"
    assert result.entries[0].status == STATUS_WILL_PROCESS


def test_admin_can_run_rnmc_zip_dry_run(tmp_path: Path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    file_log = tmp_path / "File_Log.xlsx"
    _write_file_log(file_log, [["Moscow", "old.xlsx", 3, "", "", ""]])
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(zip_path, {"Moscow/old.xlsx": b"old", "Tula/new.xlsx": b"new"})

    connection = connect(db_path)
    try:
        init_database(connection)
        import_legacy_file_log(connection, file_log)
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        with zip_path.open("rb") as handle:
            response = client.post(
                "/admin/imports/rnmc-dry-run",
                data={"region_override": ""},
                files={"rnmc_zip": ("rnmc.zip", handle, "application/zip")},
            )

    assert response.status_code == 200
    assert "Результат dry-run ZIP" in response.text
    assert "old.xlsx" in response.text
    assert "skipped_processed" in response.text
    assert "new.xlsx" in response.text
    assert "will_process" in response.text


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
