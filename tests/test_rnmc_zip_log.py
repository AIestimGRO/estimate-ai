"""Tests for recording RNMC zip upload plans in imported_files."""

from pathlib import Path
from zipfile import ZipFile

from fastapi.testclient import TestClient
from openpyxl import Workbook

from app.services.rnmc_zip import commit_rnmc_zip_import_log
from app.web.app import create_app
from core.storage import connect, init_database
from core.storage.catalog import (
    STATUS_DUPLICATE_NAME,
    STATUS_LEGACY_IMPORTED,
    STATUS_PENDING,
    STATUS_SKIPPED,
    filename_is_processed,
    import_legacy_file_log,
    list_imported_files,
)


def test_commit_rnmc_zip_import_log_records_pending_skipped_and_duplicates(
    tmp_path: Path,
) -> None:
    file_log = tmp_path / "File_Log.xlsx"
    _write_file_log(
        file_log,
        [
            ["Moscow", "old.xlsx", 3, "", "", ""],
            ["Moscow", "legacy.xlsx", 4, "", "", ""],
        ],
    )
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(
        zip_path,
        {
            "Moscow/old.xlsx": b"already logged in same region",
            "Tula/legacy.xlsx": b"already logged by filename in another region",
            "Kazan/new.xlsx": b"new",
            "Samara/new.xlsx": b"duplicate name",
            "readme.txt": b"ignored",
        },
    )

    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        import_legacy_file_log(connection, file_log)
        result = commit_rnmc_zip_import_log(connection, str(zip_path))
        records = list_imported_files(connection)
        assert filename_is_processed(connection, "new.xlsx")
    finally:
        connection.close()

    assert result.pending_recorded == 1
    assert result.skipped_recorded == 1
    assert result.duplicates_recorded == 1
    assert result.existing_records_kept == 1
    assert result.total_recorded == 3
    assert result.dry_run.ignored_files == 1

    by_region_file = {(record.region_folder, record.filename): record for record in records}
    assert by_region_file[("Moscow", "old.xlsx")].status == STATUS_LEGACY_IMPORTED
    assert by_region_file[("Moscow", "old.xlsx")].rows_ok == 3
    assert by_region_file[("Tula", "legacy.xlsx")].status == STATUS_SKIPPED
    assert "already exists" in by_region_file[("Tula", "legacy.xlsx")].failure_reason
    assert by_region_file[("Kazan", "new.xlsx")].status == STATUS_PENDING
    assert by_region_file[("Samara", "new.xlsx")].status == STATUS_DUPLICATE_NAME


def test_admin_can_commit_rnmc_zip_import_log(tmp_path: Path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(zip_path, {"ManualFolder/admin-new.xlsx": b"new"})

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        with zip_path.open("rb") as handle:
            response = client.post(
                "/admin/imports/rnmc-log",
                data={"region_override": "Manual Region"},
                files={"rnmc_zip": ("rnmc.zip", handle, "application/zip")},
            )
        page = client.get("/admin/imports")

    assert response.status_code == 200
    assert "ZIP зафиксирован в журнале" in response.text
    assert "admin-new.xlsx" in page.text
    assert "Manual Region" in page.text
    assert "pending" in page.text


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
