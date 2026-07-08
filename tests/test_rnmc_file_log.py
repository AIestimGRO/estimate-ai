"""Tests for RNMC legacy FileLog import tracking."""

from pathlib import Path

from fastapi.testclient import TestClient
from openpyxl import Workbook

from app.web.app import create_app
from core.storage import connect, init_database
from core.storage.catalog import (
    STATUS_DUPLICATE_NAME,
    STATUS_LEGACY_IMPORTED,
    filename_is_processed,
    import_legacy_file_log,
    list_imported_files,
    normalize_import_filename,
)


def test_import_legacy_file_log_records_rows_and_duplicate_names(tmp_path: Path) -> None:
    workbook_path = tmp_path / "File_Log.xlsx"
    _write_file_log(
        workbook_path,
        [
            ["C:/Root/Moscow", "rnmc-1.xlsx", 25, "2026 Q1", "2026-01-10", "2026-02-20"],
            ["Tula", "rnmc-2.xlsx", "нет данных", "2026 Q2", "", ""],
            ["Kazan", "rnmc-1.xlsx", "новая РНМЦ", "2026 Q3", "", ""],
        ],
    )
    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        result = import_legacy_file_log(connection, workbook_path)
        records = list_imported_files(connection)
    finally:
        connection.close()

    assert result.rows_seen == 3
    assert result.rows_imported == 3
    assert result.duplicates == 1
    by_region = {record.region_folder: record for record in records}
    assert by_region["Moscow"].status == STATUS_LEGACY_IMPORTED
    assert by_region["Moscow"].rows_ok == 25
    assert by_region["Moscow"].legacy_note == "25"
    assert by_region["Moscow"].lsr_quarter == "2026 Q1"
    assert by_region["Moscow"].planned_start == "2026-01-10"
    assert by_region["Moscow"].planned_finish == "2026-02-20"
    assert by_region["Kazan"].status == STATUS_DUPLICATE_NAME
    assert "Duplicate filename" in by_region["Kazan"].failure_reason


def test_filename_is_processed_uses_file_name_only(tmp_path: Path) -> None:
    workbook_path = tmp_path / "File_Log.xlsx"
    _write_file_log(workbook_path, [["Moscow", "RNMC-1.xlsx", 1, "", "", ""]])
    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        import_legacy_file_log(connection, workbook_path)
        assert filename_is_processed(connection, "RNMC-1.xlsx")
        assert filename_is_processed(connection, "Other/Folder/rnmc-1.xlsx")
        assert not filename_is_processed(connection, "rnmc-2.xlsx")
    finally:
        connection.close()


def test_normalize_import_filename_keeps_basename_case_insensitive() -> None:
    assert normalize_import_filename("Folder\\RNMC-1.XLSX") == "rnmc-1.xlsx"


def test_admin_can_upload_legacy_file_log(tmp_path: Path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    workbook_path = tmp_path / "File_Log.xlsx"
    _write_file_log(workbook_path, [["Moscow", "rnmc-admin.xlsx", 7, "2026 Q1", "", ""]])

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        with workbook_path.open("rb") as handle:
            response = client.post(
                "/admin/imports/file-log",
                files={
                    "file_log": (
                        "File_Log.xlsx",
                        handle,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                },
            )
        page = client.get("/admin/imports")

    assert response.status_code == 200
    assert "rnmc-admin.xlsx" in page.text
    assert "legacy_imported" in page.text
    assert "2026 Q1" in page.text


def test_schema_migration_adds_import_log_columns(tmp_path: Path) -> None:
    db_path = tmp_path / "estimate_ai.db"
    connection = connect(db_path)
    try:
        connection.executescript(
            """
            CREATE TABLE schema_migrations (version INTEGER PRIMARY KEY);
            INSERT INTO schema_migrations(version) VALUES (2);
            CREATE TABLE catalog_sources (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                kind TEXT NOT NULL DEFAULT 'excel_bulk',
                created_at TEXT NOT NULL DEFAULT (datetime('now'))
            );
            CREATE TABLE imported_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_id INTEGER REFERENCES catalog_sources(id) ON DELETE SET NULL,
                region_folder TEXT NOT NULL DEFAULT '',
                filename TEXT NOT NULL,
                status TEXT NOT NULL,
                imported_at TEXT NOT NULL DEFAULT (datetime('now')),
                task_number TEXT NOT NULL DEFAULT '',
                rows_ok INTEGER NOT NULL DEFAULT 0,
                rows_rejected INTEGER NOT NULL DEFAULT 0,
                failure_reason TEXT NOT NULL DEFAULT '',
                UNIQUE(region_folder, filename)
            );
            INSERT INTO imported_files(region_folder, filename, status)
            VALUES ('Moscow', 'RNMC.xlsx', 'success');
            """
        )
        init_database(connection)
        columns = {
            str(row["name"])
            for row in connection.execute("PRAGMA table_info(imported_files)")
        }
        record = connection.execute(
            "SELECT filename_key FROM imported_files WHERE filename = ?",
            ("RNMC.xlsx",),
        ).fetchone()
    finally:
        connection.close()

    assert "filename_key" in columns
    assert "legacy_note" in columns
    assert "lsr_quarter" in columns
    assert record is not None
    assert record["filename_key"] == "rnmc.xlsx"


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
