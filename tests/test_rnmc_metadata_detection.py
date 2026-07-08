"""Tests for automatic RNMC workbook metadata detection."""

from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook

from app.services.rnmc_excel import analyze_rnmc_zip_row_preview, import_rnmc_zip_catalog_rows
from core.storage import connect, init_database
from core.storage.catalog import get_imported_file, list_imported_files


def test_import_detects_lsr_quarter_and_planned_dates(tmp_path: Path) -> None:
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(zip_path, {"Moscow/metadata.xlsx": _workbook_with_metadata_and_rows()})

    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        result = import_rnmc_zip_catalog_rows(connection, str(zip_path))
        record = next(item for item in list_imported_files(connection) if item.filename == "metadata.xlsx")
        stored = get_imported_file(connection, record.id)
    finally:
        connection.close()

    assert result.success_count == 1
    entry = result.entries[0]
    assert entry.lsr_quarter == "2026 Q1"
    assert entry.planned_start == "2026-01-10"
    assert entry.planned_finish == "2026-02-20"
    assert stored is not None
    assert stored.lsr_quarter == "2026 Q1"
    assert stored.planned_start == "2026-01-10"
    assert stored.planned_finish == "2026-02-20"


def test_row_preview_shows_metadata_without_database_writes(tmp_path: Path) -> None:
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(zip_path, {"Kazan/preview.xlsx": _workbook_with_metadata_and_rows()})

    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        result = analyze_rnmc_zip_row_preview(connection, str(zip_path))
        records = list_imported_files(connection)
    finally:
        connection.close()

    assert records == []
    assert result.preview_ok_count == 1
    entry = result.entries[0]
    assert entry.lsr_quarter == "2026 Q1"
    assert entry.planned_start == "2026-01-10"
    assert entry.planned_finish == "2026-02-20"


def test_metadata_is_detected_even_when_table_is_missing(tmp_path: Path) -> None:
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(zip_path, {"Tula/no-table.xlsx": _workbook_with_metadata_only()})

    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        preview = analyze_rnmc_zip_row_preview(connection, str(zip_path))
        import_result = import_rnmc_zip_catalog_rows(connection, str(zip_path))
        record = next(item for item in list_imported_files(connection) if item.filename == "no-table.xlsx")
    finally:
        connection.close()

    assert preview.no_table_count == 1
    assert preview.entries[0].lsr_quarter == "2026 Q1"
    assert preview.entries[0].planned_start == "2026-01-10"
    assert preview.entries[0].planned_finish == "2026-02-20"
    assert import_result.no_data_count == 1
    assert record.lsr_quarter == "2026 Q1"
    assert record.planned_start == "2026-01-10"
    assert record.planned_finish == "2026-02-20"


def _write_zip(path: Path, files: dict[str, bytes]) -> None:
    with ZipFile(path, "w") as archive:
        for name, data in files.items():
            archive.writestr(name, data)


def _workbook_with_metadata_and_rows() -> bytes:
    workbook = _base_metadata_workbook()
    sheet = workbook.active
    sheet.append([])
    sheet.append([
        "№ п/п",
        "Наименование работ",
        "Ед.изм.",
        "Кол-во",
        "ГЭСН/ФЕР/Перечень",
        "Цена",
    ])
    sheet.append([1, "Work A", "м", 10, "GESN01-01-001-01", 100])
    sheet.append([None, None, None, None, None, None])
    sheet.append([None, None, None, None, None, None])
    sheet.append([None, None, None, None, None, None])
    return _to_bytes(workbook)


def _workbook_with_metadata_only() -> bytes:
    return _to_bytes(_base_metadata_workbook())


def _base_metadata_workbook() -> Workbook:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = "№ задачи 1Ф: TASK-META"
    sheet["A2"] = "Год Квартал ЛСР"
    sheet["B2"] = "1 квартал 2026"
    sheet["A3"] = "Планирумый срок начала работ"
    sheet["B3"] = date(2026, 1, 10)
    sheet["A4"] = "Планируемый срок окончания работ: 20.02.2026"
    return workbook


def _to_bytes(workbook: Workbook) -> bytes:
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def test_metadata_is_detected_from_legacy_note_text(tmp_path: Path) -> None:
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(zip_path, {"Moscow/legacy-note.xlsx": _workbook_with_legacy_note_metadata()})

    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        result = analyze_rnmc_zip_row_preview(connection, str(zip_path))
    finally:
        connection.close()

    entry = result.entries[0]
    assert entry.lsr_quarter == "2025 Q4"
    assert entry.planned_start == "2026-06-01"
    assert entry.planned_finish == "2026-09-01"


def test_metadata_accepts_excel_serial_dates_near_labels(tmp_path: Path) -> None:
    zip_path = tmp_path / "rnmc.zip"
    _write_zip(zip_path, {"Kazan/serial-dates.xlsx": _workbook_with_serial_date_metadata()})

    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        result = analyze_rnmc_zip_row_preview(connection, str(zip_path))
    finally:
        connection.close()

    entry = result.entries[0]
    assert entry.lsr_quarter == "2025 Q2"
    assert entry.planned_start == "2026-07-01"
    assert entry.planned_finish == "2026-09-01"


def _workbook_with_legacy_note_metadata() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = "№ задачи 1Ф: TASK-LEGACY"
    sheet["A2"] = (
        "Расчет выполнен на основании предоставленных смет (в ценах IV кв.2025 г.) "
        "с применением индексов фактической инфляции и дефлятора, рассчитанных "
        "на период выполнения работ – с июня 2026 г. по сентябрь 2026 г."
    )
    sheet.append([])
    sheet.append(["№ п/п", "Наименование работ", "Ед.изм.", "Кол-во"])
    sheet.append([1, "Work A", "м", 10])
    sheet.append([None, None, None, None])
    sheet.append([None, None, None, None])
    sheet.append([None, None, None, None])
    return _to_bytes(workbook)


def _workbook_with_serial_date_metadata() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = "№ задачи 1Ф: TASK-SERIAL"
    sheet["A2"] = "1.1. Предоставленных смет в ценах:"
    sheet["B2"] = "2 кв. 25г."
    sheet["A3"] = "начало работ:"
    sheet["B3"] = 46204
    sheet["A4"] = "окончание работ:"
    sheet["B4"] = 46266
    sheet.append([])
    sheet.append(["№ п/п", "Наименование работ", "Ед.изм.", "Кол-во"])
    sheet.append([1, "Work A", "м", 10])
    sheet.append([None, None, None, None])
    sheet.append([None, None, None, None])
    sheet.append([None, None, None, None])
    return _to_bytes(workbook)
