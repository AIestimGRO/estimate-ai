"""Catalog source and item persistence."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import date, datetime
from numbers import Real
from pathlib import Path
from typing import BinaryIO

from openpyxl import load_workbook

from core.catalog import CatalogRow
from core.excel_io import Settings, read_catalog_rows_with_positions
from core.normalize import NormCode, NormUnit


BATCH_SIZE = 2000
DEFAULT_SOURCE_NAME = "main"
LEGACY_FILE_LOG_SOURCE_NAME = "legacy_file_log"
LEGACY_FILE_LOG_SOURCE_KIND = "legacy_file_log"
STATUS_LEGACY_IMPORTED = "legacy_imported"
STATUS_SUCCESS = "success"
STATUS_SKIPPED = "skipped"
STATUS_NO_DATA = "no_data"
STATUS_FAILED = "failed"
STATUS_DUPLICATE_NAME = "duplicate_name"
STATUS_MANUAL_CHECKED = "manual_checked"
STATUS_PENDING = "pending"
ROW_LOG_STATUS_REJECTED = "rejected"
RNMC_ZIP_SOURCE_NAME = "rnmc_zip_upload"
RNMC_ZIP_SOURCE_KIND = "rnmc_zip"
PROCESSED_FILE_STATUSES = frozenset(
    {
        STATUS_LEGACY_IMPORTED,
        STATUS_SUCCESS,
        STATUS_SKIPPED,
        STATUS_NO_DATA,
        STATUS_DUPLICATE_NAME,
        STATUS_MANUAL_CHECKED,
        STATUS_PENDING,
    }
)
FINAL_PREVIEW_SKIP_STATUSES = frozenset(
    {
        STATUS_LEGACY_IMPORTED,
        STATUS_SUCCESS,
        STATUS_SKIPPED,
        STATUS_NO_DATA,
        STATUS_DUPLICATE_NAME,
        STATUS_MANUAL_CHECKED,
    }
)



@dataclass(frozen=True)
class CatalogSource:
    id: int
    name: str
    kind: str
    created_at: str
    item_count: int


@dataclass(frozen=True)
class ImportedFileRecord:
    id: int
    source_name: str
    source_kind: str
    region_folder: str
    filename: str
    status: str
    imported_at: str
    task_number: str
    rows_ok: int
    rows_rejected: int
    failure_reason: str
    filename_key: str
    legacy_note: str
    lsr_quarter: str
    planned_start: str
    planned_finish: str


@dataclass(frozen=True)
class CatalogItemRecord:
    id: int
    task_id: str
    region: str
    code: str
    unit: str
    work_name: str
    price: float
    total_price: float | None
    labor_unit: float | None
    labor_total: float | None
    machine_labor_unit: float | None
    machine_labor_total: float | None
    source_row_number: int


@dataclass(frozen=True)
class ImportRowLogRecord:
    id: int
    file_id: int
    row_number: int
    status: str
    reason: str


@dataclass(frozen=True)
class LegacyFileLogImportResult:
    rows_seen: int
    rows_imported: int
    duplicates: int
    empty_rows: int


@dataclass(frozen=True)
class CatalogImportResult:
    source_name: str
    source_id: int
    rows_imported: int
    rows_skipped: int
    source_filename: str


@dataclass(frozen=True)
class CatalogRowStorageItem:
    catalog_row: CatalogRow
    source_region_folder: str
    source_filename: str
    source_row_number: int


@dataclass(frozen=True)
class CatalogFileRowsImportResult:
    source_name: str
    source_id: int
    rows_imported: int
    rows_skipped: int
    source_filename: str
    region_folder: str


def list_catalog_sources(connection: sqlite3.Connection) -> list[CatalogSource]:
    rows = connection.execute(
        """
        SELECT
            catalog_sources.id,
            catalog_sources.name,
            catalog_sources.kind,
            catalog_sources.created_at,
            COUNT(catalog_items.id) AS item_count
        FROM catalog_sources
        LEFT JOIN catalog_items ON catalog_items.source_id = catalog_sources.id
        GROUP BY
            catalog_sources.id,
            catalog_sources.name,
            catalog_sources.kind,
            catalog_sources.created_at
        ORDER BY catalog_sources.name
        """
    ).fetchall()
    return [
        CatalogSource(
            id=int(row["id"]),
            name=str(row["name"]),
            kind=str(row["kind"]),
            created_at=str(row["created_at"]),
            item_count=int(row["item_count"]),
        )
        for row in rows
    ]


def list_imported_files(
    connection: sqlite3.Connection,
    *,
    status: str = "",
) -> list[ImportedFileRecord]:
    status_filter = _text(status)
    where = ""
    params: tuple[object, ...] = ()
    if status_filter:
        where = "WHERE imported_files.status = ?"
        params = (status_filter,)
    return _imported_file_records_from_query(connection, where, params)


def _imported_file_records_from_query(
    connection: sqlite3.Connection,
    where: str = "",
    params: tuple[object, ...] = (),
) -> list[ImportedFileRecord]:
    rows = connection.execute(
        """
        SELECT
            imported_files.id,
            COALESCE(catalog_sources.name, '') AS source_name,
            COALESCE(catalog_sources.kind, '') AS source_kind,
            imported_files.region_folder,
            imported_files.filename,
            imported_files.status,
            imported_files.imported_at,
            imported_files.task_number,
            imported_files.rows_ok,
            imported_files.rows_rejected,
            imported_files.failure_reason,
            imported_files.filename_key,
            imported_files.legacy_note,
            imported_files.lsr_quarter,
            imported_files.planned_start,
            imported_files.planned_finish
        FROM imported_files
        LEFT JOIN catalog_sources ON catalog_sources.id = imported_files.source_id
        {where}
        ORDER BY imported_files.imported_at DESC, imported_files.id DESC
        """.format(where=where),
        params,
    ).fetchall()
    return [
        ImportedFileRecord(
            id=int(row["id"]),
            source_name=str(row["source_name"]),
            source_kind=str(row["source_kind"]),
            region_folder=str(row["region_folder"]),
            filename=str(row["filename"]),
            status=str(row["status"]),
            imported_at=str(row["imported_at"]),
            task_number=str(row["task_number"]),
            rows_ok=int(row["rows_ok"]),
            rows_rejected=int(row["rows_rejected"]),
            failure_reason=str(row["failure_reason"]),
            filename_key=str(row["filename_key"]),
            legacy_note=str(row["legacy_note"]),
            lsr_quarter=str(row["lsr_quarter"]),
            planned_start=str(row["planned_start"]),
            planned_finish=str(row["planned_finish"]),
        )
        for row in rows
    ]


def list_catalog_rows(
    connection: sqlite3.Connection,
    *,
    source_name: str = DEFAULT_SOURCE_NAME,
) -> list[CatalogRow]:
    source_id = _require_source_id(connection, source_name)
    rows = connection.execute(
        """
        SELECT
            task_id, region, code, unit, work_name, price, added_date,
            total_price, labor_unit, labor_total,
            machine_labor_unit, machine_labor_total
        FROM catalog_items
        WHERE source_id = ?
        ORDER BY id
        """,
        (source_id,),
    ).fetchall()
    return [_row_to_catalog_row(row) for row in rows]


def count_catalog_rows(
    connection: sqlite3.Connection,
    *,
    source_name: str = DEFAULT_SOURCE_NAME,
) -> int:
    row = connection.execute(
        """
        SELECT COUNT(*) AS row_count
        FROM catalog_items
        INNER JOIN catalog_sources ON catalog_sources.id = catalog_items.source_id
        WHERE catalog_sources.name = ?
        """,
        (source_name,),
    ).fetchone()
    return 0 if row is None else int(row["row_count"])




def import_legacy_file_log(
    connection: sqlite3.Connection,
    workbook_path: str | Path | BinaryIO,
    *,
    source_name: str = LEGACY_FILE_LOG_SOURCE_NAME,
) -> LegacyFileLogImportResult:
    source_id = _get_or_create_source(
        connection,
        source_name,
        kind=LEGACY_FILE_LOG_SOURCE_KIND,
    )
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        sheet = workbook["FileLog"] if "FileLog" in workbook.sheetnames else workbook.active
        headers = _legacy_file_log_headers(sheet)
        seen: dict[str, str] = {}
        rows_seen = 0
        rows_imported = 0
        duplicates = 0
        empty_rows = 0

        for values in sheet.iter_rows(min_row=2, values_only=True):
            row = _legacy_file_log_row(values, headers)
            if _row_is_empty(row):
                empty_rows += 1
                continue
            filename = _text(row.get("filename"))
            if filename == "":
                empty_rows += 1
                continue
            rows_seen += 1

            region = _region_from_folder_text(row.get("region_folder"))
            key = normalize_import_filename(filename)
            legacy_note = _text(row.get("legacy_note"))
            rows_ok = _legacy_rows_ok(legacy_note)
            failure_reason = ""
            status = STATUS_LEGACY_IMPORTED

            previous_region = seen.get(key)
            if previous_region is not None:
                duplicates += 1
                status = STATUS_DUPLICATE_NAME
                failure_reason = f"Duplicate filename in FileLog; first region: {previous_region}"
            else:
                seen[key] = region

            _record_imported_file(
                connection,
                source_id=source_id,
                region_folder=region,
                filename=filename,
                status=status,
                rows_ok=rows_ok,
                rows_rejected=0,
                failure_reason=failure_reason,
                legacy_note=legacy_note,
                lsr_quarter=_text(row.get("lsr_quarter")),
                planned_start=_text(row.get("planned_start")),
                planned_finish=_text(row.get("planned_finish")),
            )
            rows_imported += 1
        connection.commit()
        return LegacyFileLogImportResult(
            rows_seen=rows_seen,
            rows_imported=rows_imported,
            duplicates=duplicates,
            empty_rows=empty_rows,
        )
    finally:
        workbook.close()


def filename_is_processed(connection: sqlite3.Connection, filename: str) -> bool:
    return _filename_has_status(connection, filename, PROCESSED_FILE_STATUSES)


def filename_is_final_for_preview(connection: sqlite3.Connection, filename: str) -> bool:
    """Return True when a file name should not be parsed in row preview.

    Pending and failed files are intentionally not final: a later preview/retry
    can parse them again from a newly uploaded zip.
    """
    return _filename_has_status(connection, filename, FINAL_PREVIEW_SKIP_STATUSES)


def _filename_has_status(
    connection: sqlite3.Connection,
    filename: str,
    statuses: frozenset[str],
) -> bool:
    key = normalize_import_filename(filename)
    if key == "":
        return False
    placeholders = ", ".join("?" for _ in statuses)
    row = connection.execute(
        f"""
        SELECT 1
        FROM imported_files
        WHERE filename_key = ?
          AND status IN ({placeholders})
        LIMIT 1
        """,
        (key, *sorted(statuses)),
    ).fetchone()
    return row is not None


def imported_file_exists_for_region(
    connection: sqlite3.Connection,
    *,
    region_folder: str,
    filename: str,
) -> bool:
    row = connection.execute(
        """
        SELECT 1
        FROM imported_files
        WHERE region_folder = ? AND filename = ?
        LIMIT 1
        """,
        (_text(region_folder), _text(filename)),
    ).fetchone()
    return row is not None


def record_imported_file(
    connection: sqlite3.Connection,
    *,
    region_folder: str,
    filename: str,
    status: str,
    rows_ok: int = 0,
    rows_rejected: int = 0,
    failure_reason: str = "",
    task_number: str = "",
    legacy_note: str = "",
    lsr_quarter: str = "",
    planned_start: str = "",
    planned_finish: str = "",
    source_name: str = RNMC_ZIP_SOURCE_NAME,
    source_kind: str = RNMC_ZIP_SOURCE_KIND,
) -> int:
    source_id = _get_or_create_source(connection, source_name, kind=source_kind)
    return _record_imported_file(
        connection,
        source_id=source_id,
        region_folder=region_folder,
        filename=filename,
        status=status,
        rows_ok=rows_ok,
        rows_rejected=rows_rejected,
        failure_reason=failure_reason,
        task_number=task_number,
        legacy_note=legacy_note,
        lsr_quarter=lsr_quarter,
        planned_start=planned_start,
        planned_finish=planned_finish,
    )


def get_imported_file(connection: sqlite3.Connection, import_id: int) -> ImportedFileRecord | None:
    records = _imported_file_records_from_query(
        connection,
        "WHERE imported_files.id = ?",
        (int(import_id),),
    )
    return records[0] if records else None


def update_imported_file_metadata(
    connection: sqlite3.Connection,
    import_id: int,
    *,
    region_folder: str,
    task_number: str,
    lsr_quarter: str,
    planned_start: str,
    planned_finish: str,
) -> bool:
    cursor = connection.execute(
        """
        UPDATE imported_files
        SET region_folder = ?,
            task_number = ?,
            lsr_quarter = ?,
            planned_start = ?,
            planned_finish = ?
        WHERE id = ?
        """,
        (
            _text(region_folder),
            _text(task_number),
            _text(lsr_quarter),
            _text(planned_start),
            _text(planned_finish),
            int(import_id),
        ),
    )
    connection.commit()
    return cursor.rowcount > 0


def allow_import_retry(connection: sqlite3.Connection, import_id: int) -> bool:
    cursor = connection.execute(
        """
        UPDATE imported_files
        SET status = ?,
            failure_reason = 'Retry allowed; upload the ZIP again to reprocess this file'
        WHERE id = ?
          AND status IN (?, ?)
        """,
        (STATUS_PENDING, int(import_id), STATUS_FAILED, STATUS_NO_DATA),
    )
    connection.commit()
    return cursor.rowcount > 0


def list_catalog_items_for_imported_file(
    connection: sqlite3.Connection,
    import_id: int,
) -> list[CatalogItemRecord]:
    record = get_imported_file(connection, import_id)
    if record is None:
        return []
    rows = connection.execute(
        """
        SELECT
            catalog_items.id,
            catalog_items.task_id,
            catalog_items.region,
            catalog_items.code,
            catalog_items.unit,
            catalog_items.work_name,
            catalog_items.price,
            catalog_items.total_price,
            catalog_items.labor_unit,
            catalog_items.labor_total,
            catalog_items.machine_labor_unit,
            catalog_items.machine_labor_total,
            catalog_items.source_row_number
        FROM catalog_items
        WHERE catalog_items.source_region_folder = ?
          AND catalog_items.source_filename = ?
          AND catalog_items.source_id = (
              SELECT id FROM catalog_sources WHERE name = ? LIMIT 1
          )
        ORDER BY catalog_items.source_row_number, catalog_items.id
        """,
        (record.region_folder, record.filename, RNMC_ZIP_SOURCE_NAME),
    ).fetchall()
    return [
        CatalogItemRecord(
            id=int(row["id"]),
            task_id=str(row["task_id"]),
            region=str(row["region"]),
            code=str(row["code"]),
            unit=str(row["unit"]),
            work_name=str(row["work_name"]),
            price=float(row["price"]),
            total_price=_optional_float(row["total_price"]),
            labor_unit=_optional_float(row["labor_unit"]),
            labor_total=_optional_float(row["labor_total"]),
            machine_labor_unit=_optional_float(row["machine_labor_unit"]),
            machine_labor_total=_optional_float(row["machine_labor_total"]),
            source_row_number=int(row["source_row_number"]),
        )
        for row in rows
    ]


def list_import_row_logs(
    connection: sqlite3.Connection,
    import_id: int,
) -> list[ImportRowLogRecord]:
    rows = connection.execute(
        """
        SELECT id, file_id, row_number, status, reason
        FROM import_row_log
        WHERE file_id = ?
        ORDER BY row_number, id
        """,
        (int(import_id),),
    ).fetchall()
    return [
        ImportRowLogRecord(
            id=int(row["id"]),
            file_id=int(row["file_id"]),
            row_number=int(row["row_number"]),
            status=str(row["status"]),
            reason=str(row["reason"]),
        )
        for row in rows
    ]


def replace_import_row_logs(
    connection: sqlite3.Connection,
    import_id: int,
    row_logs: list[tuple[int, str, str]],
) -> None:
    connection.execute("DELETE FROM import_row_log WHERE file_id = ?", (int(import_id),))
    if row_logs:
        connection.executemany(
            """
            INSERT INTO import_row_log(file_id, row_number, status, reason)
            VALUES (?, ?, ?, ?)
            """,
            [
                (int(import_id), int(row_number), _text(status), _text(reason))
                for row_number, status, reason in row_logs
            ],
        )


def normalize_import_filename(filename: str) -> str:
    return Path(str(filename).replace("\\", "/")).name.strip().casefold()


def replace_catalog_rows_for_file(
    connection: sqlite3.Connection,
    items: list[CatalogRowStorageItem],
    *,
    region_folder: str,
    filename: str,
    source_name: str = RNMC_ZIP_SOURCE_NAME,
    source_kind: str = RNMC_ZIP_SOURCE_KIND,
) -> CatalogFileRowsImportResult:
    source_id = _get_or_create_source(connection, source_name, kind=source_kind)
    region = _text(region_folder)
    source_filename = _text(filename)

    connection.execute(
        """
        DELETE FROM catalog_items
        WHERE source_id = ?
          AND source_region_folder = ?
          AND source_filename = ?
        """,
        (source_id, region, source_filename),
    )

    payload: list[tuple] = []
    skipped = 0
    for item in items:
        row = item.catalog_row
        if not _is_storable_row(row):
            skipped += 1
            continue
        price = _parse_positive_price(row.price)
        if price is None:
            skipped += 1
            continue
        payload.append(
            (
                source_id,
                _text(row.task_id),
                _text(row.region),
                _text(row.code),
                _text(row.unit),
                _text(row.work_name),
                price,
                _parse_optional_number(row.total_price),
                _parse_optional_number(row.labor_unit),
                _parse_optional_number(row.labor_total),
                _parse_optional_number(row.machine_labor_unit),
                _parse_optional_number(row.machine_labor_total),
                _serialize_date(row.added_date),
                _text(item.source_region_folder),
                _text(item.source_filename),
                int(item.source_row_number),
            )
        )

    for offset in range(0, len(payload), BATCH_SIZE):
        connection.executemany(
            """
            INSERT INTO catalog_items (
                source_id, task_id, region, code, unit, work_name, price,
                total_price, labor_unit, labor_total, machine_labor_unit,
                machine_labor_total, added_date, source_region_folder,
                source_filename, source_row_number
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            payload[offset : offset + BATCH_SIZE],
        )

    return CatalogFileRowsImportResult(
        source_name=source_name,
        source_id=source_id,
        rows_imported=len(payload),
        rows_skipped=skipped,
        source_filename=source_filename,
        region_folder=region,
    )


def import_catalog_from_excel(
    connection: sqlite3.Connection,
    workbook_path: str | Path,
    *,
    source_name: str = DEFAULT_SOURCE_NAME,
    replace: bool = True,
    settings: Settings | None = None,
) -> CatalogImportResult:
    path = Path(workbook_path).resolve()
    source_id = _get_or_create_source(
        connection,
        source_name,
        kind="excel_bulk",
    )

    if replace:
        connection.execute("DELETE FROM catalog_items WHERE source_id = ?", (source_id,))

    positioned_rows = read_catalog_rows_with_positions(path, settings)
    payload: list[tuple] = []
    skipped = 0

    for row_number, catalog_row in positioned_rows:
        if not _is_storable_row(catalog_row):
            skipped += 1
            continue
        payload.append(
            (
                source_id,
                _text(catalog_row.task_id),
                _text(catalog_row.region),
                _text(catalog_row.code),
                _text(catalog_row.unit),
                _text(catalog_row.work_name),
                float(catalog_row.price),
                _serialize_date(catalog_row.added_date),
                "",
                path.name,
                row_number,
            )
        )

    for offset in range(0, len(payload), BATCH_SIZE):
        connection.executemany(
            """
            INSERT INTO catalog_items (
                source_id, task_id, region, code, unit, work_name, price,
                added_date, source_region_folder, source_filename, source_row_number
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            payload[offset : offset + BATCH_SIZE],
        )

    _record_imported_file(
        connection,
        source_id=source_id,
        region_folder="",
        filename=path.name,
        status=STATUS_SUCCESS,
        rows_ok=len(payload),
        rows_rejected=skipped,
    )
    connection.commit()

    return CatalogImportResult(
        source_name=source_name,
        source_id=source_id,
        rows_imported=len(payload),
        rows_skipped=skipped,
        source_filename=path.name,
    )


def _legacy_file_log_headers(sheet) -> dict[str, int]:
    raw_headers = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
    normalized = {_normalize_header(value): index for index, value in enumerate(raw_headers)}
    return {
        "region_folder": _find_header(normalized, ["folder"]),
        "filename": _find_header(normalized, ["file"]),
        "legacy_note": _find_header(normalized, ["status"]),
        "lsr_quarter": _find_header(normalized, ["год квартал лср"]),
        "planned_start": _find_header(
            normalized,
            [
                "планирумый срок начала работ",
                "планируемый срок начала работ",
            ],
        ),
        "planned_finish": _find_header(
            normalized,
            ["планируемый срок окончания работ"],
        ),
    }


def _legacy_file_log_row(values: tuple[object, ...], headers: dict[str, int]) -> dict[str, object]:
    result: dict[str, object] = {}
    for name, index in headers.items():
        result[name] = values[index] if 0 <= index < len(values) else None
    return result


def _find_header(normalized_headers: dict[str, int], aliases: list[str]) -> int:
    for alias in aliases:
        if alias in normalized_headers:
            return normalized_headers[alias]
    return -1


def _normalize_header(value: object) -> str:
    return " ".join(_text(value).casefold().split())


def _row_is_empty(row: dict[str, object]) -> bool:
    return all(_text(value) == "" for value in row.values())


def _region_from_folder_text(value: object) -> str:
    text = _text(value)
    if text == "":
        return ""
    parts = [part.strip() for part in text.replace("/", "\\").split("\\") if part.strip()]
    return parts[-1] if parts else text


def _legacy_rows_ok(value: object) -> int:
    if value is None or isinstance(value, bool):
        return 0
    if isinstance(value, Real):
        return max(0, int(value))
    text = _text(value).replace(",", ".")
    try:
        return max(0, int(float(text)))
    except ValueError:
        return 0



def _get_or_create_source(
    connection: sqlite3.Connection,
    name: str,
    *,
    kind: str,
) -> int:
    existing = connection.execute(
        "SELECT id FROM catalog_sources WHERE name = ?",
        (name,),
    ).fetchone()
    if existing is not None:
        return int(existing["id"])

    cursor = connection.execute(
        "INSERT INTO catalog_sources(name, kind) VALUES (?, ?)",
        (name, kind),
    )
    return int(cursor.lastrowid)


def _require_source_id(connection: sqlite3.Connection, source_name: str) -> int:
    row = connection.execute(
        "SELECT id FROM catalog_sources WHERE name = ?",
        (source_name,),
    ).fetchone()
    if row is None:
        raise LookupError(f"catalog source not found: {source_name}")
    return int(row["id"])


def _record_imported_file(
    connection: sqlite3.Connection,
    *,
    source_id: int,
    region_folder: str,
    filename: str,
    status: str,
    rows_ok: int,
    rows_rejected: int,
    failure_reason: str = "",
    task_number: str = "",
    legacy_note: str = "",
    lsr_quarter: str = "",
    planned_start: str = "",
    planned_finish: str = "",
) -> int:
    connection.execute(
        """
        INSERT INTO imported_files (
            source_id, region_folder, filename, status, task_number,
            rows_ok, rows_rejected, failure_reason, filename_key, legacy_note,
            lsr_quarter, planned_start, planned_finish
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(region_folder, filename) DO UPDATE SET
            source_id = excluded.source_id,
            status = excluded.status,
            imported_at = datetime('now'),
            task_number = excluded.task_number,
            rows_ok = excluded.rows_ok,
            rows_rejected = excluded.rows_rejected,
            failure_reason = excluded.failure_reason,
            filename_key = excluded.filename_key,
            legacy_note = excluded.legacy_note,
            lsr_quarter = excluded.lsr_quarter,
            planned_start = excluded.planned_start,
            planned_finish = excluded.planned_finish
        """,
        (
            source_id,
            region_folder,
            filename,
            status,
            task_number,
            rows_ok,
            rows_rejected,
            failure_reason,
            normalize_import_filename(filename),
            legacy_note,
            lsr_quarter,
            planned_start,
            planned_finish,
        ),
    )
    row = connection.execute(
        """
        SELECT id FROM imported_files
        WHERE region_folder = ? AND filename = ?
        LIMIT 1
        """,
        (region_folder, filename),
    ).fetchone()
    if row is None:  # pragma: no cover - sqlite upsert defensive boundary
        raise RuntimeError("imported file row was not recorded")
    return int(row["id"])


def _is_storable_row(catalog_row: CatalogRow) -> bool:
    if _text(catalog_row.task_id) == "":
        return False
    if NormCode(catalog_row.code) == "":
        return False
    if NormUnit(catalog_row.unit) == "":
        return False
    return _parse_positive_price(catalog_row.price) is not None


def _row_to_catalog_row(row: sqlite3.Row) -> CatalogRow:
    added_date = row["added_date"]
    return CatalogRow(
        task_id=row["task_id"],
        price=row["price"],
        code=row["code"],
        unit=row["unit"],
        work_name=row["work_name"],
        region=row["region"],
        added_date=None if added_date is None else str(added_date),
        total_price=row["total_price"],
        labor_unit=row["labor_unit"],
        labor_total=row["labor_total"],
        machine_labor_unit=row["machine_labor_unit"],
        machine_labor_total=row["machine_labor_total"],
    )


def _text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _parse_positive_price(value: object) -> float | None:
    number = _parse_optional_number(value)
    if number is None or number <= 0:
        return None
    return number


def _parse_optional_number(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Real):
        return float(value)
    text = str(value).strip()
    if text == "":
        return None
    text = text.replace(" ", " ").replace(" ", "")
    text = "".join(char for char in text if char.isdigit() or char in {".", ",", "-"})
    if text in {"", ".", ",", "-"}:
        return None
    if "," in text and "." in text:
        comma_pos = text.rfind(",")
        dot_pos = text.rfind(".")
        decimal_sep = "," if comma_pos > dot_pos else "."
        thousands_sep = "." if decimal_sep == "," else ","
        text = text.replace(thousands_sep, "")
        text = text.replace(decimal_sep, ".")
    else:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _optional_float(value: object) -> float | None:
    if value is None:
        return None
    return float(value)


def _serialize_date(value: object) -> str | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    return text or None
