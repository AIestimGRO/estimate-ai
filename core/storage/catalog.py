"""Catalog source and item persistence."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import date, datetime
from numbers import Real
from pathlib import Path

from core.catalog import CatalogRow
from core.excel_io import Settings, read_catalog_rows_with_positions
from core.normalize import NormCode, NormUnit


BATCH_SIZE = 2000
DEFAULT_SOURCE_NAME = "main"


@dataclass(frozen=True)
class CatalogSource:
    id: int
    name: str
    kind: str
    created_at: str


@dataclass(frozen=True)
class CatalogImportResult:
    source_name: str
    source_id: int
    rows_imported: int
    rows_skipped: int
    source_filename: str


def list_catalog_sources(connection: sqlite3.Connection) -> list[CatalogSource]:
    rows = connection.execute(
        "SELECT id, name, kind, created_at FROM catalog_sources ORDER BY name"
    ).fetchall()
    return [
        CatalogSource(
            id=int(row["id"]),
            name=str(row["name"]),
            kind=str(row["kind"]),
            created_at=str(row["created_at"]),
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
        SELECT task_id, region, code, unit, work_name, price, added_date
        FROM catalog_items
        WHERE source_id = ?
        ORDER BY id
        """,
        (source_id,),
    ).fetchall()
    return [_row_to_catalog_row(row) for row in rows]


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
        status="success",
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
) -> None:
    connection.execute(
        """
        INSERT INTO imported_files (
            source_id, region_folder, filename, status, task_number,
            rows_ok, rows_rejected, failure_reason
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(region_folder, filename) DO UPDATE SET
            source_id = excluded.source_id,
            status = excluded.status,
            imported_at = datetime('now'),
            task_number = excluded.task_number,
            rows_ok = excluded.rows_ok,
            rows_rejected = excluded.rows_rejected,
            failure_reason = excluded.failure_reason
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
        ),
    )


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
    )


def _text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _parse_positive_price(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Real):
        number = float(value)
    else:
        text = str(value).strip()
        if text == "":
            return None
        try:
            number = float(text)
        except ValueError:
            return None
    if number <= 0:
        return None
    return number


def _serialize_date(value: object) -> str | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    return text or None
