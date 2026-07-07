"""Price risk log and GESN exception persistence.

Ports the Price_Check_Log / GESN_Exceptions workflow from Module6,
DOMAIN_RULES.md section 5.2. Open risks are deduplicated by
exception_key (unit||code||dem); repeat runs update metrics instead of
inserting duplicates.
"""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from pathlib import Path

from core.approval import ApproveGesnExceptionRange
from core.risk import GesnException

STATUS_OPEN = "open"
STATUS_APPROVED = "approved"


@dataclass(frozen=True)
class FlaggedRiskSnapshot:
    """Minimal flagged-row payload for price_risk_log upsert."""

    exception_key: str
    reason: str
    code: str
    unit: str
    min_price: float | None
    max_price: float | None
    ratio: float | None
    recommended_price: float | None
    estimate_row: int


@dataclass(frozen=True)
class PriceRiskLogEntry:
    """One row in price_risk_log."""

    id: int
    exception_key: str
    status: str
    reason: str
    code: str
    unit: str
    min_price: float | None
    max_price: float | None
    ratio: float | None
    recommended_price: float | None
    estimate_row: int | None
    first_seen_at: str
    last_seen_at: str


def load_gesn_exceptions(connection: sqlite3.Connection) -> dict[str, GesnException]:
    return {
        item.exception_key: item
        for item in list_gesn_exceptions(connection)
    }


def list_gesn_exceptions(connection: sqlite3.Connection) -> list[GesnException]:
    rows = connection.execute(
        """
        SELECT exception_key, approved_min, approved_max, last_range_update_date
        FROM gesn_exceptions
        ORDER BY exception_key
        """
    ).fetchall()
    return [
        GesnException(
            exception_key=str(row["exception_key"]),
            approved_min=float(row["approved_min"]),
            approved_max=float(row["approved_max"]),
            last_range_update_date=float(row["last_range_update_date"]),
        )
        for row in rows
    ]


def save_gesn_exception(connection: sqlite3.Connection, exception: GesnException) -> None:
    connection.execute(
        """
        INSERT INTO gesn_exceptions (
            exception_key, approved_min, approved_max, last_range_update_date
        ) VALUES (?, ?, ?, ?)
        ON CONFLICT(exception_key) DO UPDATE SET
            approved_min = excluded.approved_min,
            approved_max = excluded.approved_max,
            last_range_update_date = excluded.last_range_update_date
        """,
        (
            exception.exception_key,
            exception.approved_min,
            exception.approved_max,
            exception.last_range_update_date,
        ),
    )
    connection.commit()


def list_price_risks(
    connection: sqlite3.Connection,
    *,
    status: str | None = None,
) -> list[PriceRiskLogEntry]:
    if status is None:
        rows = connection.execute(
            """
            SELECT id, exception_key, status, reason, code, unit,
                   min_price, max_price, ratio, recommended_price,
                   estimate_row, first_seen_at, last_seen_at
            FROM price_risk_log
            ORDER BY last_seen_at DESC, id DESC
            """
        ).fetchall()
    else:
        rows = connection.execute(
            """
            SELECT id, exception_key, status, reason, code, unit,
                   min_price, max_price, ratio, recommended_price,
                   estimate_row, first_seen_at, last_seen_at
            FROM price_risk_log
            WHERE status = ?
            ORDER BY last_seen_at DESC, id DESC
            """,
            (status,),
        ).fetchall()
    return [_row_to_entry(row) for row in rows]


def upsert_open_risk(
    connection: sqlite3.Connection,
    *,
    exception_key: str,
    reason: str,
    code: str,
    unit: str,
    min_price: float | None,
    max_price: float | None,
    ratio: float | None,
    recommended_price: float | None,
    estimate_row: int | None,
) -> None:
    """Insert or update the single open risk row for exception_key."""
    existing = connection.execute(
        "SELECT id, status FROM price_risk_log WHERE exception_key = ?",
        (exception_key,),
    ).fetchone()

    if existing is None:
        connection.execute(
            """
            INSERT INTO price_risk_log (
                exception_key, status, reason, code, unit,
                min_price, max_price, ratio, recommended_price, estimate_row
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                exception_key,
                STATUS_OPEN,
                reason,
                code,
                unit,
                min_price,
                max_price,
                ratio,
                recommended_price,
                estimate_row,
            ),
        )
    else:
        connection.execute(
            """
            UPDATE price_risk_log
            SET status = ?,
                reason = ?,
                code = ?,
                unit = ?,
                min_price = ?,
                max_price = ?,
                ratio = ?,
                recommended_price = ?,
                estimate_row = ?,
                last_seen_at = datetime('now')
            WHERE exception_key = ?
            """,
            (
                STATUS_OPEN,
                reason,
                code,
                unit,
                min_price,
                max_price,
                ratio,
                recommended_price,
                estimate_row,
                exception_key,
            ),
        )
    connection.commit()


def approve_risk(
    connection: sqlite3.Connection,
    exception_key: str,
    proposed_min: float,
    proposed_max: float,
    proposed_date_serial: float,
) -> GesnException:
    """Approve an open risk: widen gesn_exceptions and mark the log row approved."""
    row = connection.execute(
        """
        SELECT id, status FROM price_risk_log
        WHERE exception_key = ? AND status = ?
        """,
        (exception_key, STATUS_OPEN),
    ).fetchone()
    if row is None:
        raise ValueError(f"no open risk log entry for exception_key {exception_key!r}")

    existing = load_gesn_exceptions(connection).get(exception_key)
    approved = ApproveGesnExceptionRange(
        exception_key=exception_key,
        proposed_min=proposed_min,
        proposed_max=proposed_max,
        proposed_date_serial=proposed_date_serial,
        existing_exception=existing,
    )
    save_gesn_exception(connection, approved)
    connection.execute(
        """
        UPDATE price_risk_log
        SET status = ?, last_seen_at = datetime('now')
        WHERE exception_key = ?
        """,
        (STATUS_APPROVED, exception_key),
    )
    connection.commit()
    return approved


def persist_flagged_risks(
    connection: sqlite3.Connection,
    snapshots: list[FlaggedRiskSnapshot],
) -> int:
    """Upsert open risk rows for every flagged estimate row."""
    for snapshot in snapshots:
        upsert_open_risk(
            connection,
            exception_key=snapshot.exception_key,
            reason=snapshot.reason,
            code=snapshot.code,
            unit=snapshot.unit,
            min_price=snapshot.min_price,
            max_price=snapshot.max_price,
            ratio=snapshot.ratio,
            recommended_price=snapshot.recommended_price,
            estimate_row=snapshot.estimate_row,
        )
    return len(snapshots)


def database_is_available(database_path: str | Path | None) -> bool:
    """Return whether run_and_write should use SQLite for gesn/risk persistence."""
    if database_path is not None:
        return True

    from core.storage.connection import default_database_path

    return default_database_path().is_file()


def _row_to_entry(row: sqlite3.Row) -> PriceRiskLogEntry:
    return PriceRiskLogEntry(
        id=int(row["id"]),
        exception_key=str(row["exception_key"]),
        status=str(row["status"]),
        reason=str(row["reason"]),
        code=str(row["code"]),
        unit=str(row["unit"]),
        min_price=_optional_float(row["min_price"]),
        max_price=_optional_float(row["max_price"]),
        ratio=_optional_float(row["ratio"]),
        recommended_price=_optional_float(row["recommended_price"]),
        estimate_row=_optional_int(row["estimate_row"]),
        first_seen_at=str(row["first_seen_at"]),
        last_seen_at=str(row["last_seen_at"]),
    )


def _optional_float(value: object) -> float | None:
    if value is None:
        return None
    return float(value)


def _optional_int(value: object) -> int | None:
    if value is None:
        return None
    return int(value)
