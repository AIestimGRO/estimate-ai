"""Approval workflow and audit log for durable catalog corrections."""

from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from pathlib import Path

from core.storage.catalog import (
    CATALOG_EDITOR_NUMERIC_FIELDS,
    CATALOG_EDITOR_TEXT_FIELDS,
    REQUIRED_CATALOG_NUMERIC_FIELDS,
    _clean_catalog_editor_values,
    _parse_catalog_editor_number,
)


ACTION_UPDATE = "update"
ACTION_DELETE = "delete"
STATUS_PENDING = "pending"
STATUS_APPROVED = "approved"
STATUS_REJECTED = "rejected"
ROLE_SPECIALIST = "specialist"
ROLE_SENIOR = "senior"
ROLE_ADMIN = "admin"
APPROVER_ROLES = frozenset({ROLE_SENIOR, ROLE_ADMIN})
SUPPORTED_ACTIONS = frozenset({ACTION_UPDATE, ACTION_DELETE})
SUPPORTED_ROLES = frozenset({ROLE_SPECIALIST, ROLE_SENIOR, ROLE_ADMIN})
VALUE_TEXT = "text"
VALUE_NUMBER = "number"
VALUE_NULL = "null"

_CONFIG_PATH = (
    Path(__file__).resolve().parents[2]
    / "data"
    / "config"
    / "expert_corrections.json"
)


@dataclass(frozen=True)
class CatalogCorrectionChange:
    field_name: str
    old_value: object
    new_value: object
    value_type: str


@dataclass(frozen=True)
class CatalogCorrectionEvent:
    event_type: str
    actor: str
    actor_role: str
    details: str
    created_at: str


@dataclass(frozen=True)
class CatalogCorrectionRecord:
    id: int
    seed_key: str
    action: str
    status: str
    target_item_id: int | None
    target_source_name: str
    target_source_filename: str
    target_source_row_number: int
    target_task_id: str
    target_code: str
    target_unit: str
    target_work_name: str
    reason: str
    submitted_by: str
    submitted_role: str
    submitted_at: str
    reviewed_by: str
    reviewed_role: str
    reviewed_at: str
    review_comment: str
    applied_at: str
    changes: list[CatalogCorrectionChange]
    events: list[CatalogCorrectionEvent]


def create_catalog_correction(
    connection: sqlite3.Connection,
    item_id: int,
    *,
    values: dict[str, object],
    reason: str,
    actor: str,
    actor_role: str,
    action: str = ACTION_UPDATE,
) -> int:
    """Create a pending request without mutating catalog_items."""
    normalized_action = _validate_action(action)
    normalized_reason = _required_text(reason, "Correction reason is required")
    normalized_actor = _required_text(actor, "Correction actor is required")
    normalized_role = _validate_role(actor_role)
    target = _catalog_target(connection, int(item_id))
    if target is None:
        raise ValueError("Catalog item was not found")
    if _pending_request_exists(connection, target, normalized_action):
        raise ValueError("A pending correction already exists for this catalog row")

    changes: list[tuple[str, object, object]] = []
    if normalized_action == ACTION_UPDATE:
        cleaned = _clean_catalog_editor_values(values)
        for field_name, new_value in cleaned.items():
            old_value = target[field_name]
            if not _values_equal(old_value, new_value):
                changes.append((field_name, old_value, new_value))
        if not changes:
            raise ValueError("The correction does not change any catalog fields")

    with connection:
        correction_id = _insert_request(
            connection,
            target=target,
            action=normalized_action,
            status=STATUS_PENDING,
            reason=normalized_reason,
            submitted_by=normalized_actor,
            submitted_role=normalized_role,
        )
        _insert_changes(connection, correction_id, changes)
        _insert_event(
            connection,
            correction_id,
            "submitted",
            normalized_actor,
            normalized_role,
            normalized_reason,
        )
    return correction_id


def create_bulk_catalog_corrections(
    connection: sqlite3.Connection,
    item_ids: list[int],
    *,
    action: str,
    reason: str,
    actor: str,
    actor_role: str,
    field: str = "",
    operation: str = "",
    value: object = None,
) -> int:
    """Create one auditable request per selected catalog row."""
    ids = sorted({int(item_id) for item_id in item_ids if int(item_id) > 0})
    if not ids:
        return 0
    normalized_action = _validate_action(action)
    created = 0
    for item_id in ids:
        values: dict[str, object] = {}
        if normalized_action == ACTION_UPDATE:
            target = _catalog_target(connection, item_id)
            if target is None:
                continue
            current_value = target[field] if field in target.keys() else None
            new_value = _bulk_result_value(
                field,
                operation,
                value,
                current_value,
            )
            if _values_equal(current_value, new_value):
                continue
            values = {field: new_value}
        try:
            create_catalog_correction(
                connection,
                item_id,
                values=values,
                reason=reason,
                actor=actor,
                actor_role=actor_role,
                action=normalized_action,
            )
        except ValueError as error:
            if "pending correction already exists" in str(error):
                continue
            raise
        created += 1
    return created


def approve_catalog_correction(
    connection: sqlite3.Connection,
    correction_id: int,
    *,
    actor: str,
    actor_role: str,
    comment: str = "",
) -> None:
    normalized_actor = _required_text(actor, "Reviewer is required")
    normalized_role = _validate_role(actor_role)
    if normalized_role not in APPROVER_ROLES:
        raise PermissionError("Only a senior or admin can approve corrections")
    request = _request_row(connection, int(correction_id))
    if request is None or str(request["status"]) != STATUS_PENDING:
        raise ValueError("Pending correction was not found")

    with connection:
        connection.execute(
            """
            UPDATE catalog_correction_requests
            SET status = ?, reviewed_by = ?, reviewed_role = ?,
                reviewed_at = datetime('now'), review_comment = ?
            WHERE id = ?
            """,
            (
                STATUS_APPROVED,
                normalized_actor,
                normalized_role,
                _text(comment),
                int(correction_id),
            ),
        )
        _insert_event(
            connection,
            int(correction_id),
            "approved",
            normalized_actor,
            normalized_role,
            _text(comment),
        )
        if not _apply_request(connection, request, event_type="applied"):
            raise ValueError("Catalog row for the correction was not found")


def reject_catalog_correction(
    connection: sqlite3.Connection,
    correction_id: int,
    *,
    actor: str,
    actor_role: str,
    comment: str,
) -> None:
    normalized_actor = _required_text(actor, "Reviewer is required")
    normalized_role = _validate_role(actor_role)
    normalized_comment = _required_text(comment, "Rejection comment is required")
    if normalized_role not in APPROVER_ROLES:
        raise PermissionError("Only a senior or admin can reject corrections")
    request = _request_row(connection, int(correction_id))
    if request is None or str(request["status"]) != STATUS_PENDING:
        raise ValueError("Pending correction was not found")
    with connection:
        connection.execute(
            """
            UPDATE catalog_correction_requests
            SET status = ?, reviewed_by = ?, reviewed_role = ?,
                reviewed_at = datetime('now'), review_comment = ?
            WHERE id = ?
            """,
            (
                STATUS_REJECTED,
                normalized_actor,
                normalized_role,
                normalized_comment,
                int(correction_id),
            ),
        )
        _insert_event(
            connection,
            int(correction_id),
            "rejected",
            normalized_actor,
            normalized_role,
            normalized_comment,
        )


def synchronize_catalog_corrections(connection: sqlite3.Connection) -> int:
    """Seed approved expert fixes and reapply every approved request."""
    _seed_expert_corrections(connection)
    applied = 0
    requests = connection.execute(
        """
        SELECT *
        FROM catalog_correction_requests
        WHERE status = ?
        ORDER BY id
        """,
        (STATUS_APPROVED,),
    ).fetchall()
    for request in requests:
        if _apply_request(connection, request, event_type="reapplied"):
            applied += 1
    return applied


def list_catalog_corrections(
    connection: sqlite3.Connection,
    *,
    status: str = "",
    limit: int = 500,
) -> list[CatalogCorrectionRecord]:
    params: list[object] = []
    where = ""
    normalized_status = _text(status)
    if normalized_status:
        where = "WHERE status = ?"
        params.append(normalized_status)
    params.append(max(1, min(2000, int(limit))))
    rows = connection.execute(
        f"""
        SELECT *
        FROM catalog_correction_requests
        {where}
        ORDER BY CASE status WHEN 'pending' THEN 0 ELSE 1 END, id DESC
        LIMIT ?
        """,
        params,
    ).fetchall()
    return [_correction_record(connection, row) for row in rows]


def _seed_expert_corrections(connection: sqlite3.Connection) -> None:
    try:
        payload = json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return
    for item in payload.get("corrections", []):
        seed_key = _text(item.get("seed_key"))
        if seed_key == "":
            continue
        existing = connection.execute(
            "SELECT id FROM catalog_correction_requests WHERE seed_key = ?",
            (seed_key,),
        ).fetchone()
        if existing is not None:
            continue
        target = _seed_target(
            connection,
            _text(item.get("source_filename")),
            int(item.get("source_row_number") or 0),
        )
        if target is None:
            continue
        cleaned = _clean_catalog_editor_values(dict(item.get("values") or {}))
        changes = [
            (field_name, target[field_name], new_value)
            for field_name, new_value in cleaned.items()
        ]
        reason = _required_text(item.get("reason"), "Seed reason is required")
        correction_id = _insert_request(
            connection,
            target=target,
            action=ACTION_UPDATE,
            status=STATUS_APPROVED,
            reason=reason,
            submitted_by="expert.review.6444312",
            submitted_role=ROLE_SPECIALIST,
            seed_key=seed_key,
            reviewed_by="system.senior",
            reviewed_role=ROLE_SENIOR,
            review_comment="Imported approved expert correction",
        )
        _insert_changes(connection, correction_id, changes)
        _insert_event(
            connection,
            correction_id,
            "submitted",
            "expert.review.6444312",
            ROLE_SPECIALIST,
            reason,
        )
        _insert_event(
            connection,
            correction_id,
            "approved",
            "system.senior",
            ROLE_SENIOR,
            "Imported approved expert correction",
        )
        request = _request_row(connection, correction_id)
        if request is not None:
            _apply_request(connection, request, event_type="applied", force_event=True)


def _insert_request(
    connection: sqlite3.Connection,
    *,
    target: sqlite3.Row,
    action: str,
    status: str,
    reason: str,
    submitted_by: str,
    submitted_role: str,
    seed_key: str = "",
    reviewed_by: str = "",
    reviewed_role: str = "",
    review_comment: str = "",
) -> int:
    cursor = connection.execute(
        """
        INSERT INTO catalog_correction_requests (
            seed_key, action, status, target_item_id, target_source_name,
            target_source_filename, target_source_row_number, target_task_id,
            target_code, target_unit, target_work_name, reason, submitted_by,
            submitted_role, reviewed_by, reviewed_role, reviewed_at,
            review_comment
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            seed_key or None,
            action,
            status,
            int(target["id"]),
            str(target["source_name"]),
            str(target["source_filename"]),
            int(target["source_row_number"]),
            str(target["task_id"]),
            str(target["code"]),
            str(target["unit"]),
            str(target["work_name"]),
            reason,
            submitted_by,
            submitted_role,
            reviewed_by,
            reviewed_role,
            "now" if reviewed_by else None,
            review_comment,
        ),
    )
    if reviewed_by:
        connection.execute(
            """
            UPDATE catalog_correction_requests
            SET reviewed_at = datetime('now')
            WHERE id = ?
            """,
            (int(cursor.lastrowid),),
        )
    return int(cursor.lastrowid)


def _insert_changes(
    connection: sqlite3.Connection,
    correction_id: int,
    changes: list[tuple[str, object, object]],
) -> None:
    connection.executemany(
        """
        INSERT INTO catalog_correction_changes (
            correction_id, field_name, value_type, old_value, new_value
        ) VALUES (?, ?, ?, ?, ?)
        """,
        [
            (
                int(correction_id),
                field_name,
                _value_type(new_value),
                _serialize_value(old_value),
                _serialize_value(new_value),
            )
            for field_name, old_value, new_value in changes
        ],
    )


def _insert_event(
    connection: sqlite3.Connection,
    correction_id: int,
    event_type: str,
    actor: str,
    actor_role: str,
    details: str,
) -> None:
    connection.execute(
        """
        INSERT INTO catalog_correction_events (
            correction_id, event_type, actor, actor_role, details
        ) VALUES (?, ?, ?, ?, ?)
        """,
        (
            int(correction_id),
            _text(event_type),
            _text(actor),
            _text(actor_role),
            _text(details),
        ),
    )


def _apply_request(
    connection: sqlite3.Connection,
    request: sqlite3.Row,
    *,
    event_type: str,
    force_event: bool = False,
) -> bool:
    target = _resolve_request_target(connection, request)
    if target is None:
        return False
    correction_id = int(request["id"])
    action = str(request["action"])
    changed = False
    if action == ACTION_DELETE:
        cursor = connection.execute(
            "DELETE FROM catalog_items WHERE id = ?",
            (int(target["id"]),),
        )
        changed = cursor.rowcount > 0
    else:
        changes = _request_changes(connection, correction_id)
        updates = {
            change.field_name: change.new_value
            for change in changes
            if not _values_equal(target[change.field_name], change.new_value)
        }
        if updates:
            assignments = ", ".join(f"{field_name} = ?" for field_name in updates)
            connection.execute(
                f"UPDATE catalog_items SET {assignments} WHERE id = ?",
                (*updates.values(), int(target["id"])),
            )
            changed = True
    if changed or force_event:
        connection.execute(
            """
            UPDATE catalog_correction_requests
            SET target_item_id = ?, applied_at = datetime('now')
            WHERE id = ?
            """,
            (int(target["id"]), correction_id),
        )
        _insert_event(
            connection,
            correction_id,
            event_type,
            "system",
            ROLE_ADMIN,
            f"catalog_item_id={int(target['id'])}",
        )
    return True


def _resolve_request_target(
    connection: sqlite3.Connection,
    request: sqlite3.Row,
) -> sqlite3.Row | None:
    item_id = request["target_item_id"]
    if item_id is not None:
        row = _catalog_target(connection, int(item_id))
        if row is not None and _same_stable_target(row, request):
            return row
    rows = connection.execute(
        """
        SELECT catalog_items.*, catalog_sources.name AS source_name
        FROM catalog_items
        INNER JOIN catalog_sources ON catalog_sources.id = catalog_items.source_id
        WHERE catalog_items.source_filename = ?
          AND catalog_items.source_row_number = ?
        ORDER BY catalog_items.id
        """,
        (
            str(request["target_source_filename"]),
            int(request["target_source_row_number"]),
        ),
    ).fetchall()
    if len(rows) == 1:
        return rows[0]
    same_source = [
        row for row in rows if str(row["source_name"]) == str(request["target_source_name"])
    ]
    return same_source[0] if len(same_source) == 1 else None


def _catalog_target(
    connection: sqlite3.Connection,
    item_id: int,
) -> sqlite3.Row | None:
    return connection.execute(
        """
        SELECT catalog_items.*, catalog_sources.name AS source_name
        FROM catalog_items
        INNER JOIN catalog_sources ON catalog_sources.id = catalog_items.source_id
        WHERE catalog_items.id = ?
        """,
        (int(item_id),),
    ).fetchone()


def _seed_target(
    connection: sqlite3.Connection,
    filename: str,
    row_number: int,
) -> sqlite3.Row | None:
    rows = connection.execute(
        """
        SELECT catalog_items.*, catalog_sources.name AS source_name
        FROM catalog_items
        INNER JOIN catalog_sources ON catalog_sources.id = catalog_items.source_id
        WHERE catalog_items.source_filename = ?
          AND catalog_items.source_row_number = ?
        ORDER BY catalog_items.id
        """,
        (filename, int(row_number)),
    ).fetchall()
    return rows[0] if len(rows) == 1 else None


def _same_stable_target(target: sqlite3.Row, request: sqlite3.Row) -> bool:
    return (
        str(target["source_filename"]) == str(request["target_source_filename"])
        and int(target["source_row_number"]) == int(request["target_source_row_number"])
    )


def _request_row(
    connection: sqlite3.Connection,
    correction_id: int,
) -> sqlite3.Row | None:
    return connection.execute(
        "SELECT * FROM catalog_correction_requests WHERE id = ?",
        (int(correction_id),),
    ).fetchone()


def _request_changes(
    connection: sqlite3.Connection,
    correction_id: int,
) -> list[CatalogCorrectionChange]:
    rows = connection.execute(
        """
        SELECT field_name, value_type, old_value, new_value
        FROM catalog_correction_changes
        WHERE correction_id = ?
        ORDER BY id
        """,
        (int(correction_id),),
    ).fetchall()
    return [
        CatalogCorrectionChange(
            field_name=str(row["field_name"]),
            old_value=_deserialize_value(row["old_value"], str(row["value_type"])),
            new_value=_deserialize_value(row["new_value"], str(row["value_type"])),
            value_type=str(row["value_type"]),
        )
        for row in rows
    ]


def _request_events(
    connection: sqlite3.Connection,
    correction_id: int,
) -> list[CatalogCorrectionEvent]:
    rows = connection.execute(
        """
        SELECT event_type, actor, actor_role, details, created_at
        FROM catalog_correction_events
        WHERE correction_id = ?
        ORDER BY id
        """,
        (int(correction_id),),
    ).fetchall()
    return [
        CatalogCorrectionEvent(
            event_type=str(row["event_type"]),
            actor=str(row["actor"]),
            actor_role=str(row["actor_role"]),
            details=str(row["details"]),
            created_at=str(row["created_at"]),
        )
        for row in rows
    ]


def _correction_record(
    connection: sqlite3.Connection,
    row: sqlite3.Row,
) -> CatalogCorrectionRecord:
    correction_id = int(row["id"])
    return CatalogCorrectionRecord(
        id=correction_id,
        seed_key="" if row["seed_key"] is None else str(row["seed_key"]),
        action=str(row["action"]),
        status=str(row["status"]),
        target_item_id=None if row["target_item_id"] is None else int(row["target_item_id"]),
        target_source_name=str(row["target_source_name"]),
        target_source_filename=str(row["target_source_filename"]),
        target_source_row_number=int(row["target_source_row_number"]),
        target_task_id=str(row["target_task_id"]),
        target_code=str(row["target_code"]),
        target_unit=str(row["target_unit"]),
        target_work_name=str(row["target_work_name"]),
        reason=str(row["reason"]),
        submitted_by=str(row["submitted_by"]),
        submitted_role=str(row["submitted_role"]),
        submitted_at=str(row["submitted_at"]),
        reviewed_by=str(row["reviewed_by"]),
        reviewed_role=str(row["reviewed_role"]),
        reviewed_at="" if row["reviewed_at"] is None else str(row["reviewed_at"]),
        review_comment=str(row["review_comment"]),
        applied_at="" if row["applied_at"] is None else str(row["applied_at"]),
        changes=_request_changes(connection, correction_id),
        events=_request_events(connection, correction_id),
    )


def _pending_request_exists(
    connection: sqlite3.Connection,
    target: sqlite3.Row,
    action: str,
) -> bool:
    row = connection.execute(
        """
        SELECT id
        FROM catalog_correction_requests
        WHERE status = ?
          AND action = ?
          AND target_source_filename = ?
          AND target_source_row_number = ?
        LIMIT 1
        """,
        (
            STATUS_PENDING,
            action,
            str(target["source_filename"]),
            int(target["source_row_number"]),
        ),
    ).fetchone()
    return row is not None


def _bulk_result_value(
    field: str,
    operation: str,
    raw_value: object,
    current_value: object,
) -> object:
    field_name = _text(field)
    op = _text(operation)
    if field_name in CATALOG_EDITOR_TEXT_FIELDS:
        if op != "set":
            raise ValueError("Text fields support only set operation")
        cleaned = _clean_catalog_editor_values({field_name: raw_value})
        return cleaned[field_name]
    if field_name not in CATALOG_EDITOR_NUMERIC_FIELDS:
        raise ValueError("Unsupported catalog field")
    number = _parse_catalog_editor_number(field_name, raw_value)
    if op == "set":
        result = number
    elif op == "add":
        result = (0.0 if current_value is None else float(current_value)) + number
    elif op == "multiply":
        result = None if current_value is None else float(current_value) * number
    else:
        raise ValueError("Unsupported bulk operation")
    if field_name in REQUIRED_CATALOG_NUMERIC_FIELDS and (
        result is None or float(result) <= 0
    ):
        raise ValueError("Required numeric fields must stay positive")
    return result


def _validate_action(value: object) -> str:
    action = _text(value)
    if action not in SUPPORTED_ACTIONS:
        raise ValueError("Unsupported correction action")
    return action


def _validate_role(value: object) -> str:
    role = _text(value)
    if role not in SUPPORTED_ROLES:
        raise ValueError("Unsupported user role")
    return role


def _required_text(value: object, message: str) -> str:
    text = _text(value)
    if text == "":
        raise ValueError(message)
    return text


def _text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _value_type(value: object) -> str:
    if value is None:
        return VALUE_NULL
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return VALUE_NUMBER
    return VALUE_TEXT


def _serialize_value(value: object) -> str | None:
    if value is None:
        return None
    if isinstance(value, float):
        return format(value, ".17g")
    return str(value)


def _deserialize_value(value: object, value_type: str) -> object:
    if value_type == VALUE_NULL or value is None:
        return None
    if value_type == VALUE_NUMBER:
        return float(value)
    return str(value)


def _values_equal(first: object, second: object) -> bool:
    if first is None or second is None:
        return first is None and second is None
    if isinstance(first, (int, float)) and isinstance(second, (int, float)):
        return abs(float(first) - float(second)) <= 1e-12
    return str(first) == str(second)
