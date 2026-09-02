"""Persistent processing workspaces, edits, requests, and audit history."""

from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass


STATUS_DRAFT = "draft"
STATUS_READY = "ready"
STATUS_COMPLETE = "complete"
REQUEST_PENDING = "pending"
REQUEST_APPROVED = "approved"
REQUEST_REJECTED = "rejected"
REQUEST_BLUE_TASK = "blue_task"
REQUEST_ANALOG_CHANGE = "analog_change"
REQUEST_RULE_CHANGE = "rule_change"
REQUEST_OTHER = "other"
SUPPORTED_REQUEST_TYPES = frozenset(
    {
        REQUEST_BLUE_TASK,
        REQUEST_ANALOG_CHANGE,
        REQUEST_RULE_CHANGE,
        REQUEST_OTHER,
    }
)


@dataclass(frozen=True)
class ProcessingJob:
    id: str
    owner_user_id: int
    owner_name: str
    estimate_filename: str
    source_path: str
    output_path: str
    final_output_path: str
    sheet_title: str
    header_row: int
    coefficient: float
    region: str
    use_tkp_analogs: bool
    status: str
    total_rows: int
    matched_rows: int
    flagged_rows: int
    tkp_matched_rows: int
    column_schema: list[dict[str, object]]
    created_at: str
    updated_at: str


@dataclass(frozen=True)
class ProcessingRow:
    id: int
    job_id: str
    row_index: int
    excel_row_number: int
    values: list[object]


@dataclass(frozen=True)
class CellOverride:
    id: int
    job_id: str
    row_id: int
    column_index: int
    original_value: object
    current_value: object
    editor_user_id: int
    editor_name: str
    revision: int
    updated_at: str


@dataclass(frozen=True)
class ActivityEvent:
    id: int
    actor_user_id: int | None
    actor_name: str
    event_type: str
    entity_type: str
    entity_id: str
    job_id: str
    estimate_filename: str
    details: str
    created_at: str


@dataclass(frozen=True)
class ChangeRequest:
    id: int
    job_id: str
    estimate_filename: str
    row_id: int | None
    excel_row_number: int | None
    column_index: int | None
    request_type: str
    task_number: str
    comment: str
    status: str
    submitted_by: int
    submitted_by_name: str
    submitted_at: str
    reviewed_by: int | None
    reviewed_by_name: str
    reviewed_at: str
    review_comment: str


def upsert_processing_job(
    connection: sqlite3.Connection,
    *,
    job_id: str,
    owner_user_id: int,
    estimate_filename: str,
    source_path: str,
    output_path: str,
    sheet_title: str,
    header_row: int,
    coefficient: float,
    region: str,
    use_tkp_analogs: bool,
    status: str,
    total_rows: int,
    matched_rows: int,
    flagged_rows: int,
    tkp_matched_rows: int,
    column_schema: list[dict[str, object]],
) -> None:
    payload = json.dumps(column_schema, ensure_ascii=False, separators=(",", ":"))
    existing = connection.execute(
        "SELECT id FROM processing_jobs WHERE id = ?",
        (str(job_id),),
    ).fetchone()
    with connection:
        connection.execute(
            """
            INSERT INTO processing_jobs(
                id, owner_user_id, estimate_filename, source_path, output_path,
                sheet_title, header_row, coefficient, region, use_tkp_analogs,
                status, total_rows, matched_rows, flagged_rows,
                tkp_matched_rows, column_schema_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
                owner_user_id=excluded.owner_user_id,
                estimate_filename=excluded.estimate_filename,
                source_path=excluded.source_path,
                output_path=excluded.output_path,
                sheet_title=excluded.sheet_title,
                header_row=excluded.header_row,
                coefficient=excluded.coefficient,
                region=excluded.region,
                use_tkp_analogs=excluded.use_tkp_analogs,
                status=excluded.status,
                total_rows=excluded.total_rows,
                matched_rows=excluded.matched_rows,
                flagged_rows=excluded.flagged_rows,
                tkp_matched_rows=excluded.tkp_matched_rows,
                column_schema_json=excluded.column_schema_json,
                updated_at=datetime('now')
            """,
            (
                str(job_id),
                int(owner_user_id),
                str(estimate_filename),
                str(source_path),
                str(output_path),
                str(sheet_title),
                int(header_row),
                float(coefficient),
                str(region or ""),
                1 if use_tkp_analogs else 0,
                str(status),
                int(total_rows),
                int(matched_rows),
                int(flagged_rows),
                int(tkp_matched_rows),
                payload,
            ),
        )
        if existing is None:
            _insert_audit(
                connection,
                actor_user_id=int(owner_user_id),
                event_type="processing_created",
                entity_type="processing_job",
                entity_id=str(job_id),
                job_id=str(job_id),
                details=str(estimate_filename),
            )


def replace_processing_rows(
    connection: sqlite3.Connection,
    job_id: str,
    rows: list[tuple[int, int, list[object]]],
) -> None:
    with connection:
        connection.execute(
            "DELETE FROM processing_rows WHERE job_id = ?",
            (str(job_id),),
        )
        connection.executemany(
            """
            INSERT INTO processing_rows(
                job_id, row_index, excel_row_number, values_json
            ) VALUES (?, ?, ?, ?)
            """,
            [
                (
                    str(job_id),
                    int(row_index),
                    int(excel_row_number),
                    json.dumps(values, ensure_ascii=False, separators=(",", ":")),
                )
                for row_index, excel_row_number, values in rows
            ],
        )


def list_processing_jobs(
    connection: sqlite3.Connection,
    *,
    owner_user_id: int | None = None,
    limit: int = 500,
) -> list[ProcessingJob]:
    params: list[object] = []
    where = ""
    if owner_user_id is not None:
        where = "WHERE j.owner_user_id = ?"
        params.append(int(owner_user_id))
    params.append(max(1, min(2000, int(limit))))
    rows = connection.execute(
        f"""
        SELECT j.*, u.full_name AS owner_name
        FROM processing_jobs j
        JOIN app_users u ON u.id = j.owner_user_id
        {where}
        ORDER BY j.updated_at DESC, j.created_at DESC
        LIMIT ?
        """,
        params,
    ).fetchall()
    return [_job_record(row) for row in rows]


def get_processing_job(
    connection: sqlite3.Connection,
    job_id: str,
) -> ProcessingJob | None:
    row = connection.execute(
        """
        SELECT j.*, u.full_name AS owner_name
        FROM processing_jobs j
        JOIN app_users u ON u.id = j.owner_user_id
        WHERE j.id = ?
        """,
        (str(job_id),),
    ).fetchone()
    return None if row is None else _job_record(row)


def list_processing_rows(
    connection: sqlite3.Connection,
    job_id: str,
) -> list[ProcessingRow]:
    rows = connection.execute(
        """
        SELECT id, job_id, row_index, excel_row_number, values_json
        FROM processing_rows
        WHERE job_id = ?
        ORDER BY row_index
        """,
        (str(job_id),),
    ).fetchall()
    return [_processing_row(row) for row in rows]


def get_processing_row(
    connection: sqlite3.Connection,
    row_id: int,
) -> ProcessingRow | None:
    row = connection.execute(
        """
        SELECT id, job_id, row_index, excel_row_number, values_json
        FROM processing_rows
        WHERE id = ?
        """,
        (int(row_id),),
    ).fetchone()
    return None if row is None else _processing_row(row)


def list_cell_overrides(
    connection: sqlite3.Connection,
    job_id: str,
) -> list[CellOverride]:
    rows = connection.execute(
        """
        SELECT o.*, u.full_name AS editor_name
        FROM processing_cell_overrides o
        JOIN app_users u ON u.id = o.editor_user_id
        WHERE o.job_id = ?
        ORDER BY o.row_id, o.column_index
        """,
        (str(job_id),),
    ).fetchall()
    return [_override_record(row) for row in rows]


def save_cell_override(
    connection: sqlite3.Connection,
    *,
    job_id: str,
    row_id: int,
    column_index: int,
    new_value: object,
    editor_user_id: int,
    client_change_id: str,
) -> CellOverride:
    change_id = str(client_change_id or "").strip()
    if not change_id:
        raise ValueError("Client change id is required")
    column = int(column_index)
    if column <= 0 or column > 16_384:
        raise ValueError("Invalid column index")
    row = get_processing_row(connection, int(row_id))
    if row is None or row.job_id != str(job_id):
        raise ValueError("Processing row was not found")
    if column > len(row.values):
        raise ValueError("Column is outside the processing grid")

    existing_event = connection.execute(
        """
        SELECT override_id
        FROM processing_edit_events
        WHERE client_change_id = ?
        """,
        (change_id,),
    ).fetchone()
    if existing_event is not None:
        existing = _get_override(connection, int(existing_event["override_id"]))
        if existing is None:
            raise ValueError("Stored edit event is inconsistent")
        return existing

    current = connection.execute(
        """
        SELECT *
        FROM processing_cell_overrides
        WHERE job_id = ? AND row_id = ? AND column_index = ?
        """,
        (str(job_id), int(row_id), column),
    ).fetchone()
    original_value = row.values[column - 1]
    old_value = original_value if current is None else _json_load(current["current_value_json"])
    next_revision = 1 if current is None else int(current["revision"]) + 1

    with connection:
        if current is None:
            cursor = connection.execute(
                """
                INSERT INTO processing_cell_overrides(
                    job_id, row_id, column_index, original_value_json,
                    current_value_json, editor_user_id, revision
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    str(job_id),
                    int(row_id),
                    column,
                    _json_dump(original_value),
                    _json_dump(new_value),
                    int(editor_user_id),
                    next_revision,
                ),
            )
            override_id = int(cursor.lastrowid)
        else:
            override_id = int(current["id"])
            connection.execute(
                """
                UPDATE processing_cell_overrides
                SET current_value_json = ?, editor_user_id = ?,
                    revision = ?, updated_at = datetime('now')
                WHERE id = ?
                """,
                (
                    _json_dump(new_value),
                    int(editor_user_id),
                    next_revision,
                    override_id,
                ),
            )
        connection.execute(
            """
            INSERT INTO processing_edit_events(
                override_id, job_id, row_id, column_index,
                old_value_json, new_value_json, actor_user_id,
                client_change_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                override_id,
                str(job_id),
                int(row_id),
                column,
                _json_dump(old_value),
                _json_dump(new_value),
                int(editor_user_id),
                change_id,
            ),
        )
        _insert_audit(
            connection,
            actor_user_id=int(editor_user_id),
            event_type="cell_edited",
            entity_type="processing_cell",
            entity_id=f"{int(row_id)}:{column}",
            job_id=str(job_id),
            details=json.dumps(
                {
                    "excel_row": int(row.excel_row_number),
                    "column": column,
                    "old": old_value,
                    "new": new_value,
                    "revision": next_revision,
                },
                ensure_ascii=False,
                separators=(",", ":"),
            ),
        )
        connection.execute(
            """
            UPDATE processing_jobs
            SET status = ?, updated_at = datetime('now')
            WHERE id = ?
            """,
            (STATUS_DRAFT, str(job_id)),
        )

    result = _get_override(connection, override_id)
    if result is None:
        raise ValueError("Saved override was not found")
    return result


def create_change_request(
    connection: sqlite3.Connection,
    *,
    job_id: str,
    row_id: int | None,
    column_index: int | None,
    request_type: str,
    task_number: str,
    comment: str,
    submitted_by: int,
) -> int:
    normalized_type = str(request_type or "").strip()
    if normalized_type not in SUPPORTED_REQUEST_TYPES:
        raise ValueError("Unsupported request type")
    normalized_comment = str(comment or "").strip()
    if not normalized_comment:
        raise ValueError("Comment is required")
    normalized_task = str(task_number or "").strip()

    excel_row_number: int | None = None
    normalized_row_id: int | None = None
    if row_id is not None:
        row = get_processing_row(connection, int(row_id))
        if row is None or row.job_id != str(job_id):
            raise ValueError("Processing row was not found")
        normalized_row_id = int(row.id)
        excel_row_number = int(row.excel_row_number)

    with connection:
        cursor = connection.execute(
            """
            INSERT INTO specialist_change_requests(
                job_id, row_id, excel_row_number, column_index,
                request_type, task_number, comment, submitted_by
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                str(job_id),
                normalized_row_id,
                excel_row_number,
                int(column_index) if column_index is not None else None,
                normalized_type,
                normalized_task,
                normalized_comment,
                int(submitted_by),
            ),
        )
        request_id = int(cursor.lastrowid)
        _insert_audit(
            connection,
            actor_user_id=int(submitted_by),
            event_type="change_request_submitted",
            entity_type="change_request",
            entity_id=str(request_id),
            job_id=str(job_id),
            details=normalized_type,
        )
    return request_id


def list_change_requests(
    connection: sqlite3.Connection,
    *,
    status: str = "",
    submitted_by: int | None = None,
    limit: int = 1000,
) -> list[ChangeRequest]:
    clauses: list[str] = []
    params: list[object] = []
    if str(status or "").strip():
        clauses.append("r.status = ?")
        params.append(str(status).strip())
    if submitted_by is not None:
        clauses.append("r.submitted_by = ?")
        params.append(int(submitted_by))
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    params.append(max(1, min(5000, int(limit))))
    rows = connection.execute(
        f"""
        SELECT r.*, j.estimate_filename,
               su.full_name AS submitted_by_name,
               COALESCE(ru.full_name, '') AS reviewed_by_name
        FROM specialist_change_requests r
        JOIN processing_jobs j ON j.id = r.job_id
        JOIN app_users su ON su.id = r.submitted_by
        LEFT JOIN app_users ru ON ru.id = r.reviewed_by
        {where}
        ORDER BY CASE r.status WHEN 'pending' THEN 0 ELSE 1 END,
                 r.submitted_at DESC, r.id DESC
        LIMIT ?
        """,
        params,
    ).fetchall()
    return [_request_record(row) for row in rows]


def get_change_request(
    connection: sqlite3.Connection,
    request_id: int,
) -> ChangeRequest | None:
    rows = list_change_requests(connection, limit=5000)
    for item in rows:
        if item.id == int(request_id):
            return item
    return None


def review_change_request(
    connection: sqlite3.Connection,
    request_id: int,
    *,
    status: str,
    reviewed_by: int,
    review_comment: str = "",
) -> bool:
    normalized_status = str(status or "").strip()
    if normalized_status not in {REQUEST_APPROVED, REQUEST_REJECTED}:
        raise ValueError("Unsupported review status")
    request_row = connection.execute(
        "SELECT job_id FROM specialist_change_requests WHERE id = ?",
        (int(request_id),),
    ).fetchone()
    request_job_id = "" if request_row is None else str(request_row["job_id"])
    with connection:
        cursor = connection.execute(
            """
            UPDATE specialist_change_requests
            SET status = ?, reviewed_by = ?, reviewed_at = datetime('now'),
                review_comment = ?
            WHERE id = ? AND status = ?
            """,
            (
                normalized_status,
                int(reviewed_by),
                str(review_comment or "").strip(),
                int(request_id),
                REQUEST_PENDING,
            ),
        )
        if cursor.rowcount:
            _insert_audit(
                connection,
                actor_user_id=int(reviewed_by),
                event_type=f"change_request_{normalized_status}",
                entity_type="change_request",
                entity_id=str(int(request_id)),
                job_id=request_job_id,
                details=str(review_comment or "").strip(),
            )
    return cursor.rowcount > 0


def set_job_final_output(
    connection: sqlite3.Connection,
    job_id: str,
    final_output_path: str,
) -> None:
    with connection:
        connection.execute(
            """
            UPDATE processing_jobs
            SET final_output_path = ?, updated_at = datetime('now')
            WHERE id = ?
            """,
            (str(final_output_path), str(job_id)),
        )


def touch_job(connection: sqlite3.Connection, job_id: str) -> None:
    with connection:
        connection.execute(
            """
            UPDATE processing_jobs
            SET last_opened_at = datetime('now')
            WHERE id = ?
            """,
            (str(job_id),),
        )


def save_user_preference(
    connection: sqlite3.Connection,
    user_id: int,
    *,
    key: str,
    value: object,
) -> None:
    normalized_key = str(key or "").strip()
    if not normalized_key:
        raise ValueError("Preference key is required")
    with connection:
        connection.execute(
            """
            INSERT INTO user_preferences(user_id, preference_key, value_json)
            VALUES (?, ?, ?)
            ON CONFLICT(user_id, preference_key) DO UPDATE SET
                value_json=excluded.value_json,
                updated_at=datetime('now')
            """,
            (int(user_id), normalized_key, _json_dump(value)),
        )


def list_activity_events(
    connection: sqlite3.Connection,
    *,
    limit: int = 1000,
) -> list[ActivityEvent]:
    rows = connection.execute(
        """
        SELECT a.id, a.actor_user_id, COALESCE(u.full_name, '') AS actor_name,
               a.event_type, a.entity_type, a.entity_id, a.job_id,
               COALESCE(j.estimate_filename, '') AS estimate_filename,
               a.details, a.created_at
        FROM audit_events a
        LEFT JOIN app_users u ON u.id = a.actor_user_id
        LEFT JOIN processing_jobs j ON j.id = a.job_id
        ORDER BY a.id DESC
        LIMIT ?
        """,
        (max(1, min(10_000, int(limit))),),
    ).fetchall()
    return [
        ActivityEvent(
            id=int(row["id"]),
            actor_user_id=(
                int(row["actor_user_id"])
                if row["actor_user_id"] is not None
                else None
            ),
            actor_name=str(row["actor_name"]),
            event_type=str(row["event_type"]),
            entity_type=str(row["entity_type"]),
            entity_id=str(row["entity_id"]),
            job_id=str(row["job_id"]),
            estimate_filename=str(row["estimate_filename"]),
            details=str(row["details"]),
            created_at=str(row["created_at"]),
        )
        for row in rows
    ]


def get_user_preference(
    connection: sqlite3.Connection,
    user_id: int,
    *,
    key: str,
    default: object = None,
) -> object:
    row = connection.execute(
        """
        SELECT value_json
        FROM user_preferences
        WHERE user_id = ? AND preference_key = ?
        """,
        (int(user_id), str(key or "").strip()),
    ).fetchone()
    if row is None:
        return default
    return _json_load(row["value_json"])


def record_activity_event(
    connection: sqlite3.Connection,
    *,
    actor_user_id: int,
    event_type: str,
    entity_type: str,
    entity_id: str = "",
    job_id: str = "",
    details: str = "",
) -> None:
    with connection:
        _insert_audit(
            connection,
            actor_user_id=int(actor_user_id),
            event_type=event_type,
            entity_type=entity_type,
            entity_id=entity_id,
            job_id=job_id,
            details=details,
        )


def _insert_audit(
    connection: sqlite3.Connection,
    *,
    actor_user_id: int,
    event_type: str,
    entity_type: str,
    entity_id: str,
    job_id: str,
    details: str,
) -> None:
    connection.execute(
        """
        INSERT INTO audit_events(
            actor_user_id, event_type, entity_type, entity_id, job_id, details
        ) VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            int(actor_user_id),
            str(event_type),
            str(entity_type),
            str(entity_id),
            str(job_id),
            str(details),
        ),
    )


def _get_override(
    connection: sqlite3.Connection,
    override_id: int,
) -> CellOverride | None:
    row = connection.execute(
        """
        SELECT o.*, u.full_name AS editor_name
        FROM processing_cell_overrides o
        JOIN app_users u ON u.id = o.editor_user_id
        WHERE o.id = ?
        """,
        (int(override_id),),
    ).fetchone()
    return None if row is None else _override_record(row)


def _job_record(row: sqlite3.Row) -> ProcessingJob:
    return ProcessingJob(
        id=str(row["id"]),
        owner_user_id=int(row["owner_user_id"]),
        owner_name=str(row["owner_name"]),
        estimate_filename=str(row["estimate_filename"]),
        source_path=str(row["source_path"]),
        output_path=str(row["output_path"]),
        final_output_path=str(row["final_output_path"]),
        sheet_title=str(row["sheet_title"]),
        header_row=int(row["header_row"]),
        coefficient=float(row["coefficient"]),
        region=str(row["region"]),
        use_tkp_analogs=bool(row["use_tkp_analogs"]),
        status=str(row["status"]),
        total_rows=int(row["total_rows"]),
        matched_rows=int(row["matched_rows"]),
        flagged_rows=int(row["flagged_rows"]),
        tkp_matched_rows=int(row["tkp_matched_rows"]),
        column_schema=list(json.loads(str(row["column_schema_json"]) or "[]")),
        created_at=str(row["created_at"]),
        updated_at=str(row["updated_at"]),
    )


def _processing_row(row: sqlite3.Row) -> ProcessingRow:
    return ProcessingRow(
        id=int(row["id"]),
        job_id=str(row["job_id"]),
        row_index=int(row["row_index"]),
        excel_row_number=int(row["excel_row_number"]),
        values=list(json.loads(str(row["values_json"]) or "[]")),
    )


def _override_record(row: sqlite3.Row) -> CellOverride:
    return CellOverride(
        id=int(row["id"]),
        job_id=str(row["job_id"]),
        row_id=int(row["row_id"]),
        column_index=int(row["column_index"]),
        original_value=_json_load(row["original_value_json"]),
        current_value=_json_load(row["current_value_json"]),
        editor_user_id=int(row["editor_user_id"]),
        editor_name=str(row["editor_name"]),
        revision=int(row["revision"]),
        updated_at=str(row["updated_at"]),
    )


def _request_record(row: sqlite3.Row) -> ChangeRequest:
    return ChangeRequest(
        id=int(row["id"]),
        job_id=str(row["job_id"]),
        estimate_filename=str(row["estimate_filename"]),
        row_id=int(row["row_id"]) if row["row_id"] is not None else None,
        excel_row_number=(
            int(row["excel_row_number"])
            if row["excel_row_number"] is not None
            else None
        ),
        column_index=(
            int(row["column_index"])
            if row["column_index"] is not None
            else None
        ),
        request_type=str(row["request_type"]),
        task_number=str(row["task_number"]),
        comment=str(row["comment"]),
        status=str(row["status"]),
        submitted_by=int(row["submitted_by"]),
        submitted_by_name=str(row["submitted_by_name"]),
        submitted_at=str(row["submitted_at"]),
        reviewed_by=(
            int(row["reviewed_by"]) if row["reviewed_by"] is not None else None
        ),
        reviewed_by_name=str(row["reviewed_by_name"]),
        reviewed_at=str(row["reviewed_at"] or ""),
        review_comment=str(row["review_comment"]),
    )


def _json_dump(value: object) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def _json_load(value: object) -> object:
    if value is None:
        return None
    return json.loads(str(value))
