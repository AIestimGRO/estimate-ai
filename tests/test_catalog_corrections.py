"""Approval-backed catalog corrections and immutable audit events."""

import pytest

from core.storage import connect, init_database
from core.storage.corrections import (
    ACTION_DELETE,
    ROLE_SENIOR,
    ROLE_SPECIALIST,
    STATUS_APPROVED,
    STATUS_PENDING,
    approve_catalog_correction,
    create_catalog_correction,
    list_catalog_corrections,
    reject_catalog_correction,
)


def _seed_catalog_row(connection, *, item_id: int | None = None) -> int:
    connection.execute(
        "INSERT OR IGNORE INTO catalog_sources(name, kind) VALUES (?, ?)",
        ("legacy", "excel_bulk"),
    )
    source = connection.execute(
        "SELECT id FROM catalog_sources WHERE name = ?",
        ("legacy",),
    ).fetchone()
    columns = """
        source_id, task_id, region, code, unit, quantity, work_name, price,
        price_original, price_zlvl, total_price, labor_unit, labor_total,
        machine_labor_unit, machine_labor_total, regional_coefficient,
        source_filename, source_row_number
    """
    values = (
        int(source["id"]),
        "TASK-1",
        "Moscow",
        "GESN01-01-001-01",
        "m",
        10.0,
        "Two layer work",
        100.0,
        100.0,
        100.0,
        1000.0,
        1.0,
        10.0,
        0.2,
        2.0,
        1.0,
        "source.xlsx",
        42,
    )
    placeholders = ", ".join("?" for _ in values)
    if item_id is None:
        cursor = connection.execute(
            f"INSERT INTO catalog_items ({columns}) VALUES ({placeholders})",
            values,
        )
    else:
        cursor = connection.execute(
            f"INSERT INTO catalog_items (id, {columns}) VALUES (?, {placeholders})",
            (item_id, *values),
        )
    connection.commit()
    return int(cursor.lastrowid)


def test_pending_correction_does_not_change_catalog_until_approved(tmp_path) -> None:
    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        item_id = _seed_catalog_row(connection)
        correction_id = create_catalog_correction(
            connection,
            item_id,
            values={"price": 80.0, "price_original": 80.0, "price_zlvl": 80.0},
            reason="Expert correction",
            actor="specialist.one",
            actor_role=ROLE_SPECIALIST,
        )

        price_before = connection.execute(
            "SELECT price FROM catalog_items WHERE id = ?",
            (item_id,),
        ).fetchone()["price"]
        pending = list_catalog_corrections(connection, status=STATUS_PENDING)
        assert price_before == 100.0
        assert len(pending) == 1
        assert pending[0].id == correction_id
        assert pending[0].submitted_by == "specialist.one"
        assert {change.field_name for change in pending[0].changes} == {
            "price",
            "price_original",
            "price_zlvl",
        }

        approve_catalog_correction(
            connection,
            correction_id,
            actor="senior.one",
            actor_role=ROLE_SENIOR,
            comment="Checked",
        )

        price_after = connection.execute(
            "SELECT price FROM catalog_items WHERE id = ?",
            (item_id,),
        ).fetchone()["price"]
        approved = list_catalog_corrections(connection, status=STATUS_APPROVED)
        assert price_after == 80.0
        assert len(approved) == 1
        assert approved[0].reviewed_by == "senior.one"
        assert [event.event_type for event in approved[0].events] == [
            "submitted",
            "approved",
            "applied",
        ]
    finally:
        connection.close()


def test_approved_correction_reapplies_to_rebuilt_catalog_row(tmp_path) -> None:
    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        item_id = _seed_catalog_row(connection)
        correction_id = create_catalog_correction(
            connection,
            item_id,
            values={"price": 80.0},
            reason="Expert correction",
            actor="specialist.one",
            actor_role=ROLE_SPECIALIST,
        )
        approve_catalog_correction(
            connection,
            correction_id,
            actor="senior.one",
            actor_role=ROLE_SENIOR,
        )

        connection.execute("DELETE FROM catalog_items WHERE id = ?", (item_id,))
        rebuilt_id = _seed_catalog_row(connection, item_id=item_id + 100)
        init_database(connection)

        rebuilt_price = connection.execute(
            "SELECT price FROM catalog_items WHERE id = ?",
            (rebuilt_id,),
        ).fetchone()["price"]
        approved = list_catalog_corrections(connection, status=STATUS_APPROVED)
        assert rebuilt_price == 80.0
        assert approved[0].events[-1].event_type == "reapplied"
    finally:
        connection.close()


def test_approved_delete_reapplies_after_rebuild(tmp_path) -> None:
    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        item_id = _seed_catalog_row(connection)
        correction_id = create_catalog_correction(
            connection,
            item_id,
            values={},
            reason="Invalid analog",
            actor="specialist.one",
            actor_role=ROLE_SPECIALIST,
            action=ACTION_DELETE,
        )
        approve_catalog_correction(
            connection,
            correction_id,
            actor="senior.one",
            actor_role=ROLE_SENIOR,
        )
        assert connection.execute(
            "SELECT COUNT(*) AS count FROM catalog_items"
        ).fetchone()["count"] == 0

        _seed_catalog_row(connection)
        init_database(connection)
        assert connection.execute(
            "SELECT COUNT(*) AS count FROM catalog_items"
        ).fetchone()["count"] == 0
        approved = list_catalog_corrections(connection, status=STATUS_APPROVED)
        assert approved[0].events[-1].event_type == "reapplied"
    finally:
        connection.close()


def test_only_senior_role_can_approve(tmp_path) -> None:
    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        item_id = _seed_catalog_row(connection)
        correction_id = create_catalog_correction(
            connection,
            item_id,
            values={"price": 80.0},
            reason="Expert correction",
            actor="specialist.one",
            actor_role=ROLE_SPECIALIST,
        )

        with pytest.raises(PermissionError):
            approve_catalog_correction(
                connection,
                correction_id,
                actor="specialist.two",
                actor_role=ROLE_SPECIALIST,
            )
    finally:
        connection.close()


def test_rejected_correction_keeps_catalog_and_logs_decision(tmp_path) -> None:
    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        item_id = _seed_catalog_row(connection)
        correction_id = create_catalog_correction(
            connection,
            item_id,
            values={"price": 80.0},
            reason="Expert correction",
            actor="specialist.one",
            actor_role=ROLE_SPECIALIST,
        )
        reject_catalog_correction(
            connection,
            correction_id,
            actor="senior.one",
            actor_role=ROLE_SENIOR,
            comment="Source does not confirm the value",
        )

        price = connection.execute(
            "SELECT price FROM catalog_items WHERE id = ?",
            (item_id,),
        ).fetchone()["price"]
        rejected = list_catalog_corrections(connection, status="rejected")
        assert price == 100.0
        assert len(rejected) == 1
        assert [event.event_type for event in rejected[0].events] == [
            "submitted",
            "rejected",
        ]
    finally:
        connection.close()


def test_real_expert_seed_is_approved_and_applied_once(tmp_path) -> None:
    connection = connect(tmp_path / "estimate_ai.db")
    try:
        init_database(connection)
        source = connection.execute(
            "INSERT INTO catalog_sources(name, kind) VALUES (?, ?)",
            ("main", "excel_bulk"),
        )
        connection.execute(
            """
            INSERT INTO catalog_items (
                source_id, task_id, code, unit, quantity, work_name, price,
                price_original, price_zlvl, total_price, regional_coefficient,
                source_filename, source_row_number
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(source.lastrowid),
                "5805212_5667536",
                "GESN06-14-002-03",
                "m",
                54.0,
                "Expansion joint",
                1525.3458526666666,
                1525.3458526666666,
                1525.3458526666666,
                82368.67604399999,
                1.0,
                "\u0420\u041d\u041c\u0426 _5805212_5667536.xlsx",
                7072,
            ),
        )
        connection.commit()

        init_database(connection)
        init_database(connection)

        row = connection.execute(
            """
            SELECT price, price_original, price_zlvl, total_price
            FROM catalog_items
            WHERE source_row_number = 7072
            """
        ).fetchone()
        corrections = list_catalog_corrections(connection, status=STATUS_APPROVED)
        event_count = connection.execute(
            "SELECT COUNT(*) AS count FROM catalog_correction_events"
        ).fetchone()["count"]
        assert dict(row) == {
            "price": 342.2,
            "price_original": 342.2,
            "price_zlvl": 342.2,
            "total_price": 18478.8,
        }
        assert len(corrections) == 1
        assert corrections[0].seed_key == "review-6444312-gesn06-14-002-03-5805212"
        assert event_count == 3
    finally:
        connection.close()
