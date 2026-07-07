"""Tests for approving price risks from the admin UI."""

from fastapi.testclient import TestClient

from app.web.app import create_app
from core.storage import (
    connect,
    get_price_risk,
    init_database,
    list_gesn_exceptions,
    list_price_risks,
)
from core.storage.risk_log import STATUS_APPROVED, STATUS_OPEN


def _insert_open_risk(connection, *, key: str = "m||GESN01-01-001-01||NO_DEM") -> None:
    connection.execute(
        """
        INSERT INTO price_risk_log (
            exception_key, status, reason, code, unit,
            min_price, max_price, ratio, recommended_price, estimate_row
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            key,
            STATUS_OPEN,
            "RATIO_EXCEEDED",
            "GESN01-01-001-01",
            "m",
            100.0,
            350.0,
            3.5,
            200.0,
            12,
        ),
    )
    connection.commit()


def test_admin_approvals_page_has_approve_button(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    connection = connect(db_path)
    try:
        init_database(connection)
        _insert_open_risk(connection)
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/approvals")

    assert response.status_code == 200
    assert 'action="/admin/approvals/approve"' in response.text
    assert 'name="exception_key"' in response.text
    assert "Одобрить" in response.text


def test_admin_approves_open_risk_and_writes_exception(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    key = "m||GESN01-01-001-01||NO_DEM"

    connection = connect(db_path)
    try:
        init_database(connection)
        _insert_open_risk(connection, key=key)
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/approvals/approve",
            data={"exception_key": key},
            follow_redirects=False,
        )

    assert response.status_code == 303
    assert response.headers["location"] == "/admin/approvals"

    connection = connect(db_path)
    try:
        exceptions = list_gesn_exceptions(connection)
        approved = get_price_risk(
            connection,
            exception_key=key,
            status=STATUS_APPROVED,
        )
        open_rows = list_price_risks(connection, status=STATUS_OPEN)
    finally:
        connection.close()

    assert len(exceptions) == 1
    assert exceptions[0].exception_key == key
    assert exceptions[0].approved_min == 100.0
    assert exceptions[0].approved_max == 350.0
    assert exceptions[0].last_range_update_date > 0
    assert approved is not None
    assert open_rows == []


def test_admin_approval_widens_existing_exception(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    key = "m||GESN01-01-001-01||NO_DEM"

    connection = connect(db_path)
    try:
        init_database(connection)
        connection.execute(
            """
            INSERT INTO gesn_exceptions (
                exception_key, approved_min, approved_max, last_range_update_date
            ) VALUES (?, ?, ?, ?)
            """,
            (key, 150.0, 250.0, 10.0),
        )
        _insert_open_risk(connection, key=key)
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/approvals/approve",
            data={"exception_key": key},
            follow_redirects=False,
        )

    assert response.status_code == 303

    connection = connect(db_path)
    try:
        exceptions = list_gesn_exceptions(connection)
    finally:
        connection.close()

    assert len(exceptions) == 1
    assert exceptions[0].approved_min == 100.0
    assert exceptions[0].approved_max == 350.0
    assert exceptions[0].last_range_update_date > 10.0


def test_admin_approval_rejects_unknown_open_risk(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/approvals/approve",
            data={"exception_key": "missing||key||NO_DEM"},
        )

    assert response.status_code == 404
    assert "Открытый риск для одобрения не найден" in response.text
