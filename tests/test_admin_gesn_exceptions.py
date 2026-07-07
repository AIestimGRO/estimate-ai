"""Tests for the read-only admin GESN exceptions page."""

from fastapi.testclient import TestClient

from app.web.app import create_app
from core.storage import connect, init_database


def test_admin_gesn_exceptions_shows_empty_state(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/gesn-exceptions")

    assert response.status_code == 200
    assert "GESN exceptions" in response.text
    assert "Одобренные диапазоны пока не записаны" in response.text
    assert 'class="admin-nav-link active" href="/admin/gesn-exceptions"' in response.text


def test_admin_gesn_exceptions_lists_approved_ranges(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    connection = connect(db_path)
    try:
        init_database(connection)
        connection.execute(
            """
            INSERT INTO gesn_exceptions (
                exception_key, approved_min, approved_max, last_range_update_date
            ) VALUES (?, ?, ?, ?)
            """,
            (
                "m||GESN01-01-001-01||0",
                100.0,
                350.0,
                45200.0,
            ),
        )
        connection.commit()
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/gesn-exceptions")

    assert response.status_code == 200
    assert "m||GESN01-01-001-01||0" in response.text
    assert "GESN01-01-001-01" in response.text
    assert "<td>m</td>" in response.text
    assert "<td>нет</td>" in response.text
    assert "<td>100</td>" in response.text
    assert "<td>350</td>" in response.text
    assert "<td>45200</td>" in response.text
    assert 'class="admin-nav-link active" href="/admin/gesn-exceptions"' in response.text
