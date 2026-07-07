"""Tests for the read-only admin price risk log page."""

from fastapi.testclient import TestClient

from app.web.app import create_app
from core.storage import connect, init_database


def test_admin_risks_shows_empty_state(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/risks")

    assert response.status_code == 200
    assert "Риск-лог" in response.text
    assert "Риск-лог пока пуст" in response.text
    assert 'class="admin-nav-link active" href="/admin/risks"' in response.text


def test_admin_risks_lists_price_risk_log(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    connection = connect(db_path)
    try:
        init_database(connection)
        connection.execute(
            """
            INSERT INTO price_risk_log (
                exception_key, status, reason, code, unit,
                min_price, max_price, ratio, recommended_price, estimate_row
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "m||GESN01-01-001-01||0",
                "open",
                "ratio_exceeded",
                "GESN01-01-001-01",
                "m",
                100.0,
                350.0,
                3.5,
                120.0,
                42,
            ),
        )
        connection.commit()
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/risks")

    assert response.status_code == 200
    assert "m||GESN01-01-001-01||0" in response.text
    assert "ratio_exceeded" in response.text
    assert "GESN01-01-001-01" in response.text
    assert "open" in response.text
    assert "<td>100</td>" in response.text
    assert "<td>350</td>" in response.text
    assert "<td>3.5</td>" in response.text
    assert "<td>120</td>" in response.text
    assert "<td>42</td>" in response.text
