"""Tests for the read-only admin catalog sources page."""

from fastapi.testclient import TestClient

from app.web.app import create_app
from core.storage import connect, init_database


def test_admin_sources_shows_empty_state(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/sources")

    assert response.status_code == 200
    assert "\u0418\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u0438 \u0431\u0430\u0437\u044b" in response.text
    assert "\u0418\u0441\u0442\u043e\u0447\u043d\u0438\u043a\u0438 \u043f\u043e\u043a\u0430 \u043d\u0435 \u0434\u043e\u0431\u0430\u0432\u043b\u0435\u043d\u044b" in response.text


def test_admin_sources_lists_catalog_sources(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    connection = connect(db_path)
    try:
        init_database(connection)
        cursor = connection.execute(
            "INSERT INTO catalog_sources(name, kind) VALUES (?, ?)",
            ("RNMC 2026", "historical_excel"),
        )
        source_id = int(cursor.lastrowid)
        connection.execute(
            """
            INSERT INTO catalog_items (
                source_id, task_id, region, code, unit, work_name, price
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                source_id,
                "TASK-1",
                "Moscow",
                "GESN01-01-001-01",
                "m",
                "work",
                100.0,
            ),
        )
        connection.commit()
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/sources")

    assert response.status_code == 200
    assert "RNMC 2026" in response.text
    assert "historical_excel" in response.text
    assert "<td>1</td>" in response.text
    assert 'class="admin-nav-link active" href="/admin/sources"' in response.text
