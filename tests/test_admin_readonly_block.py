"""Tests for the read-only admin block pages."""

from fastapi.testclient import TestClient

from app.web.app import create_app
from core.storage import connect, init_database


def test_admin_approvals_shows_empty_open_risks(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/approvals")

    assert response.status_code == 200
    assert "Одобрение диапазонов" in response.text
    assert "Открытых рисков для одобрения пока нет" in response.text
    assert 'class="admin-nav-link active" href="/admin/approvals"' in response.text


def test_admin_approvals_lists_only_open_risks(tmp_path, monkeypatch) -> None:
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
                "spread_limit",
                "GESN01-01-001-01",
                "m",
                100.0,
                350.0,
                3.5,
                200.0,
                12,
            ),
        )
        connection.execute(
            """
            INSERT INTO price_risk_log (
                exception_key, status, reason, code, unit,
                min_price, max_price, ratio, recommended_price, estimate_row
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "pcs||GESN02-02-002-02||1",
                "approved",
                "spread_limit",
                "GESN02-02-002-02",
                "pcs",
                10.0,
                20.0,
                2.0,
                15.0,
                14,
            ),
        )
        connection.commit()
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/approvals")

    assert response.status_code == 200
    assert "m||GESN01-01-001-01||0" in response.text
    assert "GESN01-01-001-01" in response.text
    assert "spread_limit" in response.text
    assert "pcs||GESN02-02-002-02||1" not in response.text


def test_admin_name_exclusions_shows_empty_state(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/name-exclusions")

    assert response.status_code == 200
    assert "Исключения по наименованиям" in response.text
    assert "Правила исключений пока не записаны" in response.text
    assert 'class="admin-nav-link active" href="/admin/name-exclusions"' in response.text


def test_admin_name_exclusions_lists_rules(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    connection = connect(db_path)
    try:
        init_database(connection)
        connection.execute(
            """
            INSERT INTO name_exclusion_rules (
                enabled, scope, match_mode, pattern, rule_group, comment, sort_order
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                1,
                "CATALOG",
                "ALL_WORDS",
                "temporary|work",
                "noise",
                "sample comment",
                0,
            ),
        )
        connection.commit()
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/name-exclusions")

    assert response.status_code == 200
    assert "CATALOG" in response.text
    assert "ALL_WORDS" in response.text
    assert "temporary|work" in response.text
    assert "noise" in response.text
    assert "sample comment" in response.text
    assert "<td>да</td>" in response.text


def test_admin_settings_shows_diagnostics(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    connection = connect(db_path)
    try:
        init_database(connection)
        connection.execute(
            "INSERT INTO catalog_sources (name, kind) VALUES (?, ?)",
            ("main", "excel_bulk"),
        )
        source_id = int(connection.execute(
            "SELECT id FROM catalog_sources WHERE name = ?", ("main",)
        ).fetchone()["id"])
        connection.execute(
            """
            INSERT INTO catalog_items (
                source_id, task_id, region, code, unit, work_name, price
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (source_id, "TASK-1", "Moscow", "GESN01", "m", "Work", 100.0),
        )
        connection.execute(
            """
            INSERT INTO task_color_entries (
                enabled, task_number, reason, comment, sort_order
            ) VALUES (?, ?, ?, ?, ?)
            """,
            (1, "TASK-1", "manual", "comment", 0),
        )
        connection.commit()
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/settings")

    assert response.status_code == 200
    assert "Настройки и диагностика" in response.text
    assert "Database path" in response.text
    assert str(db_path) in response.text
    assert "Catalog rows" in response.text
    assert "Task color entries" in response.text
    assert "Name exclusion rules" in response.text
    assert 'class="admin-nav-link active" href="/admin/settings"' in response.text
