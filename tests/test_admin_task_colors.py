"""Tests for the admin task colors page."""

from fastapi.testclient import TestClient

from app.web.app import create_app
from core.storage import (
    connect,
    init_database,
    list_task_color_entries,
    list_task_highlight_reasons,
)


def test_admin_task_colors_shows_empty_state(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/task-colors")

    assert response.status_code == 200
    assert "Синие задачи" in response.text
    assert "Синие задачи пока не записаны" in response.text
    assert 'action="/admin/task-colors/add"' in response.text
    assert 'class="admin-nav-link active" href="/admin/task-colors"' in response.text


def test_admin_task_colors_lists_entries(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    connection = connect(db_path)
    try:
        init_database(connection)
        connection.execute(
            """
            INSERT INTO task_color_entries (
                enabled, task_number, reason, comment, sort_order
            ) VALUES (?, ?, ?, ?, ?)
            """,
            (
                1,
                "TASK-123",
                "manual_review",
                "show as blue analog column",
                0,
            ),
        )
        connection.execute(
            """
            INSERT INTO task_color_entries (
                enabled, task_number, reason, comment, sort_order
            ) VALUES (?, ?, ?, ?, ?)
            """,
            (
                0,
                "TASK-999",
                "disabled",
                "kept for history",
                1,
            ),
        )
        connection.commit()
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/task-colors")

    assert response.status_code == 200
    assert "TASK-123" in response.text
    assert "manual_review" in response.text
    assert "show as blue analog column" in response.text
    assert "TASK-999" in response.text
    assert "disabled" in response.text
    assert "kept for history" in response.text
    assert "<td>да</td>" in response.text
    assert "<td>нет</td>" in response.text
    assert 'action="/admin/task-colors/toggle"' in response.text
    assert "Выключить" in response.text
    assert "Включить" in response.text
    assert 'class="admin-nav-link active" href="/admin/task-colors"' in response.text


def test_admin_task_colors_adds_entry(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/task-colors/add",
            data={
                "task_number": " TASK-456 ",
                "reason": "manual_review",
                "comment": "added from admin",
            },
            follow_redirects=False,
        )

    assert response.status_code == 303
    assert response.headers["location"] == "/admin/task-colors"

    connection = connect(db_path)
    try:
        entries = list_task_color_entries(connection)
    finally:
        connection.close()

    assert len(entries) == 1
    assert entries[0].enabled is True
    assert entries[0].task_number == "TASK-456"
    assert entries[0].reason == "manual_review"
    assert entries[0].comment == "added from admin"


def test_admin_task_colors_rejects_empty_task_number(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/task-colors/add",
            data={"task_number": "   ", "reason": "manual_review", "comment": ""},
        )

    assert response.status_code == 400
    assert "Номер задачи обязателен" in response.text


def test_admin_task_colors_toggle_disables_entry(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    connection = connect(db_path)
    try:
        init_database(connection)
        connection.execute(
            """
            INSERT INTO task_color_entries (
                enabled, task_number, reason, comment, sort_order
            ) VALUES (?, ?, ?, ?, ?)
            """,
            (1, "TASK-777", "manual_review", "temporary blue task", 0),
        )
        connection.commit()
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/task-colors/toggle",
            data={"task_number": "TASK-777", "enabled": "0"},
            follow_redirects=False,
        )

    assert response.status_code == 303
    assert response.headers["location"] == "/admin/task-colors"

    connection = connect(db_path)
    try:
        entries = list_task_color_entries(connection)
    finally:
        connection.close()

    assert len(entries) == 1
    assert entries[0].enabled is False


def test_admin_task_colors_page_seeds_default_reasons(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/task-colors")

    assert response.status_code == 200
    assert "\u0422\u041a\u041f +1%" in response.text
    assert "\u0424\u041e\u0422" in response.text
    assert 'action="/admin/task-colors/reasons/add"' in response.text
    assert '<select name="reason" required>' in response.text


def test_admin_task_colors_add_reason(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/task-colors/reasons/add",
            data={"key": "custom", "label": "\u041e\u0441\u043e\u0431\u0430\u044f", "color_hex": "#AABBCC"},
            follow_redirects=False,
        )

    assert response.status_code == 303
    assert response.headers["location"] == "/admin/task-colors"

    connection = connect(db_path)
    try:
        reasons = list_task_highlight_reasons(connection)
    finally:
        connection.close()

    added = [r for r in reasons if r.key == "CUSTOM"]
    assert len(added) == 1
    assert added[0].label == "\u041e\u0441\u043e\u0431\u0430\u044f"
    assert added[0].color_hex == "AABBCC"
    assert added[0].enabled is True


def test_admin_task_colors_add_reason_rejects_bad_color(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/task-colors/reasons/add",
            data={"key": "bad", "label": "Bad", "color_hex": "not-a-color"},
        )

    assert response.status_code == 400
    assert "\u041f\u0440\u0438\u0447\u0438\u043d\u0430 \u043d\u0435 \u0441\u043e\u0445\u0440\u0430\u043d\u0435\u043d\u0430" in response.text


def test_admin_task_colors_toggle_reason_disables_it(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        client.get("/admin/task-colors")  # triggers init_database + default seeding
        response = client.post(
            "/admin/task-colors/reasons/toggle",
            data={"key": "FOT", "enabled": "0"},
            follow_redirects=False,
        )

    assert response.status_code == 303

    connection = connect(db_path)
    try:
        reasons = list_task_highlight_reasons(connection)
    finally:
        connection.close()

    fot = next(r for r in reasons if r.key == "FOT")
    assert fot.enabled is False


def test_admin_task_colors_toggle_reason_missing_key_returns_404(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/task-colors/reasons/toggle",
            data={"key": "DOES-NOT-EXIST", "enabled": "0"},
        )

    assert response.status_code == 404
