"""Tests for the admin task colors page."""

from fastapi.testclient import TestClient

from app.web.app import create_app
from core.storage import connect, init_database, list_task_color_entries


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
