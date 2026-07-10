"""Tests for the read-only admin import history page."""

from fastapi.testclient import TestClient

from app.web.app import create_app
from core.storage import connect, init_database


def test_admin_imports_shows_empty_state(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/imports")

    assert response.status_code == 200
    assert "\u0418\u043c\u043f\u043e\u0440\u0442\u044b \u0444\u0430\u0439\u043b\u043e\u0432" in response.text
    assert "\u0418\u043c\u043f\u043e\u0440\u0442\u044b \u043f\u043e\u043a\u0430 \u043d\u0435 \u0437\u0430\u043f\u0438\u0441\u0430\u043d\u044b" in response.text


def test_admin_imports_lists_imported_files(tmp_path, monkeypatch) -> None:
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
            INSERT INTO imported_files (
                source_id, region_folder, filename, status, task_number,
                rows_ok, rows_rejected, failure_reason
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                source_id,
                "Moscow",
                "catalog.xlsx",
                "success",
                "TASK-1",
                25,
                2,
                "",
            ),
        )
        connection.commit()
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/imports")

    assert response.status_code == 200
    assert "catalog.xlsx" in response.text
    assert "RNMC 2026" in response.text
    assert "historical_excel" in response.text
    assert "Moscow" in response.text
    assert "TASK-1" in response.text
    assert "success" in response.text
    assert "<td>25</td>" in response.text
    assert "<td>2</td>" in response.text
    assert 'class="admin-nav-link active" href="/admin/imports"' in response.text


def test_admin_imports_uses_single_regular_import_flow(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/imports")

    assert response.status_code == 200
    assert "Загрузка новых РНМЦ" in response.text
    assert "Загрузить и проверить" in response.text
    assert "Как проходит импорт" in response.text
    assert "История импортов" in response.text
    assert "Импорт старого File_Log.xlsx" not in response.text
    assert "Зафиксировать ZIP в журнале" not in response.text
    assert "Импортировать строки РНМЦ из ZIP в каталог" not in response.text
