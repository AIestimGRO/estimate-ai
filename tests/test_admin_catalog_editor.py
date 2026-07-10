"""Tests for the editable admin catalog page."""

from fastapi.testclient import TestClient

from app.web.app import create_app
from core.storage import connect, init_database


def _seed_catalog_row(db_path, *, price: float = 100.0, quantity: float = 2.0) -> int:
    connection = connect(db_path)
    try:
        init_database(connection)
        cursor = connection.execute(
            "INSERT INTO catalog_sources(name, kind) VALUES (?, ?)",
            ("rnmc_zip_upload", "rnmc_zip"),
        )
        source_id = int(cursor.lastrowid)
        cursor = connection.execute(
            """
            INSERT INTO catalog_items (
                source_id, task_id, region, code, unit, quantity, work_name, price,
                total_price, labor_unit, labor_total, machine_labor_unit,
                machine_labor_total, regional_coefficient, source_region_folder,
                source_filename, source_row_number
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                source_id,
                "TASK-1",
                "Moscow",
                "GESN01-01-001-01",
                "m",
                quantity,
                "Test work",
                price,
                price * quantity,
                1.5,
                3.0,
                0.2,
                0.4,
                1.0,
                "Moscow",
                "catalog.xlsx",
                42,
            ),
        )
        row_id = int(cursor.lastrowid)
        connection.commit()
        return row_id
    finally:
        connection.close()


def test_admin_catalog_lists_and_filters_rows(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    _seed_catalog_row(db_path)

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/catalog?code=GESN01")

    assert response.status_code == 200
    assert "Каталог аналогов" in response.text
    assert "GESN01-01-001-01" in response.text
    assert "Test work" in response.text
    assert "catalog.xlsx" in response.text
    assert 'class="admin-nav-link active" href="/admin/catalog"' in response.text


def test_admin_catalog_updates_row_values(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    row_id = _seed_catalog_row(db_path)

    data = {
        "return_url": "/admin/catalog",
        "row_id": str(row_id),
        f"task_id_{row_id}": "TASK-2",
        f"region_{row_id}": "SPb",
        f"code_{row_id}": "GESN02-02-002-02",
        f"unit_{row_id}": "шт",
        f"quantity_{row_id}": "3,5",
        f"work_name_{row_id}": "Updated work",
        f"price_{row_id}": "120,25",
        f"total_price_{row_id}": "420,875",
        f"labor_unit_{row_id}": "2",
        f"labor_total_{row_id}": "7",
        f"machine_labor_unit_{row_id}": "0,5",
        f"machine_labor_total_{row_id}": "1,75",
        f"regional_coefficient_{row_id}": "1,2",
        f"source_region_folder_{row_id}": "SPb",
        f"source_filename_{row_id}": "updated.xlsx",
        f"source_row_number_{row_id}": "55",
    }

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post("/admin/catalog/update-row", data=data, follow_redirects=False)

    assert response.status_code == 303

    connection = connect(db_path)
    try:
        row = connection.execute(
            """
            SELECT task_id, region, code, unit, quantity, work_name, price,
                   total_price, regional_coefficient, source_filename, source_row_number
            FROM catalog_items WHERE id = ?
            """,
            (row_id,),
        ).fetchone()
    finally:
        connection.close()

    assert row["task_id"] == "TASK-2"
    assert row["region"] == "SPb"
    assert row["code"] == "GESN02-02-002-02"
    assert row["unit"] == "шт"
    assert row["quantity"] == 3.5
    assert row["work_name"] == "Updated work"
    assert row["price"] == 120.25
    assert row["total_price"] == 420.875
    assert row["regional_coefficient"] == 1.2
    assert row["source_filename"] == "updated.xlsx"
    assert row["source_row_number"] == 55


def test_admin_catalog_bulk_multiply_and_delete(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    row_id = _seed_catalog_row(db_path, price=100.0)

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/catalog/bulk",
            data={
                "return_url": "/admin/catalog",
                "selected_ids": str(row_id),
                "bulk_action": "update",
                "bulk_field": "price",
                "bulk_operation": "multiply",
                "bulk_value": "1,2",
            },
            follow_redirects=False,
        )

    assert response.status_code == 303
    connection = connect(db_path)
    try:
        price = connection.execute(
            "SELECT price FROM catalog_items WHERE id = ?",
            (row_id,),
        ).fetchone()["price"]
    finally:
        connection.close()
    assert price == 120.0

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/catalog/bulk",
            data={
                "return_url": "/admin/catalog",
                "selected_ids": str(row_id),
                "bulk_action": "delete",
            },
            follow_redirects=False,
        )

    assert response.status_code == 303
    connection = connect(db_path)
    try:
        count = connection.execute("SELECT COUNT(*) AS c FROM catalog_items").fetchone()["c"]
    finally:
        connection.close()
    assert count == 0


def test_admin_catalog_bulk_action_has_confirmation_guard(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    _seed_catalog_row(db_path)

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/catalog")

    assert response.status_code == 200
    assert "confirmCatalogBulkAction" in response.text
    assert 'onclick="return prepareCatalogBulkAction(this.form)"' in response.text
    assert "Вы уверены, что хотите" in response.text
    assert "Действие нельзя отменить автоматически" in response.text


def test_admin_catalog_clear_removes_catalog_and_import_log(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    _seed_catalog_row(db_path)

    connection = connect(db_path)
    try:
        source_id = connection.execute(
            "SELECT id FROM catalog_sources WHERE name = ?",
            ("rnmc_zip_upload",),
        ).fetchone()["id"]
        cursor = connection.execute(
            """
            INSERT INTO imported_files (
                source_id, region_folder, filename, status, filename_key
            ) VALUES (?, ?, ?, ?, ?)
            """,
            (source_id, "Moscow", "catalog.xlsx", "success", "catalog.xlsx"),
        )
        file_id = int(cursor.lastrowid)
        connection.execute(
            "INSERT INTO import_row_log(file_id, row_number, status, reason) VALUES (?, ?, ?, ?)",
            (file_id, 42, "rejected", "test"),
        )
        connection.commit()
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/catalog/clear",
            data={"confirmation": "clear_catalog"},
            follow_redirects=False,
        )

    assert response.status_code == 303
    assert "message=" in response.headers["location"]

    connection = connect(db_path)
    try:
        catalog_count = connection.execute("SELECT COUNT(*) AS c FROM catalog_items").fetchone()["c"]
        import_count = connection.execute("SELECT COUNT(*) AS c FROM imported_files").fetchone()["c"]
        row_log_count = connection.execute("SELECT COUNT(*) AS c FROM import_row_log").fetchone()["c"]
        source_count = connection.execute("SELECT COUNT(*) AS c FROM catalog_sources").fetchone()["c"]
    finally:
        connection.close()

    assert catalog_count == 0
    assert import_count == 0
    assert row_log_count == 0
    assert source_count == 1


def test_admin_catalog_clear_requires_confirmation(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    _seed_catalog_row(db_path)

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/catalog/clear",
            data={},
            follow_redirects=False,
        )

    assert response.status_code == 303
    assert "error=" in response.headers["location"]

    connection = connect(db_path)
    try:
        count = connection.execute("SELECT COUNT(*) AS c FROM catalog_items").fetchone()["c"]
    finally:
        connection.close()
    assert count == 1


def test_admin_catalog_bulk_submission_disables_row_editor_fields(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))
    _seed_catalog_row(db_path)

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/catalog")

    assert response.status_code == 200
    assert "prepareCatalogBulkAction" in response.text
    assert "tbody input:not(.catalog-row-check), tbody textarea" in response.text
    assert 'action="/admin/catalog/clear"' not in response.text
    assert "Очистить каталог для пересборки" not in response.text
