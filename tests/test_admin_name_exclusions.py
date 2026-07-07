"""Tests for editing name exclusion rules in the admin UI."""

from fastapi.testclient import TestClient

from app.web.app import create_app
from core.exclusions import is_name_excluded
from core.storage import connect, init_database, list_name_exclusion_rules


def test_admin_name_exclusions_empty_page_has_add_form(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/name-exclusions")

    assert response.status_code == 200
    assert "Исключения по наименованиям" in response.text
    assert "Правила исключений пока не записаны" in response.text
    assert 'action="/admin/name-exclusions/add"' in response.text
    assert 'name="scope"' in response.text
    assert 'name="match_mode"' in response.text
    assert 'name="pattern"' in response.text


def test_admin_name_exclusions_lists_rules_with_toggle_buttons(tmp_path, monkeypatch) -> None:
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
            (1, "CATALOG", "ALL_WORDS", "temporary|work", "noise", "sample", 0),
        )
        connection.execute(
            """
            INSERT INTO name_exclusion_rules (
                enabled, scope, match_mode, pattern, rule_group, comment, sort_order
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (0, "SMETA", "CONTAINS", "skip this", "legacy", "disabled", 1),
        )
        connection.commit()
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.get("/admin/name-exclusions")

    assert response.status_code == 200
    assert "CATALOG" in response.text
    assert "temporary|work" in response.text
    assert "skip this" in response.text
    assert "Выключить" in response.text
    assert "Включить" in response.text
    assert 'action="/admin/name-exclusions/toggle"' in response.text


def test_admin_name_exclusions_adds_rule(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/name-exclusions/add",
            data={
                "scope": " catalog ",
                "match_mode": " all_words ",
                "pattern": " temporary|work ",
                "rule_group": "noise",
                "comment": "added from admin",
            },
            follow_redirects=False,
        )

    assert response.status_code == 303
    assert response.headers["location"] == "/admin/name-exclusions"

    connection = connect(db_path)
    try:
        rules = list_name_exclusion_rules(connection)
    finally:
        connection.close()

    assert len(rules) == 1
    assert rules[0].enabled is True
    assert rules[0].scope == "CATALOG"
    assert rules[0].match_mode == "ALL_WORDS"
    assert rules[0].pattern == "temporary|work"
    assert rules[0].group == "noise"
    assert rules[0].comment == "added from admin"
    assert is_name_excluded(rules, "CATALOG", "temporary metal work") is True


def test_admin_name_exclusions_rejects_empty_pattern(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/name-exclusions/add",
            data={
                "scope": "BOTH",
                "match_mode": "ALL_WORDS",
                "pattern": "   ",
                "rule_group": "noise",
                "comment": "",
            },
        )

    assert response.status_code == 400
    assert "Pattern обязателен" in response.text


def test_admin_name_exclusions_rejects_invalid_scope(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/name-exclusions/add",
            data={
                "scope": "UNKNOWN",
                "match_mode": "ALL_WORDS",
                "pattern": "temporary",
                "rule_group": "noise",
                "comment": "",
            },
        )

    assert response.status_code == 400
    assert "Scope должен быть SMETA, CATALOG или BOTH" in response.text


def test_admin_name_exclusions_toggle_disables_rule(tmp_path, monkeypatch) -> None:
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
            (1, "BOTH", "CONTAINS", "temporary", "noise", "admin", 0),
        )
        connection.commit()
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/name-exclusions/toggle",
            data={
                "scope": "BOTH",
                "match_mode": "CONTAINS",
                "pattern": "temporary",
                "enabled": "0",
            },
            follow_redirects=False,
        )

    assert response.status_code == 303
    assert response.headers["location"] == "/admin/name-exclusions"

    connection = connect(db_path)
    try:
        rules = list_name_exclusion_rules(connection)
    finally:
        connection.close()

    assert len(rules) == 1
    assert rules[0].enabled is False


def test_admin_name_exclusions_toggle_reports_missing_rule(tmp_path, monkeypatch) -> None:
    db_path = tmp_path / "estimate_ai.db"
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    with TestClient(create_app(base_dir=tmp_path / "work")) as client:
        response = client.post(
            "/admin/name-exclusions/toggle",
            data={
                "scope": "BOTH",
                "match_mode": "CONTAINS",
                "pattern": "missing",
                "enabled": "0",
            },
        )

    assert response.status_code == 404
    assert "Правило для изменения не найдено" in response.text
