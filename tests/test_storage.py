"""Tests for SQLite storage layer."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from core.catalog import BuildCatalog
from core.exclusions import is_name_excluded
from core.macro_workbook import NAME_EXCLUSIONS_SHEET, load_all_rules_from_workbook
from core.storage import (
    connect,
    import_catalog_from_excel,
    import_rules_from_workbook,
    init_database,
    list_catalog_rows,
    list_catalog_sources,
    list_name_exclusion_rules,
    list_task_color_entries,
)
CODE = "\u0413\u042d\u0421\u041d01-01-001-01"
METER = "\u043c"


def _write_catalog_workbook(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "\u041a\u0430\u0442\u0430\u043b\u043e\u0433"
    worksheet.cell(row=4, column=2, value="TASK-100")
    worksheet.cell(row=4, column=3, value="\u0440\u0430\u0431\u043e\u0442\u0430")
    worksheet.cell(row=4, column=4, value=METER)
    worksheet.cell(row=4, column=7, value=125.5)
    worksheet.cell(row=4, column=14, value=CODE)
    workbook.save(path)
    workbook.close()


def _write_rules_workbook(path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = NAME_EXCLUSIONS_SHEET
    worksheet.cell(row=2, column=1, value=1)
    worksheet.cell(row=2, column=2, value="SMETA")
    worksheet.cell(row=2, column=3, value="ALL_WORDS")
    worksheet.cell(row=2, column=4, value="\u043c\u043c|\u0438\u0437\u043c\u0435\u043d")
    worksheet.cell(row=3, column=1, value=0)
    worksheet.cell(row=3, column=2, value="BOTH")
    worksheet.cell(row=3, column=3, value="ALL_WORDS")
    worksheet.cell(row=3, column=4, value="\u0441\u043b\u043e\u0439")
    worksheet.cell(row=2, column=8, value=1)
    worksheet.cell(row=2, column=9, value="999")
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)
    workbook.close()


def test_init_database_creates_schema() -> None:
    connection = connect(":memory:")
    try:
        init_database(connection)
        tables = {
            row[0]
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            ).fetchall()
        }
        assert "catalog_sources" in tables
        assert "catalog_items" in tables
        assert "name_exclusion_rules" in tables
    finally:
        connection.close()


def test_import_catalog_from_excel_round_trip(tmp_path: Path) -> None:
    catalog_path = tmp_path / "catalog.xlsx"
    _write_catalog_workbook(catalog_path)

    connection = connect(":memory:")
    try:
        init_database(connection)
        result = import_catalog_from_excel(connection, catalog_path, source_name="main")
        rows = list_catalog_rows(connection, source_name="main")
    finally:
        connection.close()

    assert result.rows_imported == 1
    assert len(rows) == 1
    assert rows[0].code == CODE
    assert float(rows[0].price) == pytest.approx(125.5)


def test_import_catalog_replace_clears_previous_source(tmp_path: Path) -> None:
    first = tmp_path / "first.xlsx"
    second = tmp_path / "second.xlsx"
    _write_catalog_workbook(first)

    second.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "\u041a\u0430\u0442\u0430\u043b\u043e\u0433"
    worksheet.cell(row=4, column=2, value="TASK-200")
    worksheet.cell(row=4, column=3, value="\u0440\u0430\u0431\u043e\u0442\u0430")
    worksheet.cell(row=4, column=4, value=METER)
    worksheet.cell(row=4, column=7, value=200)
    worksheet.cell(row=4, column=14, value=CODE)
    workbook.save(second)
    workbook.close()

    connection = connect(":memory:")
    try:
        init_database(connection)
        import_catalog_from_excel(connection, first, source_name="main")
        import_catalog_from_excel(connection, second, source_name="main", replace=True)
        rows = list_catalog_rows(connection, source_name="main")
        sources = list_catalog_sources(connection)
    finally:
        connection.close()

    assert len(sources) == 1
    assert len(rows) == 1
    assert rows[0].task_id == "TASK-200"


def test_imported_catalog_builds_matching_catalog(tmp_path: Path) -> None:
    catalog_path = tmp_path / "catalog.xlsx"
    _write_catalog_workbook(catalog_path)

    connection = connect(":memory:")
    try:
        init_database(connection)
        import_catalog_from_excel(connection, catalog_path)
        rows = list_catalog_rows(connection)
        catalog = BuildCatalog(rows, [])
    finally:
        connection.close()

    assert len(catalog) == 1


def test_import_rules_from_workbook(tmp_path: Path) -> None:
    xlsm = tmp_path / "macro.xlsm"
    _write_rules_workbook(xlsm)

    connection = connect(":memory:")
    try:
        init_database(connection)
        rule_count, color_count = import_rules_from_workbook(connection, xlsm)
        rules = list_name_exclusion_rules(connection)
        colors = list_task_color_entries(connection)
    finally:
        connection.close()

    assert rule_count == 2
    assert color_count == 1
    enabled = [rule for rule in rules if rule.enabled]
    assert len(enabled) == 1
    assert is_name_excluded(enabled, "SMETA", "\u0442\u0435\u043a\u0441\u0442 \u043c\u043c \u0438\u0437\u043c\u0435\u043d")


def test_load_all_rules_from_workbook_includes_disabled(tmp_path: Path) -> None:
    xlsm = tmp_path / "macro.xlsm"
    _write_rules_workbook(xlsm)

    rules, _colors = load_all_rules_from_workbook(xlsm)

    assert len(rules) == 2
    assert any(not rule.enabled for rule in rules)


def test_cli_import_catalog(tmp_path: Path, monkeypatch) -> None:
    db_path = tmp_path / "test.db"
    catalog_path = tmp_path / "catalog.xlsx"
    _write_catalog_workbook(catalog_path)
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(db_path))

    from app.cli.__main__ import main

    assert main(["init-db"]) == 0
    assert main(["import-catalog", str(catalog_path)]) == 0
    assert main(["status"]) == 0
    assert db_path.is_file()
