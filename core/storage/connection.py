"""Database connection and initialization."""

from __future__ import annotations

import json
import os
import sqlite3
from pathlib import Path

from core.storage.schema import DDL, SCHEMA_VERSION

_CONFIG_RELATIVE = ("data", "config", "db.json")
_ENV_DB_PATH = "ESTIMATE_AI_DB_PATH"


def repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


def default_database_path(config_path: str | Path | None = None) -> Path:
    env_path = os.environ.get(_ENV_DB_PATH, "").strip()
    if env_path:
        return Path(env_path).resolve()

    path = Path(config_path) if config_path is not None else _default_config_path()
    if path.is_file():
        data = json.loads(path.read_text(encoding="utf-8"))
        configured = str(data.get("database_path") or "").strip()
        if configured:
            candidate = Path(configured)
            if not candidate.is_absolute():
                candidate = repo_root() / candidate
            return candidate.resolve()

    return (repo_root() / "data" / "estimate_ai.db").resolve()


def connect(database_path: str | Path | None = None) -> sqlite3.Connection:
    path = default_database_path() if database_path is None else Path(database_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def init_database(connection: sqlite3.Connection) -> None:
    if not _schema_is_current(connection):
        connection.executescript(DDL)
        _apply_additive_migrations(connection)
        connection.execute(
            "INSERT OR REPLACE INTO schema_migrations(version) VALUES (?)",
            (SCHEMA_VERSION,),
        )
        connection.commit()

    from core.storage.corrections import synchronize_catalog_corrections

    synchronize_catalog_corrections(connection)
    connection.commit()


def _apply_additive_migrations(connection: sqlite3.Connection) -> None:
    _ensure_column(connection, "imported_files", "filename_key", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "imported_files", "legacy_note", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "imported_files", "lsr_quarter", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "imported_files", "planned_start", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "imported_files", "planned_finish", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "imported_files", "regional_coefficient", "REAL")
    _ensure_column(connection, "catalog_items", "quantity", "REAL")
    _ensure_column(connection, "catalog_items", "price_original", "REAL")
    _ensure_column(connection, "catalog_items", "price_zlvl", "REAL")
    _ensure_column(connection, "catalog_items", "total_price", "REAL")
    _ensure_column(connection, "catalog_items", "labor_unit", "REAL")
    _ensure_column(connection, "catalog_items", "labor_total", "REAL")
    _ensure_column(connection, "catalog_items", "machine_labor_unit", "REAL")
    _ensure_column(connection, "catalog_items", "machine_labor_total", "REAL")
    _ensure_column(connection, "catalog_items", "regional_coefficient", "REAL")
    _ensure_column(connection, "catalog_items", "lsr_quarter", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "catalog_items", "planned_start", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "catalog_items", "planned_finish", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "tkp_sources", "details_version", "INTEGER NOT NULL DEFAULT 0")
    _ensure_column(connection, "tkp_items", "qty_source_text", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "tkp_items", "rnmc_line_total_no_vat", "REAL")
    _ensure_column(connection, "tkp_items", "winner_group_index", "INTEGER NOT NULL DEFAULT 0")
    _ensure_column(connection, "tkp_items", "winner_start_col", "INTEGER NOT NULL DEFAULT 0")
    _ensure_column(connection, "tkp_items", "winner_start_col_letter", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "tkp_items", "winner_unit_header", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "tkp_items", "winner_total_header", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "tkp_items", "version", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "tkp_items", "winner_method", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "tkp_items", "winner_block_name", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "tkp_items", "winner_block_uin", "TEXT NOT NULL DEFAULT ''")
    _ensure_column(connection, "tkp_items", "winner_block_total_vat", "REAL")
    _ensure_column(connection, "tkp_items", "winner_block_reason", "TEXT NOT NULL DEFAULT ''")
    connection.execute("UPDATE catalog_items SET price_original = price WHERE price_original IS NULL")
    connection.execute("UPDATE catalog_items SET price_zlvl = price WHERE price_zlvl IS NULL")
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_imported_files_filename_key "
        "ON imported_files(filename_key)"
    )
    connection.execute(
        "UPDATE imported_files SET filename_key = lower(filename) "
        "WHERE filename_key = ''"
    )
    _seed_default_highlight_reasons(connection)


def _seed_default_highlight_reasons(connection: sqlite3.Connection) -> None:
    """Seed the highlight-reason registry once, if it is still empty.

    The pre-existing "blue task" highlight had no explicit reason key; it
    gets TKP_PLUS1 with the same colour it always used (DDEBF7), plus the
    new FOT reason. Runs every init_database call but is a no-op once rows
    exist, so admin edits are never overwritten.
    """
    count_row = connection.execute(
        "SELECT COUNT(*) AS n FROM task_highlight_reasons"
    ).fetchone()
    if int(count_row["n"]) > 0:
        return
    connection.executemany(
        """
        INSERT INTO task_highlight_reasons (key, label, color_hex, enabled, sort_order)
        VALUES (?, ?, ?, 1, ?)
        """,
        [
            ("TKP_PLUS1", "\u0422\u041a\u041f +1%", "DDEBF7", 0),
            ("FOT", "\u0424\u041e\u0422", "E2EFDA", 1),
        ],
    )


def _ensure_column(
    connection: sqlite3.Connection,
    table_name: str,
    column_name: str,
    declaration: str,
) -> None:
    columns = {str(row["name"]) for row in connection.execute(f"PRAGMA table_info({table_name})")}
    if column_name not in columns:
        connection.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {declaration}")


def _schema_is_current(connection: sqlite3.Connection) -> bool:
    table = connection.execute(
        "SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'schema_migrations'"
    ).fetchone()
    if table is None:
        return False

    current = connection.execute(
        "SELECT version FROM schema_migrations ORDER BY version DESC LIMIT 1"
    ).fetchone()
    return current is not None and int(current["version"]) >= SCHEMA_VERSION


def _default_config_path() -> Path:
    return repo_root().joinpath(*_CONFIG_RELATIVE)
