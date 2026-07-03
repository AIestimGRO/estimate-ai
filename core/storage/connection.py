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
    if _schema_is_current(connection):
        return

    connection.executescript(DDL)
    connection.execute(
        "INSERT OR REPLACE INTO schema_migrations(version) VALUES (?)",
        (SCHEMA_VERSION,),
    )
    connection.commit()


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
