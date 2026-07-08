"""SQLite schema for Estimate AI."""

SCHEMA_VERSION = 5

DDL = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS schema_migrations (
    version INTEGER PRIMARY KEY
);

CREATE TABLE IF NOT EXISTS catalog_sources (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    kind TEXT NOT NULL DEFAULT 'excel_bulk',
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS catalog_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_id INTEGER NOT NULL REFERENCES catalog_sources(id) ON DELETE CASCADE,
    task_id TEXT NOT NULL,
    region TEXT NOT NULL DEFAULT '',
    code TEXT NOT NULL,
    unit TEXT NOT NULL,
    work_name TEXT NOT NULL DEFAULT '',
    price REAL NOT NULL,
    total_price REAL,
    labor_unit REAL,
    labor_total REAL,
    machine_labor_unit REAL,
    machine_labor_total REAL,
    regional_coefficient REAL,
    added_date TEXT,
    source_region_folder TEXT NOT NULL DEFAULT '',
    source_filename TEXT NOT NULL DEFAULT '',
    source_row_number INTEGER NOT NULL DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_catalog_items_source_id
    ON catalog_items(source_id);
CREATE INDEX IF NOT EXISTS idx_catalog_items_task_id
    ON catalog_items(task_id);
CREATE INDEX IF NOT EXISTS idx_catalog_items_code
    ON catalog_items(code);

CREATE TABLE IF NOT EXISTS imported_files (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_id INTEGER REFERENCES catalog_sources(id) ON DELETE SET NULL,
    region_folder TEXT NOT NULL DEFAULT '',
    filename TEXT NOT NULL,
    status TEXT NOT NULL,
    imported_at TEXT NOT NULL DEFAULT (datetime('now')),
    task_number TEXT NOT NULL DEFAULT '',
    rows_ok INTEGER NOT NULL DEFAULT 0,
    rows_rejected INTEGER NOT NULL DEFAULT 0,
    failure_reason TEXT NOT NULL DEFAULT '',
    filename_key TEXT NOT NULL DEFAULT '',
    legacy_note TEXT NOT NULL DEFAULT '',
    lsr_quarter TEXT NOT NULL DEFAULT '',
    planned_start TEXT NOT NULL DEFAULT '',
    planned_finish TEXT NOT NULL DEFAULT '',
    regional_coefficient REAL,
    UNIQUE(region_folder, filename)
);


CREATE TABLE IF NOT EXISTS import_row_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    file_id INTEGER NOT NULL REFERENCES imported_files(id) ON DELETE CASCADE,
    row_number INTEGER NOT NULL,
    status TEXT NOT NULL,
    reason TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_import_row_log_file_id
    ON import_row_log(file_id);

CREATE TABLE IF NOT EXISTS name_exclusion_rules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    enabled INTEGER NOT NULL DEFAULT 1,
    scope TEXT NOT NULL,
    match_mode TEXT NOT NULL,
    pattern TEXT NOT NULL,
    rule_group TEXT NOT NULL DEFAULT '',
    comment TEXT NOT NULL DEFAULT '',
    sort_order INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS task_color_entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    enabled INTEGER NOT NULL DEFAULT 1,
    task_number TEXT NOT NULL,
    reason TEXT NOT NULL DEFAULT '',
    comment TEXT NOT NULL DEFAULT '',
    sort_order INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS gesn_exceptions (
    exception_key TEXT PRIMARY KEY,
    approved_min REAL NOT NULL,
    approved_max REAL NOT NULL,
    last_range_update_date REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS price_risk_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    exception_key TEXT NOT NULL UNIQUE,
    status TEXT NOT NULL DEFAULT 'open',
    reason TEXT NOT NULL,
    code TEXT NOT NULL DEFAULT '',
    unit TEXT NOT NULL DEFAULT '',
    min_price REAL,
    max_price REAL,
    ratio REAL,
    recommended_price REAL,
    estimate_row INTEGER,
    first_seen_at TEXT NOT NULL DEFAULT (datetime('now')),
    last_seen_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_price_risk_log_status
    ON price_risk_log(status);
"""
