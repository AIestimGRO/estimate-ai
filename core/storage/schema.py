"""SQLite schema for Estimate AI."""

SCHEMA_VERSION = 9

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
    quantity REAL,
    work_name TEXT NOT NULL DEFAULT '',
    price REAL NOT NULL,
    price_original REAL,
    price_zlvl REAL,
    total_price REAL,
    labor_unit REAL,
    labor_total REAL,
    machine_labor_unit REAL,
    machine_labor_total REAL,
    regional_coefficient REAL,
    lsr_quarter TEXT NOT NULL DEFAULT '',
    planned_start TEXT NOT NULL DEFAULT '',
    planned_finish TEXT NOT NULL DEFAULT '',
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

CREATE TABLE IF NOT EXISTS task_highlight_reasons (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    key TEXT NOT NULL UNIQUE,
    label TEXT NOT NULL,
    color_hex TEXT NOT NULL,
    enabled INTEGER NOT NULL DEFAULT 1,
    sort_order INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS tkp_sources (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id TEXT NOT NULL DEFAULT '',
    file_path TEXT NOT NULL DEFAULT '',
    file_name TEXT NOT NULL,
    modified_date TEXT NOT NULL DEFAULT '',
    sheet_name TEXT NOT NULL DEFAULT '',
    parse_status TEXT NOT NULL DEFAULT '',
    parse_message TEXT NOT NULL DEFAULT '',
    task_no TEXT NOT NULL DEFAULT '',
    request_date TEXT NOT NULL DEFAULT '',
    customer TEXT NOT NULL DEFAULT '',
    general_contractor TEXT NOT NULL DEFAULT '',
    procedure_name TEXT NOT NULL DEFAULT '',
    winner_name TEXT NOT NULL DEFAULT '',
    winner_inn TEXT NOT NULL DEFAULT '',
    winner_uin TEXT NOT NULL DEFAULT '',
    winner_total_no_vat REAL,
    winner_total_vat REAL,
    rnmc_total_no_vat REAL,
    item_count INTEGER NOT NULL DEFAULT 0,
    imported_at TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE(file_name, modified_date)
);

CREATE TABLE IF NOT EXISTS tkp_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_id INTEGER NOT NULL REFERENCES tkp_sources(id) ON DELETE CASCADE,
    source_row INTEGER NOT NULL DEFAULT 0,
    section_code TEXT NOT NULL DEFAULT '',
    section_name TEXT NOT NULL DEFAULT '',
    subsection_name TEXT NOT NULL DEFAULT '',
    item_code TEXT NOT NULL DEFAULT '',
    item_name TEXT NOT NULL,
    unit TEXT NOT NULL DEFAULT '',
    qty REAL,
    rnmc_unit_price_no_vat REAL,
    winner_unit_price_no_vat REAL,
    winner_line_total_no_vat REAL,
    winner_name TEXT NOT NULL DEFAULT '',
    winner_inn TEXT NOT NULL DEFAULT '',
    winner_uin TEXT NOT NULL DEFAULT '',
    task_no TEXT NOT NULL DEFAULT '',
    request_date TEXT NOT NULL DEFAULT '',
    customer TEXT NOT NULL DEFAULT '',
    general_contractor TEXT NOT NULL DEFAULT '',
    procedure_name TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_tkp_items_source_id
    ON tkp_items(source_id);
CREATE INDEX IF NOT EXISTS idx_tkp_items_item_name
    ON tkp_items(item_name);

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
