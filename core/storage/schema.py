"""SQLite schema for Estimate AI."""

SCHEMA_VERSION = 14

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
    rows_excluded INTEGER NOT NULL DEFAULT 0,
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


CREATE TABLE IF NOT EXISTS manual_section_mappings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL,
    code_norm TEXT NOT NULL UNIQUE,
    section_code TEXT NOT NULL,
    enabled INTEGER NOT NULL DEFAULT 1,
    comment TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_manual_section_mappings_enabled
    ON manual_section_mappings(enabled);

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
    reserve_name TEXT NOT NULL DEFAULT '',
    reserve_inn TEXT NOT NULL DEFAULT '',
    reserve_uin TEXT NOT NULL DEFAULT '',
    reserve_total_no_vat REAL,
    reserve_total_vat REAL,
    reserve_method TEXT NOT NULL DEFAULT '',
    details_version INTEGER NOT NULL DEFAULT 0,
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
    qty_source_text TEXT NOT NULL DEFAULT '',
    rnmc_unit_price_no_vat REAL,
    rnmc_line_total_no_vat REAL,
    winner_unit_price_no_vat REAL,
    winner_line_total_no_vat REAL,
    winner_name TEXT NOT NULL DEFAULT '',
    winner_inn TEXT NOT NULL DEFAULT '',
    winner_uin TEXT NOT NULL DEFAULT '',
    winner_group_index INTEGER NOT NULL DEFAULT 0,
    winner_start_col INTEGER NOT NULL DEFAULT 0,
    winner_start_col_letter TEXT NOT NULL DEFAULT '',
    winner_unit_header TEXT NOT NULL DEFAULT '',
    winner_total_header TEXT NOT NULL DEFAULT '',
    task_no TEXT NOT NULL DEFAULT '',
    request_date TEXT NOT NULL DEFAULT '',
    version TEXT NOT NULL DEFAULT '',
    customer TEXT NOT NULL DEFAULT '',
    general_contractor TEXT NOT NULL DEFAULT '',
    procedure_name TEXT NOT NULL DEFAULT '',
    winner_method TEXT NOT NULL DEFAULT '',
    winner_block_name TEXT NOT NULL DEFAULT '',
    winner_block_uin TEXT NOT NULL DEFAULT '',
    winner_block_total_vat REAL,
    winner_block_reason TEXT NOT NULL DEFAULT '',
    reserve_unit_price_no_vat REAL,
    reserve_line_total_no_vat REAL,
    reserve_name TEXT NOT NULL DEFAULT '',
    reserve_inn TEXT NOT NULL DEFAULT '',
    reserve_uin TEXT NOT NULL DEFAULT '',
    reserve_group_index INTEGER NOT NULL DEFAULT 0,
    reserve_start_col INTEGER NOT NULL DEFAULT 0,
    reserve_start_col_letter TEXT NOT NULL DEFAULT '',
    reserve_unit_header TEXT NOT NULL DEFAULT '',
    reserve_total_header TEXT NOT NULL DEFAULT '',
    reserve_method TEXT NOT NULL DEFAULT '',
    wor_schema TEXT NOT NULL DEFAULT '',
    quality_flags TEXT NOT NULL DEFAULT ''
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

CREATE TABLE IF NOT EXISTS catalog_correction_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    seed_key TEXT UNIQUE,
    action TEXT NOT NULL DEFAULT 'update',
    status TEXT NOT NULL DEFAULT 'pending',
    target_item_id INTEGER,
    target_source_name TEXT NOT NULL DEFAULT '',
    target_source_filename TEXT NOT NULL DEFAULT '',
    target_source_row_number INTEGER NOT NULL DEFAULT 0,
    target_task_id TEXT NOT NULL DEFAULT '',
    target_code TEXT NOT NULL DEFAULT '',
    target_unit TEXT NOT NULL DEFAULT '',
    target_work_name TEXT NOT NULL DEFAULT '',
    reason TEXT NOT NULL,
    submitted_by TEXT NOT NULL,
    submitted_role TEXT NOT NULL,
    submitted_at TEXT NOT NULL DEFAULT (datetime('now')),
    reviewed_by TEXT NOT NULL DEFAULT '',
    reviewed_role TEXT NOT NULL DEFAULT '',
    reviewed_at TEXT,
    review_comment TEXT NOT NULL DEFAULT '',
    applied_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_catalog_correction_requests_status
    ON catalog_correction_requests(status);
CREATE INDEX IF NOT EXISTS idx_catalog_correction_requests_target
    ON catalog_correction_requests(
        target_source_name,
        target_source_filename,
        target_source_row_number
    );

CREATE TABLE IF NOT EXISTS catalog_correction_changes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    correction_id INTEGER NOT NULL
        REFERENCES catalog_correction_requests(id) ON DELETE CASCADE,
    field_name TEXT NOT NULL,
    value_type TEXT NOT NULL,
    old_value TEXT,
    new_value TEXT,
    UNIQUE(correction_id, field_name)
);

CREATE INDEX IF NOT EXISTS idx_catalog_correction_changes_request
    ON catalog_correction_changes(correction_id);

CREATE TABLE IF NOT EXISTS catalog_correction_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    correction_id INTEGER NOT NULL
        REFERENCES catalog_correction_requests(id) ON DELETE CASCADE,
    event_type TEXT NOT NULL,
    actor TEXT NOT NULL,
    actor_role TEXT NOT NULL,
    details TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_catalog_correction_events_request
    ON catalog_correction_events(correction_id, id);

CREATE TABLE IF NOT EXISTS app_users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    full_name TEXT NOT NULL,
    login TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL,
    is_active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_app_users_role
    ON app_users(role, is_active);

CREATE TABLE IF NOT EXISTS user_sessions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL REFERENCES app_users(id) ON DELETE CASCADE,
    token_hash TEXT NOT NULL UNIQUE,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    last_seen_at TEXT NOT NULL DEFAULT (datetime('now')),
    expires_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_user_sessions_user
    ON user_sessions(user_id);
CREATE INDEX IF NOT EXISTS idx_user_sessions_expiry
    ON user_sessions(expires_at);

CREATE TABLE IF NOT EXISTS processing_jobs (
    id TEXT PRIMARY KEY,
    owner_user_id INTEGER NOT NULL REFERENCES app_users(id) ON DELETE RESTRICT,
    estimate_filename TEXT NOT NULL,
    source_path TEXT NOT NULL,
    output_path TEXT NOT NULL,
    final_output_path TEXT NOT NULL DEFAULT '',
    sheet_title TEXT NOT NULL DEFAULT '',
    header_row INTEGER NOT NULL DEFAULT 0,
    coefficient REAL NOT NULL DEFAULT 1.0,
    region TEXT NOT NULL DEFAULT '',
    use_tkp_analogs INTEGER NOT NULL DEFAULT 0,
    status TEXT NOT NULL DEFAULT 'ready',
    total_rows INTEGER NOT NULL DEFAULT 0,
    matched_rows INTEGER NOT NULL DEFAULT 0,
    flagged_rows INTEGER NOT NULL DEFAULT 0,
    tkp_matched_rows INTEGER NOT NULL DEFAULT 0,
    column_schema_json TEXT NOT NULL DEFAULT '[]',
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now')),
    last_opened_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_processing_jobs_owner
    ON processing_jobs(owner_user_id, updated_at);
CREATE INDEX IF NOT EXISTS idx_processing_jobs_status
    ON processing_jobs(status, updated_at);

CREATE TABLE IF NOT EXISTS processing_rows (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    job_id TEXT NOT NULL REFERENCES processing_jobs(id) ON DELETE CASCADE,
    row_index INTEGER NOT NULL,
    excel_row_number INTEGER NOT NULL,
    values_json TEXT NOT NULL,
    metadata_json TEXT NOT NULL DEFAULT '{}',
    UNIQUE(job_id, row_index)
);

CREATE INDEX IF NOT EXISTS idx_processing_rows_job
    ON processing_rows(job_id, row_index);

CREATE TABLE IF NOT EXISTS processing_cell_overrides (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    job_id TEXT NOT NULL REFERENCES processing_jobs(id) ON DELETE CASCADE,
    row_id INTEGER NOT NULL REFERENCES processing_rows(id) ON DELETE CASCADE,
    column_index INTEGER NOT NULL,
    original_value_json TEXT,
    current_value_json TEXT,
    editor_user_id INTEGER NOT NULL REFERENCES app_users(id) ON DELETE RESTRICT,
    revision INTEGER NOT NULL DEFAULT 1,
    updated_at TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE(job_id, row_id, column_index)
);

CREATE INDEX IF NOT EXISTS idx_processing_cell_overrides_job
    ON processing_cell_overrides(job_id, row_id, column_index);

CREATE TABLE IF NOT EXISTS processing_edit_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    override_id INTEGER NOT NULL
        REFERENCES processing_cell_overrides(id) ON DELETE CASCADE,
    job_id TEXT NOT NULL REFERENCES processing_jobs(id) ON DELETE CASCADE,
    row_id INTEGER NOT NULL REFERENCES processing_rows(id) ON DELETE CASCADE,
    column_index INTEGER NOT NULL,
    old_value_json TEXT,
    new_value_json TEXT,
    actor_user_id INTEGER NOT NULL REFERENCES app_users(id) ON DELETE RESTRICT,
    client_change_id TEXT NOT NULL UNIQUE,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_processing_edit_events_job
    ON processing_edit_events(job_id, created_at);

CREATE TABLE IF NOT EXISTS specialist_change_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    job_id TEXT NOT NULL REFERENCES processing_jobs(id) ON DELETE CASCADE,
    row_id INTEGER REFERENCES processing_rows(id) ON DELETE SET NULL,
    excel_row_number INTEGER,
    column_index INTEGER,
    request_type TEXT NOT NULL,
    task_number TEXT NOT NULL DEFAULT '',
    comment TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'pending',
    submitted_by INTEGER NOT NULL REFERENCES app_users(id) ON DELETE RESTRICT,
    submitted_at TEXT NOT NULL DEFAULT (datetime('now')),
    reviewed_by INTEGER REFERENCES app_users(id) ON DELETE SET NULL,
    reviewed_at TEXT,
    review_comment TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_specialist_change_requests_status
    ON specialist_change_requests(status, submitted_at);
CREATE INDEX IF NOT EXISTS idx_specialist_change_requests_job
    ON specialist_change_requests(job_id, submitted_at);

CREATE TABLE IF NOT EXISTS user_preferences (
    user_id INTEGER NOT NULL REFERENCES app_users(id) ON DELETE CASCADE,
    preference_key TEXT NOT NULL,
    value_json TEXT NOT NULL,
    updated_at TEXT NOT NULL DEFAULT (datetime('now')),
    PRIMARY KEY(user_id, preference_key)
);

CREATE TABLE IF NOT EXISTS audit_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    actor_user_id INTEGER REFERENCES app_users(id) ON DELETE SET NULL,
    event_type TEXT NOT NULL,
    entity_type TEXT NOT NULL,
    entity_id TEXT NOT NULL DEFAULT '',
    job_id TEXT NOT NULL DEFAULT '',
    details TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_audit_events_job
    ON audit_events(job_id, created_at);
CREATE INDEX IF NOT EXISTS idx_audit_events_actor
    ON audit_events(actor_user_id, created_at);
"""
