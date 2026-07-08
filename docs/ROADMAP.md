# Estimate AI Roadmap

## Goal

Build a web service for matching estimate/BOQ rows against a historical RNMC
catalog, producing analogs, recommended prices, risk flags, approval history,
RNMC import history, and downloadable Excel results.

## Product direction

The target product is a web service:

- import and maintain RNMC catalog data;
- upload an estimate file;
- run deterministic analog matching;
- review matched analogs and price-spread risks;
- approve or widen GESN exception ranges;
- manage task-color and name-exclusion rules;
- inspect import history and rejected rows;
- download a completed Excel result.

## Current state (2026-07)

Core VBA logic has been ported into a tested Python library. The project now has
SQLite persistence, a FastAPI web UI, an admin UI, and a working RNMC ZIP import
flow.

### Done — deterministic core

- `core/normalize.py` — code/unit normalization, demolition detection, search key.
- `core/exclusions.py` — name exclusion rules and task color metadata.
- `core/catalog.py` — catalog construction and demolition-aware 4% dedup.
- `core/matching.py` — estimate row to catalog matching with reason codes.
- `core/risk.py` — ratio risk and approved-range override.
- `core/approval.py` — create/widen `GesnException` ranges.
- `core/sections.py` — GESN prefix extraction and section code resolution.
- `core/pricing.py` — average price formula and regional coefficient.
- `core/excel_io.py` — read catalog and estimate Excel rows (`data_only=True`).
- `core/layout.py` — flexible layout resolution.
- `core/excel_writer.py` — write run result into a `WA` copy.
- `core/macro_workbook.py` — load `Name_Exclusions` / task colours from xlsm.

### Done — storage and CLI

- SQLite schema for `catalog_items`, `catalog_sources`, `imported_files`,
  `import_row_log`, `name_exclusion_rules`, `task_color_entries`,
  `gesn_exceptions`, and `price_risk_log`.
- CLI commands for database initialization, catalog import, rule import, and
  status checks.
- Risk approval storage: open risks are kept in `price_risk_log`; approved
  ranges are kept in `gesn_exceptions`.

### Done — web UI and admin UI

- Main web flow: upload estimate/catalog, run matching, choose sheet when needed,
  download WA workbook.
- `/admin/sources` — catalog source overview.
- `/admin/imports` — RNMC import dashboard and control center.
- `/admin/risks` — price risk log.
- `/admin/approvals` — approve open price risks into `gesn_exceptions`.
- `/admin/gesn-exceptions` — approved ranges.
- `/admin/task-colors` — view/edit task color metadata.
- `/admin/name-exclusions` — view/edit exclusion rules.
- `/admin/settings` — diagnostics for database and counts.

### Done — RNMC import workflow

- Import legacy `File_Log.xlsx` into `imported_files`.
- Dedup/skip by normalized filename only; duplicate filenames across folders are
  marked as `duplicate_name`.
- ZIP dry-run without writes.
- ZIP import-log recording without catalog rows.
- ZIP row preview for `.xlsx` / `.xlsm`.
- Real ZIP import into `catalog_items`.
- Per-file detail page with metadata, imported rows, rejected-row log, and retry
  unlock for `failed` / `no_data`.
- RNMC value mapping for unit price without VAT, total cost without VAT, and
  labor columns, including VAT normalization and average-value exclusion.

See `docs/RNMC_IMPORT.md` for the import specification.

### Done — validation

- Automated pytest coverage for core, storage, web, admin, and RNMC import flows.
- Real-file comparison scripts remain in `scripts/`.

## Completed milestones

- [x] Port core VBA matching/pricing/risk logic to tested Python modules.
- [x] End-to-end local run: catalog + estimate -> structured result.
- [x] Excel writer: analogs, average formula, `/KR`, section, colours.
- [x] Flexible layout read and multi-sheet choice in web UI.
- [x] SQLite catalog storage and CLI import.
- [x] Minimal web UI: upload -> run -> download WA.
- [x] Admin UI shell and read-only dashboards.
- [x] Admin edit workflows for approvals, task colors, and name exclusions.
- [x] RNMC legacy FileLog and ZIP import into the catalog.
- [x] RNMC import control center with per-file details and rejected rows.
- [x] Initial automatic detection of `lsr_quarter`, planned start, and planned
  finish from RNMC workbooks.
- [x] Strict RNMC value-column mapping with VAT normalization and labor fields.

## Next milestone — RNMC import quality automation

1. Add regional coefficient extraction/storage from the RNMC consolidation block.
2. Improve rejected-row diagnostics and export/download of rejected rows.
3. Add `.xls` support if real incoming files still require it.
4. Decide whether to store original uploaded workbooks for true one-click retry
   without re-uploading the ZIP.
5. Add a duplicate-name review workflow for `duplicate_name` records.

## Later web service path

- Refactor upload/run into proper run resources (`POST /runs`, `GET /runs/{id}`).
- Store run history and allow downloading previous WA results.
- Add authentication and roles when explicitly approved.
- Add watched-folder or scheduled import only after manual ZIP import is stable.
- Improve detected-layout Excel output formatting, including full R13 and blue
  task-color tint.

## Future phase — multi-source analogs

Beyond the historical RNMC catalog, analogs may later be drawn from additional
sources and shown in leading analog columns. This is explicitly deferred until
the deterministic exact-match core and RNMC import are proven on real data.

Do not add semantic/AI matching inside matching/pricing. Any future semantic
source must be isolated behind a separate, cache-backed, reproducible layer.

## Working rules

- Speak with the user in Russian.
- Keep code, filenames, identifiers, and code comments English/ASCII only.
- Do not invent prices.
- Matching/pricing logic must remain deterministic and testable.
- Port VBA behavior faithfully unless a change is explicitly decided.
- Prefer zip bundles of updated files for handoff in this project.
