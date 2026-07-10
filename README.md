# Estimate AI

Construction pricing assistant for estimators. The service processes BOQ / ВОР
Excel files, matches rows against a historical RNMC catalog, calculates a price
corridor, flags risky price spreads, and produces an Excel result for human
review.

The system does not replace the estimator. It prepares a deterministic draft and
keeps risky decisions visible for review and approval.

## Current stage

The project now has a tested Python core, SQLite persistence, a FastAPI web UI,
and an admin UI for catalog/import/risk control.

Implemented product flows:

- Upload an estimate file and produce a WA Excel result.
- Use a catalog from SQLite or from an uploaded Excel file.
- Review and edit catalog rows, including original/ZLVL unit prices, catalog sources, import history, risks, approvals, rules, and settings
  in `/admin`.
- Import legacy `File_Log.xlsx` records into `imported_files`.
- Upload RNMC ZIP archives, run dry-run checks, use a tabbed 30-row workbook
  preview, import valid rows into `catalog_items`, detect workbook metadata, store
  original and ZLVL unit prices, and inspect per-file import details.
- Approve price risks into `gesn_exceptions`.
- Edit task color entries and name exclusion rules from the admin UI.

Matching/pricing remains deterministic. No LLM or semantic matching is used
inside the matching/pricing path.

## How to run

```bash
python -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python -m pytest -q
python -m app.web
```

Main web routes:

- `/` — upload estimate/catalog and run matching.
- `/admin` — admin dashboard.
- `/admin/imports` — RNMC import dashboard and control center.
- `/admin/catalog` — searchable editable catalog table.
- `/admin/sources` — catalog sources.
- `/admin/risks` — price risk log.
- `/admin/approvals` — approve open price risks.
- `/admin/gesn-exceptions` — approved GESN ranges.
- `/admin/task-colors` — blue-task metadata.
- `/admin/name-exclusions` — exclusion rules.
- `/admin/settings` — database/settings diagnostics.

## Project docs

- `docs/AGENTS.md` — AI agent and coding rules.
- `docs/DOMAIN_RULES.md` — business rules extracted from VBA and accepted
  product decisions.
- `docs/MVP.md` — current MVP scope.
- `docs/ROADMAP.md` — completed milestones and next work.
- `docs/RNMC_IMPORT.md` — RNMC ZIP/File_Log import workflow.
- `docs/OPEN_ITEMS.md` — deliberately deferred decisions and follow-ups.

### Single Excel RNMC upload

The admin imports page has a second compact upload card for a single `.xlsx` or `.xlsm` file.

- `РНМЦ_КА_ЖО_ZLVL_V3.xlsx` is detected by its exact base filename and imported with the full legacy ZLVL catalog mapping.
- Any other supported Excel filename is treated as one new RNMC workbook and uses the standard preview and confirmed import flow.
- `imported_files` is updated automatically after confirmation.
