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
- Review catalog rows and submit logged corrections for original/ZLVL unit
  prices and other fields. A correction changes the database only after senior
  approval in `/admin/corrections`.
- Import legacy `File_Log.xlsx` records into `imported_files`.
- Upload RNMC ZIP archives, run dry-run checks, use a tabbed 30-row workbook
  preview, exclude selected rows or whole tasks before confirmation, import the
  remaining valid rows into `catalog_items`, detect workbook metadata, store
  original and ZLVL unit prices, and inspect per-file import details.
- Select a folder of original KL 2.0 workbooks in the TKP admin page, detect
  each winner, import the winner's priced WOR rows into SQLite, retain the
  procedure/audit fields, and optionally add one best TKP candidate to the
  estimate result without changing RNMC matching. Importing a prebuilt
  CatalogBuilder workbook remains available as a fallback.
- Approve price risks into `gesn_exceptions`.
- Edit task color entries and name exclusion rules from the admin UI.
- Resize columns in the analog catalog; browser-local widths can be reset from
  the catalog toolbar.
- Compare the live TKP result with strict deterministic filtering and optional
  local Qwen3/BGE-M3 embedding models in an isolated shadow panel.

Matching/pricing remains deterministic. No LLM or semantic matching is used
inside the matching/pricing path. Shadow TKP comparisons are read-only and
never alter an estimate result.

## How to run

```bash
python -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python -m pytest -q
python -m app.web
```

Optional local semantic experiments use `requirements-semantic.txt`. Model
weights are not included in the application archive and are never downloaded
automatically.

Main web routes:

- `/` — upload estimate/catalog and run matching.
- `/admin` — admin dashboard.
- `/admin/imports` — RNMC import dashboard and control center.
- `/admin/catalog` — searchable editable catalog table.
- `/admin/corrections` — correction journal, approval/rejection, and audit
  history.
- `/admin/sources` — catalog sources.
- `/admin/risks` — price risk log.
- `/admin/approvals` — approve open price risks.
- `/admin/gesn-exceptions` — approved GESN ranges.
- `/admin/task-colors` — blue-task metadata.
- `/admin/name-exclusions` — exclusion rules.
- `/admin/tkp` — direct original-KL folder import, fallback aggregate-catalog
  import, and a full grid with filters, sorting, pagination, configurable
  columns, and resizable widths.
- `/admin/settings` — database/settings diagnostics.

## Project docs and AI handoff (2026-10-28)

**Every new AI developer should start from [`AGENTS.md`](AGENTS.md) and
[`docs/AI_HANDOFF.md`](docs/AI_HANDOFF.md).** This handoff separates the source
ZIP state from the live Windows/SQLite deployment verified in October 2026.
Do not deploy or overwrite a production DB based on this source ZIP alone.

- `AGENTS.md` — mandatory top-level AI agent entry point.
- `docs/AI_HANDOFF.md` — current architecture, project decisions, production
  status and source-of-truth boundaries.
- `docs/AGENTS.md` — AI agent and coding rules.
- `docs/SERVER_RUNBOOK.md` — safe Windows deployment and SQLite rollback.
- `docs/DATA_IMPORT_GUARDRAILS.md` — RNMC/TKP/Excel provenance, strict QA and
  regression matrix.
- `docs/DOMAIN_RULES.md` — business rules extracted from VBA and accepted
  product decisions.
- `docs/MVP.md` — current MVP scope.
- `docs/ROADMAP.md` — completed milestones and next work.
- `docs/RNMC_IMPORT.md` — RNMC ZIP/File_Log import workflow.
- `docs/OPEN_ITEMS.md` — deliberately deferred decisions and follow-ups.
- `docs/ARCHITECTURE.md` — module and SQLite data-flow map.
- `docs/DECISIONS.md` — explicit and pending architectural/domain decisions.
- `docs/RELEASE_CHECKLIST.md` — development, review, deploy and rollback gates.
- `docs/README.md` — navigation by task.
- `docs/GITHUB_SYNC.md` — review-safe procedure for publishing the documentation to `feature/admin-ui`.

**Documentation snapshot, not a published GitHub change:** these files were
prepared from the user-provided project ZIP; they must be reconciled with
any newer `feature/admin-ui` commits before a normal Git commit/PR.
The SQLite figures in the handoff describe the **2026-10-27** server verification,
not a live connection to production.

### Single Excel RNMC upload

The admin imports page has a second compact upload card for a single `.xlsx` or `.xlsm` file.

- `РНМЦ_КА_ЖО_ZLVL_V3.xlsx` is detected by its exact base filename and imported with the full legacy ZLVL catalog mapping.
- Any other supported Excel filename is treated as one new RNMC workbook and uses the standard preview and confirmed import flow.
- `imported_files` is updated automatically after confirmation.
