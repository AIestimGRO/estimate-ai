# Estimate AI — architecture and data ownership

**Source audit:** 2026-10-28, against the user-uploaded snapshot. This is a
component map, not a deployment diagram confirmed against the current host.

## Entry points

- Local development: `python -m app.web` or `run_web.ps1` on Windows;
  admin routes in `app/web/app.py` and templates in `app/web/templates/`.
- CLI: `app/cli/`; batch/pipeline helpers under `scripts/`.
- Production (observed 2026-10-27): Scheduled Task `EstimateAI` runs a
  separate server-side wrapper `deploy/windows/server_task.ps1` with Uvicorn
  TLS on port **7777**. These server deployment files were **not in the
  supplied snapshot**; they must be verified on the live server.

## Data and calculation paths

```text
Estimate Excel
  -> app/services/read_estimate.py / core/layout.py
  -> core/normalize.py / core/catalog.py / core/matching.py
  -> core/pricing.py / core/risk.py / core/sections.py
  -> optional core/tkp_matching.py (deterministic live lexical choice)
  -> app/services/write_result.py / core/excel_writer.py
  -> WA Excel (preserve unrelated workbook parts)

RNMC ZIP + legacy File_Log.xlsx
  -> app/services/rnmc_zip.py / rnmc_excel.py
  -> core/storage/catalog.py, imported_files, import_row_log, catalog_items

Original KL 2.0
  -> core/tkp_folder_ingest.py / core/tkp_ingest.py
  -> core/storage/tkp.py -> tkp_sources, tkp_items
  -> optional TKP live candidate during estimate run

TKP diagnostics (read-only)
  -> app/services/tkp_shadow.py / core/tkp_candidate_rules.py
  -> local embeddings only when explicitly configured
```

## Persistence

Schema lives in `core/storage/schema.py` (`SCHEMA_VERSION = 12` in this
snapshot). Database connection/config in `core/storage/connection.py`;
standard SQLite file under `data/estimate_ai.db` in development (local
fixture/snapshot, **not** proof of production content). Main groups:

| Group | Tables |
| --- | --- |
| RNMC catalog | `catalog_sources`, `catalog_items` |
| RNMC import audit | `imported_files`, `import_row_log` |
| TKP catalog | `tkp_sources`, `tkp_items` |
| Risk and decisions | `price_risk_log`, `gesn_exceptions`, correction request/change/event tables |
| Rules/config | `name_exclusion_rules`, `task_color_entries`, `task_highlight_reasons` |

Always inspect `PRAGMA table_info`, foreign keys and migration status on
the **actual** database; do not assume all historic snapshots share this
schema. Do not conflate RNMC prices and winner prices.

## Important boundaries

1. RNMC code/unit matching is core deterministic behavior; do not change
   legacy normalization/4% dedup/demolition policy by intuition.
2. TKP live candidate scoring is **not** strict engineering equivalence: the
   unit contributes to score rather than always rejecting mismatches.
3. Original KL may contain many bidders; the persisted TKP winner record
   does **not** prove absence of other bidders or prior FRIZ quotes.
4. Semantic TKP shadow analysis cannot modify a live price selection.
5. No runtime database, client workbook, model weights, server logs or keys
   belong in source deliverables.

See `DATA_IMPORT_GUARDRAILS.md` for column provenance and exceptions;
`AI_HANDOFF.md` for validated project history.
