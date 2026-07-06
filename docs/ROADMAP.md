# Estimate AI Roadmap

## Goal

Build a web service for matching estimate/BOQ rows against a historical
RNMC catalog, producing analogs, recommended prices, risk flags, approval
history, and downloadable Excel results.

## Product Direction

The target product is a web service:

- Upload or import RNMC catalog files.
- Upload an estimate file.
- Run deterministic analog matching.
- Review matched analogs and price-spread risks.
- Approve or widen GESN exception ranges.
- Download a completed Excel result.
- Keep import history and approval history in a database.

## Current State (2026-07)

Core VBA logic has been ported into a local Python library with a working
end-to-end pipeline, SQLite persistence, and a minimal web UI.

### Done — core library

- `core/normalize.py` — code/unit normalization, demolition detection, search key.
- `core/exclusions.py` — name exclusion rules and task color metadata.
- `core/catalog.py` — catalog construction and demolition-aware 4% dedup.
- `core/matching.py` — estimate row to catalog matching with reason codes.
- `core/risk.py` — ratio risk and approved-range override.
- `core/approval.py` — create/widen `GesnException` ranges.
- `core/sections.py` — GESN prefix extraction and section code resolution.
- `core/pricing.py` — average price formula and regional coefficient.
- `core/excel_io.py` — read catalog and estimate Excel rows (`data_only=True`).
- `core/ingest.py` — folder-walk RNMC ingestion into plain Python structures.
- `core/layout.py` — flexible layout resolution (sheet/columns/regional
  coefficient/average placement, R1/R5/R6-R9/R12/R16/R19/R20).
- `core/excel_writer.py` — write run result into a `WA` copy (analogs,
  average formula, `/KR`, section, cell colours). Risk rows are **not**
  written to Excel anymore (see storage below).
- `core/macro_workbook.py` — load `Name_Exclusions` / task colours from xlsm.

### Done — storage (SQLite, schema v2)

- `core/storage/schema.py` — catalog, import log, rules, `gesn_exceptions`,
  `price_risk_log`.
- `core/storage/catalog.py` — import catalog from Excel, list rows.
- `core/storage/rules.py` — name exclusion rules and task colours.
- `core/storage/risk_log.py` — upsert open risks by `exception_key`,
  load/approve → `gesn_exceptions` (ports Module6, DOMAIN_RULES §5.2).
- `app/cli/` — `init-db`, `import-catalog`, `import-rules`, `status`.

### Done — application services

- `app/services/run_matching.py` — one end-to-end matching run over rows.
- `app/services/read_estimate.py` — flexible read (sheet selection,
  template-or-detected columns, clear errors).
- `app/services/write_result.py` — read + coefficient + match + persist
  risks to SQLite + write WA Excel.
- `app/services/catalog_source.py` — catalog from upload or database.

### Done — web UI (minimal)

- `app/web/` — FastAPI + uvicorn: upload estimate (+ optional catalog),
  run matching, multi-sheet choice, download WA. Run with `python -m app.web`.
- Catalog can come from SQLite when the DB is populated (`import-catalog`).

### Done — validation

- 186 automated tests (pytest).
- Real-file comparisons against VBA macro output (scripts in `scripts/`).
- Example: estimate 6458203 — WAW=WAM for matching; Price_Check_Log channel
  differs from macro by design (risks now in SQLite, not WA sheet).

### Important docs

- `docs/AGENTS.md`
- `docs/DOMAIN_RULES.md`
- `docs/MVP.md`
- `docs/OPEN_ITEMS.md`

## Completed Milestones

- [x] Port core VBA matching/pricing/risk logic to tested Python modules.
- [x] End-to-end local run: catalog + estimate → structured result.
- [x] Excel writer: analogs, average formula, `/KR`, section, colours.
- [x] Flexible layout read (template + detected headers).
- [x] SQLite catalog storage and CLI import.
- [x] Minimal web UI (upload → run → download WA).
- [x] Risk log in SQLite (`price_risk_log`), `gesn_exceptions` from DB only;
  red cell highlighting in estimate preserved; `Price_Check_Log` sheet removed
  from WA output.

## Next Milestone — Admin UI (`feature/admin-ui`)

Backend for risk review and approval is in place (`core/storage/risk_log.py`,
`core/approval.py`). Next: admin screens and API wiring.

1. **Risk review UI** — list open rows from `price_risk_log` (filter by status,
   reason, code, unit).
2. **Approve workflow** — call `approve_risk()` → widen `gesn_exceptions`,
   mark log row `approved`; verify re-run no longer flags the same spread.
3. **Import dashboard** — show `imported_files` / `import_row_log` history;
   trigger `import-catalog` / force re-import (DOMAIN_RULES §9.6).
4. **Rules management** — view/edit `name_exclusion_rules` and task colours
   (or re-import from macro workbook).
5. **REST endpoints** (if UI needs them):
   - `GET /risks`, `POST /risks/{key}/approve`
   - `GET /catalog/imports`, `POST /catalog/import`
   - optional: `POST /runs` refactor of current `/run` flow

No new dependencies unless explicitly approved.

## Web Service Path — remaining

After admin UI:

1. Refactor upload/run into proper REST resources (`POST /runs`, `GET /runs/{id}`).
2. Auth / user roles (explicitly deferred).
3. Watched-folder auto-import.
4. Better Excel output formatting (OPEN_ITEMS: blue task-colour tint, R13 full).

## Multi-source analogs (future phase, registered 2026-07)

Beyond the historical RNMC price catalog, analogs will be drawn from two
additional sources and shown in the **leading analog columns**, ahead of the
exact-match catalog analogs. Sequenced **after** the deterministic exact-match
core is proven on real data. See `docs/OPEN_ITEMS.md` ("Multi-source analogs")
for open decisions. Do not implement before those decisions are made.

## Later Enhancements

- User roles/authentication.
- Region-aware analytics if explicitly approved.
- Fuzzy/semantic matching only after exact rule-based matching is proven.
- Full R13 analog column placement (reserved slots for multi-source analogs).

## Working Rules

- Speak with the user in Russian.
- Keep code, filenames, identifiers, and code comments English/ASCII only.
- Do not invent prices.
- Matching/pricing logic must remain deterministic and testable.
- Port VBA behavior faithfully unless a change is explicitly decided.
- Keep one task small and commit after each completed module or milestone.
