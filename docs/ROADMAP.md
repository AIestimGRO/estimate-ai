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

The current phase should still stay simple: build a working local pipeline
first, then wrap it with API and UI layers.

## Current State

Core VBA logic has been ported into a local Python library.

Implemented modules:

- `core/normalize.py` — code/unit normalization, demolition detection, search key.
- `core/exclusions.py` — name exclusion rules and task color metadata.
- `core/catalog.py` — catalog construction and demolition-aware 4% dedup.
- `core/matching.py` — estimate row to catalog matching with reason codes.
- `core/risk.py` — ratio risk and approved-range override.
- `core/approval.py` — create/widen `GesnException` ranges.
- `core/sections.py` — GESN prefix extraction and section code resolution.
- `core/pricing.py` — average price formula and regional coefficient.
- `core/excel_io.py` — read catalog and estimate Excel rows.
- `core/ingest.py` — folder-walk RNMC ingestion into plain Python structures.

Important docs:

- `docs/AGENTS.md`
- `docs/DOMAIN_RULES.md`
- `docs/MVP.md`
- `docs/OPEN_ITEMS.md`

## Next Milestone

Build the first end-to-end local matching run:

```text
catalog.xlsx + estimate.xlsx
-> read rows
-> build catalog
-> match estimate rows
-> calculate risk, section, price
-> produce structured result
```

At first, the result can be JSON or an in-memory structure. After that,
write the result back to Excel.

## Immediate Tasks

1. Create an application service for a single matching run.
   Suggested module: `app/services/run_matching.py`.

2. Define result dataclasses for:
   - estimate row result;
   - analog columns;
   - risk result;
   - section code;
   - recommended/average price;
   - row status/reason.

3. Add a simple local CLI command.
   Example target:

   ```powershell
   python -m estimate_ai run --catalog catalog.xlsx --estimate smeta.xlsx --out result.json
   ```

4. Test the pipeline on small fixtures first.

5. Test the pipeline on real anonymized files and compare against VBA output.

6. Add Excel writer:
   - copy/source estimate workbook;
   - write analog columns;
   - write average price;
   - write section code;
   - append `/KR`;
   - preserve useful formatting where practical;
   - add logs for risk checks.

## Web Service Path

After the local pipeline works:

1. Add a small backend API.
   Preferred starting point: FastAPI.

2. Add local database storage.
   Start with SQLite; keep storage interfaces swappable.

3. Add endpoints:
   - `POST /runs`
   - `GET /runs/{id}`
   - `GET /runs/{id}/download`
   - `POST /catalog/import`
   - `GET /catalog/imports`
   - `POST /exceptions/approve`

4. Add a simple web UI:
   - upload catalog/estimate;
   - run matching;
   - show status;
   - download result;
   - review risk rows.

5. Add import dashboard and approval UI.

## Later Enhancements

- Persistent RNMC catalog database.
- Watched/import folder workflow.
- Import history with failed/successful status.
- GESN exception management screen.
- Better Excel output formatting.
- User roles/authentication.
- Region-aware analytics if explicitly approved.
- Fuzzy/semantic matching only after exact rule-based matching is proven.

## Working Rules

- Speak with the user in Russian.
- Keep code, filenames, identifiers, and code comments English/ASCII only.
- Do not invent prices.
- Matching/pricing logic must remain deterministic and testable.
- Port VBA behavior faithfully unless a change is explicitly decided.
- Keep one task small and commit after each completed module or milestone.
