# MVP Scope — current

The current MVP is no longer only a VBA-to-Python port. The project now includes
a deterministic Python core, SQLite persistence, a FastAPI web UI, an admin UI,
and a working RNMC ZIP import flow.

## MVP goal

Provide a local web service that lets an estimator:

1. maintain a historical RNMC catalog in SQLite;
2. upload an estimate/BOQ Excel file;
3. run deterministic RNMC matching and optionally add one TKP candidate;
4. download a WA Excel result;
5. review risky price ranges;
6. approve ranges into `gesn_exceptions`;
7. import new RNMC files from ZIP archives;
8. inspect import history, rejected rows, and data-quality issues.

## Non-negotiable boundaries

- Matching/pricing must stay deterministic.
- No LLM or semantic matching inside the RNMC matching/pricing path. Optional
  TKP lexical scoring stays isolated and deterministic.
- Region is metadata for display/risk review, not a matching filter.
- Human approvals are stored and auditable.
- New domain rules require tests.

## Done in the MVP

### Core and Excel result

- Normalization, exact `(unit, code)` search key, demolition detection.
- Name exclusion rules and task-color metadata.
- Catalog grouping and 4% dedup.
- Matching, pricing, risk calculation, approval-range override.
- Flexible estimate layout detection and multi-sheet choice.
- WA Excel writer with analogs, average formula, `/KR`, section code, and risk
  colors, with preservation of EMF/WMF drawings, printer settings, original
  pagination metadata, and automatic formula recalculation on open.
- Optional TKP toggle with one best candidate per row, three-column Excel
  output, and inclusion of the TKP price in the average formula.

### Database

- `catalog_items` and `catalog_sources`.
- `imported_files` and `import_row_log`.
- `name_exclusion_rules` and `task_color_entries`.
- `price_risk_log` and `gesn_exceptions`.
- `tkp_sources` and `tkp_items`.

### Web/admin

- Main upload/run/download flow.
- Admin dashboards for sources, imports, risks, approvals, rules, exceptions, and
  settings.
- Approve-risk workflow.
- Edit workflows for task colors and name exclusions.
- TKP catalog import and full-grid browsing in `/admin/tkp`, including
  server-side filters/sorting/pagination and browser-persisted column layout.
- TKP storage keeps the selected 27 position, winner, procedure, and audit
  fields, including the original quantity text and both unit/line prices.
  WOR-only catalog files are accepted; all retained fields are available as
  configurable columns in the catalog grid.

### RNMC import

- Legacy `File_Log.xlsx` migration into `imported_files`.
- ZIP dry-run, import-log recording, row preview, and real catalog import.
- Filename-only dedup, with duplicate filenames marked as `duplicate_name`.
- Region from parent ZIP folder or manual override.
- Rejected-row logging and per-file import detail pages.
- Retry unlock for `failed` / `no_data` records through the next ZIP upload.

## Explicitly out of scope for the current MVP

- Authentication and user roles.
- Cloud deployment.
- Semantic/embedding matching.
- Automatic region-based price adjustment.
- Automatic extraction of LSR quarter and planned dates from every RNMC format
  (next milestone).
- `.xls` detailed parsing (deferred unless real incoming files require it).
- True one-click retry without re-uploading ZIP files.
- Watched-folder/scheduled import automation.

## Definition of done for the current MVP

For representative real files, the service should be able to:

1. import or use the RNMC catalog;
2. process an estimate and produce the expected analog assignments and price
   flags;
3. persist open risks and approvals;
4. import new RNMC ZIP files without duplicating already processed filenames;
5. show import problems clearly enough for manual correction.
