# Open Items Backlog

> Running list of things flagged during development that need a deliberate
> decision or follow-up. The deterministic core, SQLite storage, web UI, admin
> UI, and base RNMC ZIP import are implemented; see `docs/ROADMAP.md` and
> `docs/RNMC_IMPORT.md` for current state.

## Catalog editor — next decisions

1. **Persisted calculated/custom columns.** The admin catalog editor can edit
   existing text/numeric columns and perform grouped arithmetic on selected rows.
   Multi-row operations are protected by a confirmation prompt. A future block
   should define how user-created calculated columns are stored, typed, named,
   exported, and consumed by matching/pricing before adding them to the database.

## RNMC import — next decisions and improvements

2. **Review the remaining no-data RNMC files.** The current real-file pass imports
   the normal files and intentionally leaves average-only price templates as
   `no_data`, because `Цена средняя` and `Итого стоимость средняя` are not
   accepted source values. Decide whether those files should be repaired at the
   workbook level or kept out of the catalog.

3. **Support `.xls` detailed parsing if needed.** ZIP dry-run recognizes `.xls`
   as an Excel file, but row preview/import currently supports `.xlsx` and
   `.xlsm`. Add an `.xls` reader only if real incoming files still require it.

4. **True one-click retry requires storing source files.** Current retry unlock
   changes `failed` / `no_data` records to `pending`; the user uploads the ZIP
   again. To retry without re-upload, the service must store original workbooks
   in a durable file store and define retention/cleanup rules.

5. **Duplicate-name review workflow.** Duplicate filenames are correctly marked
   as `duplicate_name`, because filenames must be globally unique. A future UI
   should make it easy to review which folders/regions produced the conflict.

6. **Rejected-row export.** Rejected rows are stored in `import_row_log` and shown
   on the detail page. A download/export path would make manual cleanup easier.

## TKP folder import — follow-ups

7. **Legacy `.xls` / `.xlsb` source workbooks.** Direct folder import currently
   supports `.xlsx` and `.xlsm`, which cover the supplied real KL sample set.
   Add an isolated conversion/reader path only if incoming folders still
   contain required legacy files.

8. **Benchmark and reranker gate for semantic TKP.** The current admin-only
   shadow panel supports strict filters plus local Qwen3/BGE-M3 embedding
   adapters. Build a reviewed real-data benchmark, measure top-3/top-10 and
   no-match precision, then compare the Qwen3/BGE rerankers before allowing any
   semantic result into the live estimate path.

## From excel_io.py

9. **Silent empty result when estimate header row is not found.**
   `read_estimate_rows` returns `[]` with no error/warning if the header row
   detection fails. Higher layers should surface this as a clear message.

10. **Formula cells are read as formula text, not computed values. — RESOLVED
   (2026-07).** Reading now uses `data_only=True` where needed. Caveat:
   `data_only=True` depends on Excel-cached values being present.

## From catalog.py (lower priority)

11. `_parse_iso_date` only accepts strict ISO format (`YYYY-MM-DD`) for string
   dates. Revisit if source files provide dates as non-ISO plain text.

12. Dedup within a single task_id in `BuildCatalog` is O(n^2). Fine for current
   catalog sizes; revisit only if groups grow large.

## Flexible layout resolution — deferred rules

Context: `core/layout.py` covers column detection for `work_name`, `unit`,
`code`, `base_price`, a minimal header-row locator, a hard-stop when required
data is not found, a resolution report, regional coefficient detection, average
column placement, sheet ranking/selection, and blank-row-tolerant data ranges.
Deferred robustness rules:

- **R2/R3/R4** — richer header detection: full label-set matching,
  sub-header/units-row skipping, merged-cell header reading.
- **R4 partial** — column-enumeration rows are partly skipped, but full
  sub-header / units-row handling is still open.
- **R5 remainder** — total/summary-row skipping (`итого` / `всего`) needs a
  config-driven decision.
- **R11** — detected-layout section writing is still pending when the template
  layout is not used.
- **R13 full** — current R13-lite places analogs in the first free column;
  future multi-source analogs require reserved leading analog slots.
- **R17/R18** — tolerant numeric parsing and tolerant code parsing at detection
  time.
- **Blue task-colour highlight** — task colors exist in admin/storage, but the
  list still needs to be threaded into the matching run result before the Excel
  writer can apply blue tint.

## Multi-source analogs — future phase

Future product direction: beyond the historical RNMC catalog, analogs may later
come from additional sources and appear in leading analog columns. This remains
deferred until deterministic exact matching and RNMC import are proven on real
files.

Open decisions before building it:

1. Whether non-RNMC prices affect the recommended-price average formula or are
   display-only.
2. Similarity metric, threshold, and reproducibility model.
3. Internet/source-price provenance, freshness, and trust rules.
4. Reserved analog column ordering and configurability.
5. Determinism boundary: semantic/AI results must stay outside core
   matching/pricing and be stored so runs can be reproduced.

## Process note

Deferring a decision should not mean forgetting it. When an item is resolved,
update the relevant docs and tests and remove or mark the item here.
