# Open Items Backlog

> Running list of things flagged during Phase 0 development that need a
> deliberate decision, but were intentionally deferred rather than decided
> on the spot. Review this file in full once all core modules
> (normalize → exclusions → catalog → matching → risk → approval →
> sections → pricing → excel_io) are done, before moving to ingest.py /
> database / web layer.

## From excel_io.py

1. **Silent empty result when estimate header row is not found.**
   `read_estimate_rows` returns `[]` with no error/warning if the header
   row detection (ГЭСН/ФЕР/Пер marker search in rows 1-50) fails. This is
   fine at the `core/` layer (no I/O/UX concerns there by design), but a
   higher layer (CLI, web API) needs to surface this as a clear message to
   the user — "could not find estimate header row" is very different from
   "file processed, zero matching rows found," and right now both look
   identical from the caller's side.

2. **Formula cells are read as formula text, not computed values.**
   Workbooks are loaded with `data_only=False`. If real catalog/estimate
   files turn out to have formula-driven price/code/date cells (rather
   than static values), those cells will currently be read as formula
   strings (e.g. `"=B4*1.2"`), which will fail numeric/date parsing
   downstream rather than reading the actual computed number. Needs
   `data_only=True` (requires the file to have been saved by Excel with
   cached values) or another fallback strategy — not yet handled, only
   relevant if/when real files turn out to use formulas in these columns.

## From catalog.py (still open, lower priority — noted earlier)

3. `_parse_iso_date` only accepts strict ISO format (YYYY-MM-DD) for
   string dates. Now that excel_io.py is in place and dates mostly arrive
   as `datetime` objects from openpyxl, this is less likely to matter in
   practice — but if any source ever provides dates as plain strings in a
   non-ISO format (e.g. typed as text in a cell, not a real Excel date),
   they will silently resolve to serial 0. Revisit if this turns out to
   happen with real files.

4. Dedup within a single task_id in `BuildCatalog` is O(n^2). Fine for
   realistic catalog sizes (tens of entries per task); flag only if it
   ever needs to scale to hundreds+ entries per task.

## From DOMAIN_RULES.md §9.6 — catalog import / ingestion: RESOLVED

All five original open questions now have explicit product decisions —
see DOMAIN_RULES.md §9.6. Summary: normal import runs skip already-logged
files (filename+region-folder is the identity key, same as VBA); a
separate explicit "force re-import" action does a clean delete+reinsert
for one file; validation happens at ingestion with a per-row import log;
region-from-folder-name is confirmed permanent; one task per file is
confirmed correct (no multi-sheet handling needed); failed imports are
never auto-retried, recovery is always via the same explicit single-file
re-import action, which must be fast/low-friction in the future UI.

`core/ingest.py` can be built directly against DOMAIN_RULES.md §9.6
without further clarification needed.

Implementation note: `force_reimport=True` results from `core/ingest.py`
do not delete or replace anything by themselves. The module only parses
files and reports `IngestFileResult`; the future database/persistence
layer must perform the explicit "delete old rows for this file, then
insert new rows" operation. This is intentional storage-agnostic design,
not a missing feature in `core/ingest.py`.

Also, both fully failed files (`failed=True`) and successfully processed
files are added to the imported key set, so a fresh normal `ingest_folder`
call can only report `skipped=True` for either case. Distinguishing
"skipped because already successfully imported" from "skipped because it
failed before" must come from a future import log/history table, not from
this module's return value alone.

## Process note

This file exists because deferring a decision should not mean forgetting
it. When working through this list: resolve each item as an explicit
decision (with reasoning, not just "left as-is"), update the relevant
code/test if behavior changes, and remove the item from this file once
addressed — don't let it silently go stale.
