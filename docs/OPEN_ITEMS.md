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

2. **Formula cells are read as formula text, not computed values. —
   RESOLVED (2026-07).** Confirmed on real files: price cells are formulas
   like `=4937.682*(Дефлятор!$E$2)`. Reading now loads values with
   `data_only=True` (`read_catalog_rows`, `read_estimate_rows_with_positions`,
   `load_estimate`, and coefficient resolution), so formula cells resolve to
   the numbers Excel cached on save. The Excel *writer* keeps
   `data_only=False` to preserve formulas in the output copy.
   Caveat: `data_only=True` relies on the file having been saved by Excel with
   cached values; a workbook produced by a script without cached values would
   return `None` for formula cells.

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

## Flexible layout resolution — deferred rules

Context: `core/layout.py` (see DOMAIN_RULES.md §10) covers column detection
for `work_name` / `unit` / `code` / `base_price` (R6-R9), a minimal
header-row locator, a hard-stop when required data is not found (R20), a
resolution report (R19), the regional coefficient by label (R16),
average-column placement next to the base price (R12), sheet
ranking/selection (R1, `rank_sheets`/`select_sheets`), and a
blank-row-tolerant data range (R5 partial, `data_row_numbers`). The
following robustness rules were agreed with the product owner but
intentionally deferred, to keep each step small:

- **R2/R3/R4** — richer header detection: full label-set matching,
  sub-header/units-row skipping, merged-cell header reading. Only a
  minimal "row with most detected fields" locator exists today.
- **R4 (partial done)** — the column-enumeration row under the header (bare
  integers `1 2 3 ...`) is now skipped in the detected reader (a code cell
  that is a bare integer is treated as a header artefact). Full sub-header /
  units-row handling is still open.
- **R5 (remainder)** — total/summary-row skipping ("итого"/"всего", via a
  config dictionary). The blank-row tolerance part of R5 is done. Note: on
  large real files the detected read can run long (e.g. ~1000+ rows) if
  lower analog-detail sub-tables are separated by fewer than `max_blank_run`
  blank rows; total/section-boundary handling here needs a decision.
- **R1 (web UI)** — DONE. `app/web/` (FastAPI + uvicorn) renders the
  multi-sheet choice as buttons and forces the selection by title; the
  service layer (`app/services/read_estimate.py`) raises `MultipleSheetsError`
  with candidates and accepts `selected_sheet_title`.
- **R10** — quantity column detection (for future qty-aware output).
- **R11 (detection done, write pending)** — the section column
  ('Код раздела') is now detected as a `section` field in `layout.json`, which
  also stops 'Код раздела' from being grabbed as the code column. Writing the
  section code for a detected (non-template) layout is still pending, so the
  detected-layout WA write currently skips the section code (template write
  still fills it). Creating a section column when absent is also still open.
- **R13** — full version deferred. A minimal "R13-lite" exists: for a
  detected (non-template) layout the writer places analogs in the first free
  column after the used range (`max_column + 1`, past the average column).
  The full rule must reserve the leading analog slots for the additional
  sources below (see "Multi-source analogs"), which R13-lite does not.
- **R17/R18** — tolerant numeric parsing (spaces, nbsp, comma decimal,
  currency suffixes) and tolerant code parsing at detection time.
- **Blue task-colour highlight** (DOMAIN_RULES.md §4 step 5 / §6) — the
  task colour list is not yet threaded into the matching run result, so
  the Excel writer cannot apply the blue tint. Wire the task colour list
  into the run result before implementing blue cell colouring; do not
  reintroduce any task-level blocking (§7).

Wire these in as separate, tested steps; each should update DOMAIN_RULES.md
§10.3 and remove its entry here once landed.

## Multi-source analogs — future phase (registered, not yet built)

Product direction (decided 2026-07): beyond the historical RNMC price
catalog, analogs will be pulled from two additional sources and shown in the
leading analog columns. See `docs/ROADMAP.md` ("Multi-source analogs") for
the staged plan and the reserved column layout. These are explicitly
sequenced AFTER the deterministic exact-match core is proven on real data
(AGENTS.md rule 8: no semantic/AI matching before exact matching is solid),
and they must not weaken or override the base exact-match conditions.

Open decisions to make before building these (do not assume):

1. **Does a TKP price feed the recommended-price average formula**, or is it
   display-only alongside the exact-match analogs? (The average formula and
   its `MAX(base, ...)` guarantee must stay well-defined.)
2. **Similarity metric and threshold** for the TKP semantic match — which
   model/embedding, what minimum match % to show a row, and how the % is
   computed and stored deterministically for reproducibility.
3. **Internet-agent price trust**: how to record provenance (source URL),
   how stale a price may be, and whether such prices ever affect pricing or
   are strictly informational with a link.
4. **Reserved analog column ordering** and whether it is configurable (see
   ROADMAP) — and how it interacts with R13's "first free column" placement.
5. **Determinism boundary**: semantic/AI results are non-deterministic by
   nature; keep them out of the core deterministic matching/pricing path and
   isolate them behind a clearly separate, cache-backed layer so a run can be
   reproduced from stored results.

## Process note

This file exists because deferring a decision should not mean forgetting
it. When working through this list: resolve each item as an explicit
decision (with reasoning, not just "left as-is"), update the relevant
code/test if behavior changes, and remove the item from this file once
addressed — don't let it silently go stale.
