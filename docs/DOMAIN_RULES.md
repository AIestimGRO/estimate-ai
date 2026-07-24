# Domain Rules

> Status: extracted from the working VBA macro set (Module1–Module7),
> reviewed 2026. This is not a design proposal — it documents what the
> production macro actually does today. Each rule below cites the source
> module/function it came from. When porting to Python, behavior must match
> this file unless a change is explicitly decided and noted.

## 1. Inputs

Two Excel files per run:

- **Catalog (РНМЦ / historical prices)** — "F1" file, sheet name containing
  "Кат" (falls back to first sheet). Columns are configurable (see
  `Settings` below); defaults: task number=B, price w/o VAT=G, code=N,
  region=P, work name=C, unit=D, added-date=Q.
- **Target estimate (смета / ВОР)** — sheet name containing "ОС" (falls back
  to first sheet). Defaults: code column=N, average/total price=G, /KR
  column=N, base price=F, section code=O, first analog column=P, work
  name=C, unit=D.

All column positions are user-configurable via a settings block on an
"Instrument" sheet (`Module1`), not hardcoded — read once per run via
`LoadSettings`. A Python port should keep this as external configuration
(e.g. a config file or function arguments), not literals in matching code.

## 2. Normalization

### 2.1 Code normalization (`NormCode`, Module3)

Applied to GESN/FER/TER/"Перечень" codes before any comparison:

1. Replace CR/LF/TAB/non-breaking-space with a regular space.
2. Uppercase, trim.
3. Collapse multiple spaces to one.
4. Remove spaces immediately before/after `/`.
5. Strip a trailing `/КР` suffix (Cyrillic "КР"), if present.

Note: `/KR` stripping happens at normalization time, so the catalog/exact
matching key does **not** distinguish KR vs non-KR codes — this is
different from what was assumed in earlier planning. KR is a separate
output concern (see §6), not a matching-exclusion rule.

### 2.2 Unit normalization (`NormUnit`, Module3)

1. Lowercase, trim.
2. Normalize ё→е, non-breaking space → space, superscript ²→2, ³→3.
3. Collapse multiple spaces.
4. Strip all spaces, periods, commas, and `^` (to build a compact
   comparison key — e.g. "100 м" and "100м" become the same key).

There is no separate "multiplier" extraction in `NormUnit` — the full compact
token is kept for display. **Matching** uses `BaseUnit`, which strips a leading
numeric prefix so ``100 m2`` and ``m2`` share the same lookup key. Analog
prices are **not** scaled by that prefix; only the regional coefficient applies
at write time (see §6).

### 2.3 Matching key (`AnalogSearchKey`, Module3)

```
key = BaseUnit(unit) & "||" & NormCode(code)
```

`BaseUnit` applies `NormUnit` then removes a leading run of digits (e.g.
`100м2` → `м2`, `100м` → `м`). Prices from matched catalog rows are written
as-is; the prefix is not applied as a price multiplier.

If either side is empty, no key is built (prevents accidental matches on
missing data). **Region is intentionally not part of the matching key** —
all regions are pooled together at lookup time; region is only recorded
per price entry for display and risk-flagging, never used as a filter or
adjustment coefficient on the matching side.

### 2.4 Demolition detection (`HasDemontazh`, Module3)

1. Lowercase, normalize ё→е and nbsp→space.
2. Replace punctuation (`.,;:()[]{}/\-_+=` and whitespace chars) with
   spaces, collapse doubles.
3. Split into words; a match is any word whose **prefix** equals the root
   "демонт" (covers демонтаж, демонтажные, демонтированный, etc).
4. "монтаж" alone does NOT match — only checks for the "демонт" root, so
   plain installation words are correctly excluded.

## 3. Catalog construction (`BuildCatalog`, Module3)

For each catalog row, in order — a row is skipped (not raise an error) if
any required field is missing/invalid:

1. Task number (`catTaskCol`) must be non-empty.
2. Price (`catPriceCol`) must be numeric and > 0.
3. Code normalizes to non-empty.
4. Unit normalizes to non-empty.
5. Matching key must be non-empty (redundant with 3+4, kept as safety).
6. **Name exclusion check** (`IsNameExcluded`, scope `"CATALOG"`, Module7)
   — if the work name matches any enabled exclusion rule, the row is
   dropped from the catalog entirely (never offered as an analog).
7. Demolition flag computed from work name.
8. Added-date parsed to a date serial (0 if missing/invalid — treated as
   "no date", which matters for the exceptions logic in §5).

Surviving rows are grouped:

```
catalog[matching_key][task_id] = [ (price, region, is_demolition,
                                     catalog_row_number, full_row_copy,
                                     task_id, norm_code, norm_unit,
                                     added_date_serial), ... ]
```

### 3.1 Deduplication (4%)

Within the same `(matching_key, task_id)` group, AND within the same
demolition flag (demolition and non-demolition entries are never merged
into each other even if prices are close — this avoids losing a needed
analog after the demolition filter is applied later), entries whose prices
differ by ≤ 4% (`DEDUP_PCT` constant, Module2) are collapsed — only the
first one encountered is kept.

## 4. Matching a target row to analogs

For each estimate row with a non-empty normalized code, non-empty
normalized unit, and a positive base price (`colF`):

1. Skip the row entirely if its work name matches a name-exclusion rule
   with scope `"SMETA"` or `"BOTH"` (row gets zero analogs, flagged in log
   as `SKIPPED_BY_NAME_EXCLUSION`).
2. Look up `catalog[matching_key]`. If absent → zero analogs.
3. If found, **demolition filter** is applied (can be toggled off via
   `demontazhFilterEnabled` setting, default ON):
   - row is demolition → keep only demolition analogs
   - row is not demolition → keep only non-demolition analogs
   - filter disabled → keep everything under the key, mixed
4. All surviving (task_id, price) entries become output columns — one
   column per `(task_id, price-position-within-task)` pair, columns
   ordered by first appearance of each task_id, prices within a task
   numbered 1, 2, 3...
5. **Task color list** (Module7): tasks marked in `Name_Exclusions` H:K
   columns get their entire analog column tinted blue — informational
   highlight only, does not affect matching or risk logic.

## 5. Risk / price-spread checking

Two parallel mechanisms, mutually exclusive per matching key:

### 5.1 Default: ratio check (`IsProblemPriceGroup`, Module4)

- Needs ≥ 2 priced analogs.
- `ratio = max_price / min_price`.
- Flagged as a problem if `ratio >= priceSpreadLimit` (default threshold
  3.0 in production Instrument settings; the VBA fallback before reading
  Instrument is 2.0, configurable on Instrument sheet row 46).
- If flagged: the min and max entries are logged to `Price_Check_Log` as
  reason `RATIO_EXCEEDED` (deduplicated per run so the same pair is not
  logged twice), and all analog cells for that key are colored red in the
  estimate.

### 5.2 Approved-range override (`GESN_Exceptions` sheet, Module6)

This is a **human-in-the-loop learning mechanism**, not just a flag:

- A separate sheet `GESN_Exceptions` stores, per
  `(unit, code, demolition-flag)` key, an approved `[min, max]` price
  range plus dates.
- If a key has an approved exception, the system stops using the ratio
  check for it entirely. Instead, every analog entry added to the catalog
  **after** the exception's `LastRangeUpdateDate` is checked against the
  approved range; only out-of-range entries (newer than the exception, and
  outside [approvedMin, approvedMax]) are flagged — reason
  `OUT_OF_APPROVED_RANGE`. Entries within range, or older than the
  exception date, are not flagged again.
- A row in `Price_Check_Log` can be approved by an estimator: setting
  `Approve=1` and running `ApproveMarkedPriceExceptions` (or selecting a
  row and running `ApproveCurrentPriceException`) creates or **widens** the
  approved range in `GESN_Exceptions` to cover the new suggested
  min/max, and timestamps the update.
- Net effect: once an estimator has reviewed and approved a price range
  for a given code+unit+demolition combination, the system "remembers"
  that decision and only re-flags genuinely new outliers going forward,
  not the same already-reviewed spread every run.

This is a more sophisticated and more valuable mechanism than a simple
static threshold — it should be a first-class concept in the Python port,
not an afterthought.

## 6. Output side-effects on the estimate sheet

These are presentation/workflow details, not pure matching logic, but they
encode business rules that must be preserved:

- **Section code** (`ResolveSectionCode`, Module3): derived from the GESN
  prefix via a hardcoded prefix→section lookup table. A small set of
  prefixes (`ГЭСН09, ГЭСН27, ГЭСН28, ГЭСН46, ГЭСНР67`) are
  "demolition-priority": if the row IS demolition, section is forced to
  `08` regardless of the table; if NOT demolition, the table's non-`08`
  entry is preferred if one exists, else falls back to `08`.
- **`/КР` suffix** is appended to the code column for any row that ended up
  with at least one analog. It is added only if not already present, after
  stripping CR/LF/tab/nbsp from the base code. **(2026-07 change)** For a
  row with **no** analog, the `/КР` column is no longer left blank: the
  plain ГЭСН code is copied across as-is, without the suffix.
- **Analog column order (2026-07 rule, not a VBA port):** analog columns
  are grouped by the catalog entry's region. Columns whose region matches
  the estimate file's own declared region (read from the "Регион:" label,
  §10.1/R16) are placed first; all remaining regions follow, grouped
  together, in alphabetical (А-Я) order. A task's own columns (its price
  positions) always stay adjacent. Region-name matching between the file's
  free-text "Регион:" value and the catalog's short folder-name region is
  a best-effort heuristic (`_regions_match` in `core/excel_writer.py`) —
  see the comment there for how it handles Russian adjective forms (e.g.
  "Тула" vs "Тульская область") and its curated synonym table; extend that
  table if a real file exposes a pair it gets wrong.
- **Average price formula**: `=MAX(base_price, IFERROR(AVERAGE(base_price,
  analog_range), base_price))` — written as an actual Excel formula, not a
  static value, into the average/total price column. This guarantees the
  output is never below the base price. **(2026-07 change)** If the column
  right after the base price already holds a formula of this shape (i.e.
  the file was already processed by a prior run), it is recognized and
  overwritten in place rather than treated as "occupied" — this avoids
  inserting a duplicate average column next to a stale one. A column whose
  header is explicitly recognized as the average-price field is also reused
  even when some of its body cells contain manual numeric overrides rather
  than formulas. A genuinely unrelated occupied neighbour still triggers the
  original insert-and-shift behavior.
- **Workbook package preservation (2026-07 rule, not a VBA port):** only the
  selected estimate worksheet is rewritten. Unmodified worksheets retain
  their original OOXML, and unsupported drawing/media parts (including
  EMF/WMF), printer settings, and their relationships are restored after the
  cell write. Original row sizing metadata is retained so print pagination
  does not change merely because the file passed through openpyxl. The result
  requests a full automatic Excel recalculation on open because openpyxl does
  not preserve cached results for rewritten formulas.
- **Regional coefficient**: read from a single configurable cell address
  on the estimate sheet (e.g. `F12`); defaults to 1 if blank, non-numeric,
  ≤ 0, or the cell address itself is not configured. The estimate region is
  detected independently from the configured `Region` label, so reading the
  coefficient from a fixed cell must not discard an adjacent region value.
  Applied as a
  multiplier only to the displayed/written analog prices (and to the
  min/max/recommended values logged to `Price_Check_Log`) — catalog prices
  themselves are never modified.
- **Cell coloring**: blue = task marked in the color list; red
  (`problemFill`) = analog is part of a flagged price problem (per §5);
  grey (`dupFill`) = a second-or-later price within the same task
  (visual de-emphasis, not a data quality flag).

### 6.1 Optional TKP output (2026-07 product rule)

- A per-run toggle, off by default, controls whether the TKP winner catalog is
  searched. It does not change RNMC matching, RNMC analog columns, `/KR`, or
  RNMC price-risk checks.
- TKP matching is deterministic lexical scoring over the estimate work name,
  with unit similarity as supporting context. Only candidates with a positive
  winner unit price are eligible.
- Exactly one best candidate is selected per estimate row. If no candidate
  clears the score threshold, the TKP cells remain blank.
- The writer places a three-column TKP block before the RNMC analog block:
  `Аналог из ТКП`, `Наименование из ТКП`, and `Номер задачи ТКП`.
- The TKP winner unit price is written without applying the estimate regional
  coefficient. When present, it is included in the average-price formula;
  blank TKP cells are ignored by Excel `AVERAGE`.
- A TKP-only match does not make an RNMC row "matched", does not add `/КР`, and
  does not enter the RNMC price-spread risk calculation.

### 6.2 TKP retained fields (2026-07 product rule)

- The application does not archive every raw TKP workbook cell. It persists the
  selected position, price, winner, procedure, and audit fields needed to trace
  a candidate back to the detected winner block.
- Required position fields are `ItemName`, `Unit`, `Qty`, `QtySourceText`,
  `RnmcUnitPriceNoVat`, `RnmcLineTotalNoVat`, `WinnerUnitPriceNoVat`, and
  `WinnerLineTotalNoVat`. `QtySourceText` preserves the source spelling while
  `Qty` remains numeric.
- Required winner/audit fields are `WinnerName`, `WinnerINN`, `WinnerUIN`,
  `WinnerGroupIndex`, `WinnerStartCol`, `WinnerStartColLetter`,
  `WinnerUnitHeader`, `WinnerTotalHeader`, `WinnerMethod`, `WinnerBlockName`,
  `WinnerBlockUIN`, `WinnerBlockTotalVat`, and `WinnerBlockReason`.
- Required procedure fields are `TaskNo`, `RequestDate`, `Version`, `Customer`,
  `GeneralContractor`, and `ProcedureName`.
- Empty optional values are valid and are stored as empty/NULL. In particular,
  a missing `GeneralContractor`, `WinnerBlockUIN`, or RNMC price is not an
  import error.
- `KL20_WOR_Catalog` is sufficient for import. When `KL20_FileCatalog` is
  absent, source records are derived from WOR metadata without inventing
  missing source totals.

### 6.3 Direct KL folder import (2026-07 product rule)

- `/admin/tkp` provides the primary `Upload new TKP` action as a browser folder
  picker. All selected `.xlsx` and `.xlsm` files, including files in nested
  folders, are inspected directly; the VBA CatalogBuilder is not required and
  workbook macros are never executed.
- The preferred worksheet names normalize to `KL20` or `KL2`. If the title is
  nonstandard (a real file uses `KL 4`), the worksheet is accepted only when
  its structure contains the participant-name row, `WOR and Price` block, and
  the expected price header pair.
- Participant groups start at column K, use four columns, and repeat every four
  columns. A group must have a participant name or both a valid price-header
  pair and priced WOR data. Notes to the right of the participant blocks must
  never create phantom participants.
- Winner priority matches the CatalogBuilder business rule: final recommended
  winner (`10.1`), preliminary recommended winner (`8.1`), single participant,
  minimum no-VAT offer total, then minimum sum of WOR line totals.
- Only the selected winner's unit and line prices are stored. Section context,
  source row, quantity source text, RNMC comparison values, winner identity,
  and procedure metadata remain auditable.
- Direct uploads use a deterministic content fingerprint plus the parser
  version. Re-uploading identical bytes is skipped; changed bytes under the
  same filename replace that file's old source and item rows.
- Duplicate base filenames inside one selected folder are rejected rather than
  silently overwriting one another. Unsupported files are ignored and
  recognized workbooks with an unresolved winner are stored as `Needs review`.
- Importing the legacy two-sheet/WOR-only CatalogBuilder result remains
  available under the fallback import control.

### 6.4 TKP shadow comparison (2026-07 product rule)

- The live estimate-processing path remains the deterministic lexical matcher
  described in §6.1. Shadow results never change a workbook or database row.
- `/admin/tkp` can compare one work name/unit against the live result, a strict
  filtered lexical list, and optional local embedding models.
- Strict filtering requires a positive winner price, compatible base units,
  correct numeric unit scaling (for example `100 m2` to `m2`), compatible
  explicit multiplicity, and no installation/demolition conflict.
- The strict lexical list has its own minimum relevance threshold so unrelated
  same-unit positions are not presented as useful analogs.
- Optional Qwen3-Embedding-0.6B and BGE-M3 adapters load only from local model
  directories. The application never downloads weights automatically.
- A missing dependency/model is a visible `unavailable` shadow status and does
  not affect live matching.

## 7. Name exclusion rules (Module7)

A configurable rule table on sheet `Name_Exclusions`, columns:
`Enabled | Scope | MatchMode | Pattern | Group | Comment`.

- `Scope`: `SMETA` (estimate rows only), `CATALOG` (catalog rows only), or
  `BOTH`.
- `MatchMode`: `CONTAINS` (substring match) or `ALL_WORDS` (default — every
  `|`-separated token in `Pattern` must appear somewhere in the normalized
  text; order-independent, all tokens required).
- Text is normalized by collapsing whitespace/control characters only (NOT
  the same normalization as code/unit — case-insensitive substring/token
  match via `vbTextCompare`).
- Default shipped rules (disabled or enabled per row) cover: "каждый /
  последующий" wording variants, "см / изменен" and "мм / изменен"
  (thickness/depth change variants), "дополнительный / щит" (extra shield
  variants). Two broader optional rules exist but ship disabled, pending
  manual review.

A second, independent list on the same sheet (columns H:K) is the **task
color list** — marks specific catalog task numbers for blue highlighting
only (§4 step 5). This list does **not** exclude or block anything; an
earlier "stop list" concept was retired (`IsTaskStopped` is now a
compatibility no-op that always returns False) — do not reintroduce
task-level blocking without an explicit decision.

## 8. What is explicitly NOT happening in the live RNMC logic

These are worth stating because they differ from earlier, more
speculative product planning:

- No semantic/fuzzy work-name matching of any kind. Matching is 100%
  exact on `(normalized unit, normalized code)`. Work-name text is used
  only for demolition detection and exclusion-rule matching, never for
  similarity scoring.
- No region-based filtering or adjustment coefficient in matching. Region
  is descriptive metadata attached to a price, shown in output, not used
  to accept/reject/adjust a match.
- No multi-level "match quality" tiers (exact → normalized-unit → fuzzy
  name → manual). There is one tier: exact key match, then demolition
  filter, then risk flagging. If no catalog entry shares the exact
  (unit, code) key, the row simply gets zero analogs — there is currently
  no fallback search.
- No automatic coefficient table for unit multiples (100 m vs m) — these
  are different normalized units today, not related via a multiplier.

These gaps are reasonable candidates for the Python port to *improve on*,
but the existing approval/exception workflow (§5.2) is the differentiating
asset and must not be lost or simplified away in the process.

The isolated TKP shadow comparison in §6.4 is an explicit exception to this
description. It is an admin experiment, not part of live RNMC or TKP output.

## 9. Catalog import module (`Module8_catalog_import.bas` — file named
"Module6.bas" by the author, renamed in this repo to avoid clashing with
the GESN Exceptions module already named Module6)

> Status: working but author-flagged as unfinished. Documented here as-is;
> see "Open questions" at the end of this section for things that need a
> decision before porting, not silent assumptions.

### 9.1 What it does, end to end

Entry point `RNMC_AppendFiles_ToCatalog`:

1. User picks a root folder via a folder dialog.
2. A "logged files" dictionary is built once from the `FileLog` sheet
   (column B = filename, lowercased) — this is the dedup mechanism: a file
   already logged (by name only, not path or content hash) is skipped
   entirely on future runs.
3. The folder is walked recursively. For every `.xlsx`/`.xlsm`/`.xls` file
   not already in the logged-files dictionary, `ProcessOneFile` runs.
4. **Region is derived from the immediate parent folder name** of the file
   (`GetLastFolderName` on the file's folder path) — there is no region
   column read from inside the file itself. This means the expected
   folder layout is `<root>/<RegionName>/file.xlsx`, and region accuracy
   depends entirely on correct foldering, not on file content.
5. Each workbook is opened read-only; every worksheet in it is scanned
   (`FindHeaderRow`) until one is found with a header row containing all
   three of: a name column, a unit column, and a quantity column (see
   §9.2 for exact header matching). The **first worksheet that matches AND
   yields `added > 0` rows wins — remaining worksheets in that file are
   not processed** (`Exit For` once `added > 0`). If a workbook has
   multiple relevant sheets (e.g. several task tables), only the first
   one found gets imported.
6. A single task number is extracted once per workbook (`ExtractTaskNumber`,
   §9.3) and applied to every row appended from that workbook — there is
   no per-row or per-sheet task number, it's one value per file.
7. Rows are appended into the `Каталог` sheet (`SH_CATALOG`) starting at
   row 4, with column 1 = sequential catalog number, column 2 = task
   number, column 15 = source filename, column 16 = region name, and all
   other columns mapped from the source sheet by matching normalized
   header text against the existing `Каталог` header row (row 3) — see
   §9.4.
8. Every processed file (success or failure) gets one row appended to
   `FileLog`: folder path, filename, count of rows added, timestamp, and
   an error description column if the import failed. A file that errors
   out is still marked "logged" (added to the in-memory dict and written
   to FileLog), so it will **not** be retried automatically on the next
   run — a failed import requires manual cleanup of the FileLog row before
   the file will be picked up again.

### 9.2 Header detection (`FindHeaderRow`, `IsNameHeader`, `IsUnitHeader`,
`IsQtyHeader`)

Scans up to row 400 / column 150 of the (first matching) worksheet. A row
qualifies as the header row only if, somewhere in that row, all three
column roles are found simultaneously:

- **Name column**: cell text (normalized — lowercased, nbsp/CR/LF stripped
  to space then ALL spaces stripped) equals "наименование работ" or
  "наименование", OR starts with "наименование".
- **Unit column**: normalized text starts with "ед.изм." OR contains it
  anywhere as a substring.
- **Quantity column**: normalized text contains "кол-во" or "количество"
  anywhere as a substring.

Note `NormalizeText` here strips ALL spaces (not just collapsing them, as
`NormUnit`/`NormCode` do elsewhere) — this header-matching normalization
is intentionally looser/different from the matching-key normalization in
Module3, and is not shared code. A Python port should keep these as two
distinct normalization functions, not unify them silently.

### 9.3 Task number extraction (`ExtractTaskNumber`)

Scans the top-left area (rows 1–50, columns 1–20) of every worksheet in
the workbook, looking for the literal label text "№ задачи 1Ф" or, failing
that, "№ задачи" as a substring anywhere in a cell. When found:

1. Take the text after the label within the same cell, clean it up
   (strip `:`, `#`, collapse whitespace) — if non-empty, that's the task
   number.
2. If nothing follows the label in the same cell, look at up to 3 cells to
   the right in the same row (`NeighborTaskValue`) and use the first
   non-empty one.
3. First worksheet/match found wins; stops scanning entirely once a task
   number is extracted.

If no label is found anywhere, task number is `""` for the whole file —
there is no error raised, rows are still imported with a blank task
number.

### 9.4 Row appending (`AppendTableRows`)

- A "numbering column" (e.g. "№ п/п") is looked for in the source header
  map by substring rules (contains "№", starts with "no", contains "pp",
  "p/p", or "p-p"); if not found, column 1 of the source sheet is used as
  a fallback "numbering" reference (only used to detect blank rows, not
  written anywhere).
- Reading starts at `hdrRow + 1`, skips any fully-blank leading rows, then
  reads rows until **3 consecutive rows are blank across number+name+unit+
  qty** (end-of-table heuristic) — capped by the sheet's `UsedRange` as a
  safety ceiling.
- A row is imported only if **at least one of unit or quantity is
  non-blank** (a row with a name but no unit and no quantity is silently
  skipped, not logged as an error).
- For each imported row, only source columns whose **normalized header
  text exactly matches** a normalized header already present in the
  `Каталог` sheet's row-3 header get copied over (by column-name lookup,
  not column position) — anything in the source file that doesn't have a
  same-named column in the catalog is silently dropped. Columns 1/2/15/16
  (number, task, source file, region) are never overwritten by this
  mapping since they're set explicitly beforehand.
- All values are copied via `.Value2` (raw values, not formulas/formatted
  text) — no normalization, validation, or type-checking is applied to
  price, code, or unit values at import time. Bad data (wrong type, empty
  required fields, malformed codes) flows straight into the catalog;
  normalization/validation only happens later, when `BuildCatalog`
  (Module3, §3 above) reads the catalog for matching.

### 9.5 Region handling — important divergence from §1/§3

Region is **not** read from any column inside the source files at all in
this module — it comes exclusively from the immediate parent folder name
on disk. This is a different mechanism from `catRegionCol` in the main
matching pipeline (Module1/Module3), which reads region from a configured
column inside the `Каталог` sheet itself. The import module writes region
into catalog column 16; whatever the `catRegionCol` setting is in
`Instrument` needs to actually point at column 16 for these imports to be
picked up correctly by matching — this is an implicit dependency between
two modules that isn't enforced anywhere in code.

### 9.6 Open questions (decide before porting, do not assume)

These are things the current VBA does that may or may not be the intended
behavior for the database version — flagging rather than guessing:

1. **Dedup is by filename only**, case-insensitive, not by file content
   hash or modified-date. Re-importing a corrected version of a file with
### 9.6 Decisions for the DB-targeted ingestion module

These were open questions during VBA review. Decisions below were made
explicitly by the product owner and supersede the original VBA behavior
where noted. Do not silently revert to VBA behavior, and do not silently
reinterpret these decisions — they were clarified carefully after an
earlier draft of this section got the dedup behavior wrong (see note in
item 1).

1. **Dedup: default behavior is SKIP, same as VBA. "Clean replace" is an
   explicit, separate, user-triggered action — never automatic.**
   - Normal repeated import runs (e.g. "I added more files to the watched
     folder, run import again") must skip any filename already present in
     the import log, exactly like the VBA macro does. This is the default
     and most common path — re-running an import should never silently
     re-process or duplicate files that were already successfully
     imported.
   - A SEPARATE, explicitly-invoked action — "force re-import this file"
     — must exist for the rare case where a user knowingly corrected a
     file and wants it re-processed. Only this explicit action triggers
     "clean replace": delete all previously-imported catalog rows whose
     `source_file` matches, then re-import. This requires `source_file`
     to be a reliable, queryable attribute on every catalog row in the DB
     schema (not just descriptive metadata), so "all rows from file X"
     can be found and deleted atomically before the replacement insert.
   - **Filename-based identity is per (region folder + filename), not
     filename alone.** Two files with the same filename in two different
     region subfolders (e.g. `region1/rnmc.xlsx` and `region2/rnmc.xlsx`)
     are different files and must be tracked/deduped independently — the
     import log key must include the region folder, not just the
     filename, to avoid one region's file shadowing another's.
   - Content-hash-based dedup was considered and explicitly rejected as
     unnecessary; identity is filename+region-folder, with explicit user
     action required for any re-processing.
2. **Validate at ingestion time, and produce an import log/report.** Unlike
   the VBA macro (raw copy, validation deferred entirely to matching
   time), the DB ingestion path must check each row at import (missing
   price, missing/unparseable code, missing unit, etc.) and record
   per-row validation outcomes — this becomes the basis of an import log
   shown in the future web UI, so a user can see exactly which rows in
   which file were rejected or are incomplete, not just a silent row
   count. Rows that fail validation should still be logged (with the
   specific reason), not silently dropped without a trace — the log is
   the point.
3. **Region from immediate parent folder name is confirmed as the correct,
   permanent convention** — real files are always organized as
   `<root>/<RegionName>/file.xlsx`. This VBA behavior is kept as-is, not
   just as a stopgap. The DB ingestion module can rely on this folder
   structure being consistent and should treat a file found outside any
   region subfolder (e.g. directly in `<root>/`) as a data-quality error
   to surface, not silently default a region.
4. **One task number per file, first-matching-sheet-only is CONFIRMED
   correct and final** — real RNMC files always have exactly one task per
   file. The existing VBA assumption (single task number, first matching
   worksheet) is accurate and should be kept as-is in the Python port, no
   multi-sheet/multi-task handling is needed.
5. **Failed imports: no auto-retry, no transient/permanent distinction —
   matches VBA behavior. A fast manual single-file (re-)import path is a
   required feature, not an afterthought.** There are no known real cases
   of transient failures (e.g. file briefly locked, network drive
   briefly unavailable) that would resolve themselves without touching
   the file — so the system does not need to guess or distinguish error
   types. A file that fails import is logged as failed and is NOT
   retried automatically on subsequent folder-walk runs, same as VBA.
   Recovery is always manual: the user fixes the file (or it was a
   transient issue that's now resolved) and explicitly re-imports just
   that one file — this must be a quick, low-friction action in the
   future UI (e.g. a "retry this file" button next to its failed-import
   log entry), not a multi-step manual process. This reuses the same
   "force re-import" mechanism from item 1.

### 9.7 Implication for the planned database

Conceptually this module is already an ETL pipeline: walk folder → parse
file → normalize headers → map columns → write rows. Porting it to write
to a database table (e.g. `catalog_items`) instead of an Excel sheet range
is a natural fit — the folder-walk, header-detection, and row-mapping
logic carry over largely unchanged; only the write target changes. The
"logged files" dedup dictionary becomes a query against an
`imported_files` table, keyed by (region folder, filename) per §9.6 item
1 — normal runs query-and-skip, exactly like VBA; only the explicit
"force re-import" action drives a delete-then-reinsert. All five
originally-open questions from the VBA review now have explicit product
decisions (§9.6 items 1-5) — `core/ingest.py` can be built directly
against this section without further clarification needed.

## 10. Flexible layout resolution (decided extension, NOT a VBA port)

> Status: this is a deliberate, product-owner-approved extension beyond the
> original macros, decided 2026-07. The VBA reads fixed, user-configured
> column numbers (§1 `LoadSettings`) and a single marker-based header scan.
> The Python service additionally resolves where each logical field lives by
> matching header text against a config-driven synonym dictionary, so it can
> process files that do not exactly match the template. This is new behavior,
> flagged here per AGENTS.md rule 7 — it does not change any matching or
> pricing math, only how columns/rows are located on a worksheet.

Implemented in `core/layout.py`; synonym dictionaries live in
`data/config/layout.json` (Russian header wording stays in config, never in
code — AGENTS.md rule 3). Matching is exact/substring against a fixed
dictionary only — deterministic and testable, no fuzzy/semantic matching
(consistent with §8).

### 10.1 Resolution priority (per field)

1. **Explicit pin** — caller-provided column override wins outright.
2. **Detection** — header text matches a field synonym (equals / startswith /
   contains, per the config `mode`).
3. **Template default** — the historical fixed column (Settings), used for
   *optional* fields only.
4. **Missing** — nothing resolved.

**Required fields (code, unit, base price) never fall back to a template
default.** If a required field is neither pinned nor detected, the layout is
reported as not resolvable (`missing_required` non-empty), which the caller
surfaces as a "key data not found" result rather than silently guessing.

### 10.2 Header-row location

The header row is located by scanning the first N rows (config
`header_scan.max_rows`) and picking the row with the most detected fields;
ties keep the first (smallest-index) row; a minimum number of matched fields
is required (`header_scan.min_matched_fields`). Explicit pins do not help
locate the header row (scoring counts detected fields only).

Header-text normalization here (`_normalize_header_text`) is intentionally
looser than and separate from `NormUnit`/`NormCode` (see §9.2 for the same
"header normalization is its own thing" note).

### 10.3 Current scope and deferred rules

Implemented in `core/layout.py`:

- **Column detection** for `work_name` / `unit` / `code` / `base_price` /
  `average_price`
  (R6-R9), with a minimal header-row locator, the required-fields hard-stop
  (R20), and a human-readable resolution report (R19).
- **Regional coefficient by label** (R16, `resolve_regional_coefficient`):
  primary pattern is a "Region" label with a "Coefficient" label directly
  below it, region name to the right of the region label and the numeric
  coefficient to the right of the coefficient label; fallbacks are a
  standalone coefficient label with a number to its right, then a
  caller-provided explicit value, then the default `1.0`. This replaces the
  VBA's single fixed coefficient cell (§6) as the primary source, keeping
  the fixed cell only as an explicit fallback.
- **Average-column placement** (R12, `resolve_average_placement`): the
  average price is written into the column immediately right of the detected
  base price; if that neighbour already holds data, a new column is inserted
  there. This function only decides the target column and whether an insert
  is needed; the Excel writer performs the actual insert/shift.
  **(2026-07 fix)** `Worksheet.insert_cols()` (openpyxl) moves cell
  values/styles but does **not** rewrite formulas living in other cells of
  the sheet — any pre-existing formula referencing a column at or past the
  insertion point silently ends up pointing at the wrong data once the
  insert has shifted that data one column right. This was the root cause of
  a real bug: "ИТОГО" totals summing the neighbouring column, and a stale
  leftover average-price formula (from re-processing an already-processed
  file) surviving next to a newly inserted one with a different, also
  wrong, `AVERAGE()` range. `_shift_formulas_after_insert` in
  `core/excel_writer.py` now re-points every existing formula on the sheet
  right after each `insert_cols` call; `_plan_average_column` also now
  recognizes an existing average-formula cell (see the 2026-07 note above)
  and reuses it instead of inserting a duplicate. If the layout resolver
  detects an `average_price` header immediately to the right of the base
  price, that column is reused regardless of whether its populated body cells
  are formulas or manual numbers.
- **Sheet selection** (R1, `rank_sheets` / `select_sheets`): every worksheet
  is scored by how well its layout resolves. A single sheet is used
  automatically; when several sheets qualify, the result flags
  `needs_user_choice` and lists candidates (intended for the web flow where
  the user picks one or more sheets as buttons); an explicit selection by
  title is also supported. The sheet may be named anything — selection is by
  content, not by a fixed name/order.
- **Blank-row-tolerant data range** (R5, partial, `data_row_numbers`): body
  rows are read tolerating isolated blank rows — the scan skips blank rows
  (all key columns empty) and stops only after `data_scan.max_blank_run`
  consecutive blank rows, so one or two stray empty rows inside the table do
  not truncate the read. Fully-empty columns in the body are irrelevant
  because only the resolved key columns are inspected.

Still deferred (agreed, not yet built) — tracked in `docs/OPEN_ITEMS.md`:
richer header detection incl. merged-cell/sub-header handling (R2-R4),
total/summary-row skipping (rest of R5), section/quantity detection
(R10-R11), analog-column placement in the first free column (R13), and
tolerant number/code parsing at detection time (R17-R18).

## 11. Expert catalog corrections and explicit layer multiplicity

These are post-VBA product rules explicitly approved in 2026-07. They change
the former direct admin-edit behavior and therefore are implemented separately
from the original macro ports.

### 11.1 Correction approval and audit

- A row update, row deletion, or bulk catalog action creates a `pending`
  correction request. It does not mutate `catalog_items`.
- Every request stores the stable source identity (catalog source, source
  filename, and source row), the old and proposed values, reason, submitting
  actor/role, and submission time.
- Only the `senior` or `admin` role may approve or reject a request.
- Approval and application occur in one database transaction. Rejection never
  changes the catalog.
- Submission, approval/rejection, application, and later reapplication are
  appended to `catalog_correction_events`.
- Approved corrections are durable overlays. After a catalog row is rebuilt
  with a new database ID, the approved update or deletion is found by its
  stable source identity and applied again.
- The eight corrections from expert review `6444312` are stored as approved
  seeded requests, not as silent hard-coded updates.

Login/password authentication is not part of this rule yet. The current local
UI passes placeholder specialist/senior identities into the same workflow that
future authenticated sessions will use.

### 11.2 Explicit layer-count compatibility

The original VBA matching key remains normalized unit + normalized code, with
the demolition rule unchanged. A new deterministic post-filter applies when
both work names explicitly state a layer count:

- equal layer counts are compatible;
- different explicit counts are incompatible;
- if either side does not state a count, matching remains allowed.

Layer wording is configured in `data/config/multiplicity.json`. This prevents a
four-layer waterproofing analog from being assigned to an explicitly
two-layer estimate without introducing semantic or LLM matching.
