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

There is no separate "multiplier" extraction (e.g. "100 m" is not split
into unit="m" + multiplier=100 in current code) — the unit string is
normalized and compared as a whole token. "100 м" and "м" are currently
**different** matching keys.

### 2.3 Matching key (`AnalogSearchKey`, Module3)

```
key = NormUnit(unit) & "||" & NormCode(code)
```

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
  2.0, configurable on Instrument sheet).
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
- Confirmed against the VBA source comment in `MarkOutOfApprovedRange`:
  entries with `added_date_serial <= 0` (no recorded catalog added-date)
  are never flagged against an approved range, regardless of how far
  outside the range their price is. The macro intentionally treats them as
  "not new enough to check" to avoid false positives on data with missing
  dates.
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
  with at least one analog (not for rows without analogs). It is added
  only if not already present, after stripping CR/LF/tab/nbsp from the
  base code.
- **Average price formula**: `=MAX(base_price, IFERROR(AVERAGE(base_price,
  analog_range), base_price))` — written as an actual Excel formula, not a
  static value, into the average/total price column. This guarantees the
  output is never below the base price.
- **Regional coefficient**: read from a single configurable cell address
  on the estimate sheet (e.g. `F12`); defaults to 1 if blank, non-numeric,
  ≤ 0, or the cell address itself is not configured. Applied as a
  multiplier only to the displayed/written analog prices (and to the
  min/max/recommended values logged to `Price_Check_Log`) — catalog prices
  themselves are never modified.
- **Cell coloring**: blue = task marked in the color list; red
  (`problemFill`) = analog is part of a flagged price problem (per §5);
  grey (`dupFill`) = a second-or-later price within the same task
  (visual de-emphasis, not a data quality flag).

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

## 8. What is explicitly NOT happening in current logic

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
   the same name will be silently skipped unless its FileLog row is
   manually deleted first. Is filename-based dedup sufficient for the DB
   version, or should it be content-hash-based?
2. **No data validation at import time** — prices, codes, units are copied
   raw. Should the DB import path validate/normalize at ingestion (catch
   bad data early, at the cost of rejecting more files) or keep the
   current "import everything, validate later at matching time" approach?
3. **One task number per file, first-matching-sheet-only** — if a real
   RNMC file ever has multiple distinct task tables across multiple
   sheets, only the first is imported today. Worth confirming this
   matches how source files are actually structured in practice before
   assuming it's fine to keep.
4. **Region from folder name is a strict operational dependency** — the
   future "watched folder" needs documented, enforced subfolder-per-region
   structure, or region data will be silently wrong (folder name still
   gets used, just incorrectly) rather than failing loudly.
5. **Failed imports require manual FileLog cleanup to retry** — worth
   deciding whether the DB version should auto-retry failed files (e.g.
   distinguishing "permanently bad file" from "transient error") rather
   than requiring a manual log edit.

### 9.7 Implication for the planned database

Conceptually this module is already an ETL pipeline: walk folder → parse
file → normalize headers → map columns → write rows. Porting it to write
to a database table (e.g. `catalog_items`) instead of an Excel sheet range
is a natural fit — the folder-walk, header-detection, and row-mapping
logic carry over largely unchanged; only the write target changes. The
"logged files" dedup dictionary becomes a query against an
`imported_files` table instead of reading a sheet column. The open
questions in §9.6 should be resolved as explicit decisions during that
port, not inherited silently.
