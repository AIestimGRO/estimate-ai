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

## 6.1 RNMC catalog import value normalization

For RNMC ZIP imports, `catalog_items.price` stores unit price without VAT. Unit
price headers with auxiliary materials are valid source columns when the header
still describes a unit price. Unit-price source values marked `с НДС` are
divided by 1.2 before storage; values marked `без НДС` are stored as-is.

`Итого стоимость` is stored separately as `total_price` and follows the same VAT
normalization rule. Average values are not source values: headers containing
`средняя` or `ср знач` are ignored for both unit price and total price.

Labor import mappings are deterministic: `ТЗ` and `ТЗр` unit/total values map to
`labor_unit` / `labor_total`; `ТЗм` unit/total values map to
`machine_labor_unit` / `machine_labor_total`. Formula cells are read as cached
calculated values, and numeric strings with comma or dot decimal separators are
normalized to SQLite numeric values. File-level regional coefficient from RNMC
consolidation metadata is stored as metadata and copied to imported catalog rows;
it does not modify stored catalog prices.

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

### 9.6 Decisions for the DB/web RNMC import module

These decisions supersede the original VBA importer where noted. Do not silently
reinterpret them.

1. **Dedup/skip is by normalized base filename only.**
   - Normal repeated imports must skip any filename already present in
     `imported_files`, regardless of region/folder.
   - This is intentional: RNMC filenames are expected to be globally unique. If
     the same filename appears in two different folders/regions, that is a
     data-quality violation, not a new regional version of the same file.
   - Duplicate filenames inside one uploaded ZIP are recorded as
     `duplicate_name`.
   - Content-hash-based dedup was considered unnecessary for the current
     workflow. Filename is the business key.

2. **Legacy `File_Log.xlsx` rows are already processed.**
   - All legacy rows count as processed, including rows whose old `Status` value
     is `нет данных`, `новая РНМЦ`, `0`, or another manual note.
   - The old `Status` value is not treated as the new status. It is stored as
     `legacy_note`; if numeric, it is also stored as `rows_ok`.
   - Legacy rows use status `legacy_imported` unless they are duplicate filename
     violations.

3. **Validate at import time and store row diagnostics.**
   - The DB/web path validates rows before writing them into `catalog_items`.
   - Required values for a catalog row: task number, valid code, valid unit, and
     positive numeric price.
   - Rejected rows are written to `import_row_log` with Excel row number and a
     reason, not silently dropped.

4. **Region comes from manual override, workbook metadata, then ZIP folder.**
   - Manual upload override has highest priority and applies to all files in the upload action.
   - Without manual override, newer RNMC workbook labels such as `Регион расположения объекта`, `Регион объекта`, or `Регион` can supply the region.
   - If workbook region is not detected, the immediate parent folder inside the ZIP archive is used.
   - Detected region is stored on `imported_files.region_folder` and copied to imported catalog rows.
   - `Региональный коэффициент` / `Коэффициент` is stored as metadata and copied to imported catalog rows, but catalog prices are not multiplied by it.

5. **One task number per file and first matching worksheet are kept.**
   - The parser extracts one task number per workbook.
   - The first worksheet with a matching table and accepted rows wins.
   - Multi-task/multi-sheet import is not part of the current RNMC workflow.

6. **Retry is explicit.**
   - `failed` and `no_data` records are not automatically retried.
   - The admin UI can allow retry by changing the record to `pending`.
   - The user then uploads the ZIP again; true one-click retry without re-upload
     is deferred until original source workbooks are stored durably.

### 9.7 Current web implementation

The current admin/import implementation is documented in `docs/RNMC_IMPORT.md`.
In short:

- `/admin/imports` imports legacy `File_Log.xlsx` into `imported_files`.
- ZIP dry-run reports `will_process`, `skipped_processed`, and
  `duplicate_name` without database writes.
- ZIP log recording writes new files as `pending` without catalog rows.
- ZIP row preview opens `.xlsx` / `.xlsm` files and reports detected rows.
- ZIP row preview must skip final already-processed filenames before reading
  workbook bytes from the archive. Final statuses include legacy imported,
  success, skipped, no_data, duplicate_name, and manual_checked; pending and
  failed remain previewable for retry checks.
- ZIP row preview is intentionally limited to 30 real catalog body rows per
  workbook for UI performance. The limit starts after the detected header row;
  Excel column-number helper rows and section/title rows do not consume it. This
  limit affects preview counts only; real catalog import must still read and
  validate all rows.
- ZIP row preview should render large batches in separate views for summary,
  file statuses, workbook metadata, detected source headers, and row samples.
- ZIP row preview UI should provide column-style client-side filters to hide
  already-processed files, show only problem rows, hide empty rows, and search
  within the currently rendered table. These UI filters must not change import
  behavior or matching/pricing logic.
- ZIP row preview UI may provide local table zoom/density controls. These are
  presentation-only controls and must not change parsed values.
- RNMC `Код раздела` is a section classifier, not an analog code. It must never
  populate `catalog_items.code`. The code column should be selected from headers
  such as `Перечень ГЭСН/ФЕР/ТЕР/КР`, `Перечень ГЭСН`, `ГЭСН/ФЕР/Перечень`,
  `Обоснование`, or `Шифр`; generic `Код` is only a last-resort candidate when
  it is not a section-code header.
- RNMC section/title rows are skipped, not imported. A row with a title/work-name
  but no code, unit, quantity, price, or total is not a catalog row and should
  not be counted as a rejected work row.
- RNMC technical column-number rows immediately below headers, for example
  `1 | 2 | 3 | ...`, are skipped before preview/import row validation.
- ZIP catalog import writes accepted rows to `catalog_items`, writes rejected
  rows to `import_row_log`, and updates `imported_files` with statuses such as
  `success`, `no_data`, `failed`, `skipped`, and `duplicate_name`.
- Per-file detail pages show metadata, imported catalog rows, rejected rows, and
  retry controls for `failed` / `no_data`.

`core/ingest.py` remains a lower-level, storage-agnostic helper. The web RNMC
ZIP import path is implemented in `app/services/rnmc_zip.py`,
`app/services/rnmc_excel.py`, and the storage helpers in
`core/storage/catalog.py`.

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

- **Column detection** for `work_name` / `unit` / `code` / `base_price`
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
