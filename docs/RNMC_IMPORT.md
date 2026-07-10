# RNMC Import Workflow

This document describes the current web import workflow for adding new RNMC
files to the SQLite catalog.

## Current status

Implemented in the `feature/admin-ui` branch:

- legacy `File_Log.xlsx` import into `imported_files`;
- ZIP dry-run for RNMC files;
- ZIP import-log recording without catalog rows;
- row preview for `.xlsx` / `.xlsm` workbooks;
- real ZIP import into `catalog_items`;
- import control center in `/admin/imports`;
- per-file detail pages with catalog rows and rejected-row logs;
- automatic detection of `lsr_quarter`, planned start, planned finish, object region, and regional coefficient from `.xlsx` / `.xlsm` workbooks;
- strict mapping of RNMC value columns: unit price, total cost, and labor columns;
- manual metadata edits and retry unlock for `failed` / `no_data` records.

## File identity and duplicates

File identity is based on the normalized base filename only.

If a filename already exists in `imported_files`, the file is treated as already
processed and normal import skips it. The region/folder is not part of the
dedup key.

This is an explicit product rule: RNMC filenames are expected to be unique. If
the same filename appears in different folders/regions, that is a data-quality
violation rather than a new file version. Duplicate filenames inside one ZIP are
marked with status `duplicate_name`.

## Legacy `File_Log.xlsx`

Legacy logs can be imported from `/admin/imports`.

Rules:

- every row in `File_Log.xlsx` is considered already processed;
- this includes legacy notes such as `нет данных`, `новая РНМЦ`, and `0`;
- `новая РНМЦ` means the file was processed manually and should not be picked up
  automatically again;
- the old `Status` column is stored as `legacy_note`;
- if the old `Status` value is numeric, it is also stored as `rows_ok`;
- `Год Квартал ЛСР`, planned start, and planned finish are stored for later use.

Legacy rows receive status `legacy_imported`. Duplicate filenames in the legacy
log are marked as `duplicate_name` after the first occurrence.

## ZIP upload modes

`/admin/imports` supports four RNMC ZIP actions:

1. **Dry-run** — scans the archive and reports which files will be processed,
   skipped, or marked duplicate. No database writes.
2. **Record ZIP in import log** — records new files as `pending`, skipped files
   as `skipped`, and duplicate names as `duplicate_name`. No catalog rows are
   imported.
3. **Row preview** — opens `.xlsx` / `.xlsm` files and shows a readable
   tabbed result: summary, file statuses, workbook metadata, detected source
   headers, and preview rows. Final already-processed filenames are skipped
   before workbook bytes are read. Preview shows at most 30 real body rows per
   workbook after the detected header row; blank technical rows before the table
   body do not consume the limit. The real catalog import still validates and
   imports all rows. No catalog rows are imported.
4. **Import rows into catalog** — validates and writes accepted rows to
   `catalog_items`, updates `imported_files`, stores detected workbook metadata,
   and writes rejected-row details to `import_row_log`.

## Row preview UI

The preview result is intentionally split into tabs so large RNMC batches remain
reviewable:

- **Summary** aggregates counts by status and row outcome.
- **Files and statuses** shows one row per workbook with status, reason, sheet,
  header row, task number, OK/rejected counts, and preview-limit marker.
- **Metadata** shows resolved region, regional coefficient, LSR quarter, planned
  start, and planned finish.
- **Headers** shows the original Excel header texts that were mapped to code,
  work name, unit, quantity, unit price, total price, and labor columns.
- **Rows** shows up to 30 real body rows per workbook with normalized unit price
  without VAT, total without VAT, labor fields, and preview-only row issues.

Preview tables include client-side filters for search, status, already-processed
files, problem rows, and empty rows. They also include a local table zoom/density
control so large RNMC batches can be reviewed without changing browser zoom.

## Region handling

Default region = immediate parent folder inside the ZIP, unless a workbook consolidation block provides an object region.

Example:

```text
upload.zip
  Moscow/file1.xlsx   -> region = Moscow
  Yakutia/file2.xlsx  -> region = Yakutia
```

The admin UI also accepts a manual region override. When provided, it applies to
all files in that upload action and wins over workbook-detected region values.

For newer RNMC templates, the parser can read workbook labels such as `Регион
расположения объекта`, `Регион объекта`, or `Регион`. When found and no manual
override is provided, this value is used as the imported file region and copied
to imported catalog rows. If no workbook region is detected, the ZIP folder name
remains the fallback.

## Workbook parsing rules

The parser mirrors the legacy VBA importer where possible:

- scans up to 400 rows and 150 columns for a header row;
- required logical headers: work name, unit, and quantity;
- supported unit headers include `Ед.изм.`, `Ед.изм`, and `Единица измерения`;
- extracts one task number per workbook from `№ задачи 1Ф` / `№ задачи`;
- scans the first rows/columns for metadata labels such as `Год Квартал ЛСР`,
  planned start, planned finish, object region, and regional coefficient;
- reads the first worksheet that has a matching header and accepted rows;
- skips technical numbering rows directly below the header, such as `1 | 2 | 3 | ...`;
- skips section/subsection rows that have a name but no unit, quantity, code, price, or total;
- stops after 3 consecutive blank rows across number/name/unit/quantity;
- ignores temporary Excel lock files starting with `~$`;
- `.xlsx` and `.xlsm` are supported for preview/import;
- `.xls` is discovered as an Excel file, but detailed parsing is deferred.

## Metadata detection

The RNMC parser now detects and stores workbook-level metadata during preview
and real import:

- `lsr_quarter`;
- `planned_start`;
- `planned_finish`;
- `region_folder`;
- `regional_coefficient`.

Detection is intentionally conservative and deterministic. It looks for known
Russian labels in the workbook, then reads an inline value after a separator or
nearby values to the right / below the label. Bare local-estimate labels such as `ЛСР 02-01-01 ...` are not treated as the `Год/квартал ЛСР` metadata value. Parseable dates are normalized to
ISO format (`YYYY-MM-DD`). Quarter values such as `1 квартал 2026`, `IV кв.2025 г.`, and `2 кв. 25г.` are
normalized to `2026 Q1`, `2025 Q4`, and `2025 Q2`. Legacy note text such as
`в ценах IV кв.2025 г.` and monthly work periods such as `с июня 2026 г. по
сентябрь 2026 г.` are also parsed. Month-only periods are stored as the first
day of the month in ISO format. Excel serial date values near start/finish
labels are converted to ISO dates only within a plausible metadata date range, so numeric prices/codes are not mistaken for dates.

Newer consolidation blocks are also parsed. Region labels such as `Регион
расположения объекта`, `Регион объекта`, or `Регион` can override the ZIP folder
when no manual region override is supplied. `Региональный коэффициент` /
`Коэффициент` is parsed as a numeric value with comma or dot decimals, stored on
`imported_files.regional_coefficient`, and copied to imported
`catalog_items.regional_coefficient` rows. Implausible values such as years (`2026`, `2027`) are ignored.

Detected `lsr_quarter`, `planned_start`, and `planned_finish` are stored both on
`imported_files` and on every imported `catalog_items` row from that workbook.
Manual editing remains available on the import detail page because real RNMC
layouts may require further iterations and additional label patterns. When a
regional coefficient is entered manually for an import file, it is propagated to
that file's catalog rows and their ZLVL unit price is recalculated.

## RNMC value column mapping

The catalog stores deterministic numeric values only. Workbooks are opened with
`data_only=True`, so formula cells are read from their cached calculated value
when Excel has saved one. Numeric parsing accepts spaces, comma decimals, and
dot decimals, then stores SQLite `REAL` values.

`Кол-во` is stored in `catalog_items.quantity` as the source row quantity.

The catalog stores unit price in two levels. `catalog_items.price_original` is
the source regional unit price without VAT. `catalog_items.price_zlvl` is the
ZLVL / Moscow-level unit price. `catalog_items.price` remains the working price
used by matching/pricing and equals `price_zlvl`. For new RNMC ZIP files,
`price_original` comes from the workbook and `price_zlvl = price_original /
regional_coefficient` when a valid coefficient is available. If the source
header says `с НДС`, the value is divided by `1.2` before `price_original` is
stored. If the source header says `без НДС`, the value is stored as-is.

`Итого стоимость` is stored separately in `catalog_items.total_price` as one
without-VAT total value. Total cost is not stored in two ZLVL/original levels.
Average headers such as `Цена средняя`, `Итого стоимость средняя`, or `ср знач`
are not used as source values for either unit price or `total_price`.

Labor columns are stored separately:

- `ТЗ на ед., чел-час` -> `labor_unit`;
- `ТЗ всего, чел-час` -> `labor_total`;
- `ТЗм на ед., чел-час` -> `machine_labor_unit`;
- `ТЗм всего, чел-час` -> `machine_labor_total`;
- `ТЗр на ед., чел-час` -> `labor_unit`;
- `ТЗр всего, чел-час` -> `labor_total`;
- `ЗТР на ед., чел-час` -> `labor_unit`;
- `ЗТР всего, чел-час` -> `labor_total`.


## Code column selection

`Код раздела` is never a catalog code. The parser first looks for RNMC code/list
headers such as `Перечень ГЭСН/ФЕР/ТЕР/КР`, `Перечень ГЭСН`, `Обоснование`, or
`Шифр`. A generic `Код` header is accepted only if it is not `Код раздела`.

Some real RNMC workbooks have a blank header cell immediately before `Код
раздела`, while the body cells in that blank column contain the actual GESN/FER
code. In that specific layout, the importer uses that unlabeled column as the
catalog code column and shows it in preview as `[без заголовка перед Код
раздела]`.

## Catalog row validation

A row is written to `catalog_items` only when it has:

- task number;
- valid code / GESN / FER / перечень value;
- valid unit;
- positive numeric price.

Rows with missing section data are skipped when they are clearly section/subsection labels. Data rows are written to `catalog_items` only when they pass validation. Rejected data rows are not silently lost; they are written to `import_row_log` with
an Excel row number and reason, for example:

- `missing_task_number`;
- `missing_or_invalid_code`;
- `missing_or_invalid_unit`;
- `missing_or_invalid_price`;
- `missing_unit_and_quantity`.

## Statuses

Current import statuses:

- `legacy_imported` — imported from old `File_Log.xlsx` and treated as already
  processed;
- `pending` — recorded from ZIP, not yet imported into catalog, eligible for
  future processing;
- `success` — rows imported into `catalog_items`;
- `no_data` — workbook/header found but no valid catalog rows, or required
  headers were not found;
- `failed` — parse/import failed;
- `skipped` — filename already exists in import history;
- `duplicate_name` — duplicate filename rule violation.

Final statuses block normal re-import by filename. `pending` remains eligible
for row preview/import.

## Retry model

Current retry is explicit and requires the user to upload the ZIP again.

For `failed` and `no_data` records, the detail page has an action to allow retry.
This changes the record to `pending`, so the same filename can be processed when
the ZIP is uploaded again.

True one-click retry without re-upload is deferred until the service stores the
original uploaded Excel files in a durable file store.

## Deferred improvements

- expand metadata label patterns based on real RNMC files that are not detected
  yet;
- `.xls` reader support;
- durable storage of source files for one-click retry;
- richer duplicate-name review UI;
- richer rejected-row diagnostics and exports.

## Legacy ZLVL catalog reload

`РНМЦ_КА_ЖО_ZLVL_V3.xlsx` is imported as a prepared catalog, not as a raw RNMC
ZIP file. Its direct column mapping is:

- `Номер задачи` -> `task_id`;
- `Наименование работ` -> `work_name`;
- `Ед.изм.` -> `unit`;
- `Кол-во` -> `quantity`;
- `Цена единицы работ (с учетом вспомогательных материалов), руб. без НДС` -> `price_original`;
- `Цена единицы работ (с учетом вспомогательных материалов), руб. без НДС ZLVL` -> `price_zlvl` and `price`;
- `Итого стоимость, руб. без НДС` -> `total_price`;
- `ТЗ на ед., чел-час` / `ТЗ всего, чел-час` -> `labor_unit` / `labor_total`;
- `ТЗм на ед., чел-час` / `ТЗм всего, чел-час` -> `machine_labor_unit` / `machine_labor_total`;
- `Перечень ГЭСН/ФЕР/ТЕР/КР` -> `code`;
- `source_file` -> `source_filename`;
- `Регион` -> `region`;
- `Год Квартал ЛСР` -> `lsr_quarter`;
- `Планирумый срок начала работ` -> `planned_start`;
- `Планируемый срок окончания работ` -> `planned_finish`;
- `Региональный коэффициент` -> `regional_coefficient`;
- `Дата добавления в каталог` -> `added_date`.

For this legacy reload, `Итого стоимость, руб. с НДС` is ignored.

Each distinct `source_file` value from the prepared catalog is also recorded in
`imported_files` with status `legacy_imported`. This preserves the global
filename-based skip rule: if the same source RNMC later appears in a ZIP upload,
dry-run/preview/import can skip it by normalized base filename without opening the
workbook again. The consolidated catalog file itself is also recorded as a
`success` import for auditability.
## Admin import workflow

The regular admin workflow is intentionally limited to one staged action:

1. Upload one RNMC ZIP and run preview.
2. Review summary, file statuses, metadata, detected headers, and sample rows.
3. Confirm the staged ZIP to write catalog rows and update `imported_files` automatically.

Legacy `File_Log.xlsx` migration and low-level log-only endpoints remain available for compatibility and tests, but they are not shown in the regular admin interface. One-time catalog migrations should be run as maintenance scripts instead of permanent UI actions.


## Processed-file history during rebuilds

`imported_files` is the durable replacement for legacy `File_Log.xlsx` and is preserved when catalog rows are cleared for a rebuild. ZIP preview uses both final `imported_files.filename_key` values and distinct `catalog_items.source_filename` values as a recovery fallback. This prevents already loaded source files from being parsed again when an older catalog import reconstructed the log incompletely.
