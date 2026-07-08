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
- automatic detection of `lsr_quarter`, planned start, and planned finish from `.xlsx` / `.xlsm` workbooks;
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
3. **Row preview** — opens `.xlsx` / `.xlsm` files and shows detected rows,
   task number, workbook metadata, sheet name, header row, and rejected counts.
   No catalog rows are imported.
4. **Import rows into catalog** — validates and writes accepted rows to
   `catalog_items`, updates `imported_files`, stores detected workbook metadata,
   and writes rejected-row details to `import_row_log`.

## Region handling

Default region = immediate parent folder inside the ZIP.

Example:

```text
upload.zip
  Moscow/file1.xlsx   -> region = Moscow
  Yakutia/file2.xlsx  -> region = Yakutia
```

The admin UI also accepts a manual region override. When provided, it applies to
all files in that upload action.

Future work may add automatic region detection from workbook content, but folder
name remains the current default convention.

## Workbook parsing rules

The parser mirrors the legacy VBA importer where possible:

- scans up to 400 rows and 150 columns for a header row;
- required logical headers: work name, unit, and quantity;
- supported unit headers include `Ед.изм.`, `Ед.изм`, and `Единица измерения`;
- extracts one task number per workbook from `№ задачи 1Ф` / `№ задачи`;
- scans the first 120 rows and 60 columns for metadata labels such as
  `Год Квартал ЛСР`, planned start, and planned finish;
- reads the first worksheet that has a matching header and accepted rows;
- stops after 3 consecutive blank rows across number/name/unit/quantity;
- ignores temporary Excel lock files starting with `~$`;
- `.xlsx` and `.xlsm` are supported for preview/import;
- `.xls` is discovered as an Excel file, but detailed parsing is deferred.

## Metadata detection

The RNMC parser now detects and stores workbook-level metadata during preview
and real import:

- `lsr_quarter`;
- `planned_start`;
- `planned_finish`.

Detection is intentionally conservative and deterministic. It looks for known
Russian labels in the workbook, then reads an inline value after a separator or
nearby values to the right / below the label. Parseable dates are normalized to
ISO format (`YYYY-MM-DD`). Quarter values such as `1 квартал 2026`, `IV кв.2025 г.`, and `2 кв. 25г.` are
normalized to `2026 Q1`, `2025 Q4`, and `2025 Q2`. Legacy note text such as
`в ценах IV кв.2025 г.` and monthly work periods such as `с июня 2026 г. по
сентябрь 2026 г.` are also parsed. Month-only periods are stored as the first
day of the month in ISO format. Excel serial date values near start/finish
labels are converted to ISO dates.

Manual editing remains available on the import detail page because real RNMC
layouts may require further iterations and additional label patterns.

## RNMC value column mapping

The catalog stores deterministic numeric values only. Workbooks are opened with
`data_only=True`, so formula cells are read from their cached calculated value
when Excel has saved one. Numeric parsing accepts spaces, comma decimals, and
dot decimals, then stores SQLite `REAL` values.

The main matching price is `catalog_items.price`, which means unit price without
VAT. The importer accepts unit-price headers with auxiliary materials when the
header is still a unit price. If the source header says `с НДС`, the value is
divided by `1.2` before storage. If the source header says `без НДС`, the value
is stored as-is.

`Итого стоимость` is stored separately in `catalog_items.total_price`, also
normalized to without VAT by dividing `с НДС` source values by `1.2`. Average
headers such as `Цена средняя`, `Итого стоимость средняя`, or `ср знач` are not
used as source values for either `price` or `total_price`.

Labor columns are stored separately:

- `ТЗ на ед., чел-час` -> `labor_unit`;
- `ТЗ всего, чел-час` -> `labor_total`;
- `ТЗм на ед., чел-час` -> `machine_labor_unit`;
- `ТЗм всего, чел-час` -> `machine_labor_total`;
- `ТЗр на ед., чел-час` -> `labor_unit`;
- `ТЗр всего, чел-час` -> `labor_total`.

## Catalog row validation

A row is written to `catalog_items` only when it has:

- task number;
- valid code / GESN / FER / перечень value;
- valid unit;
- positive numeric price.

Rejected rows are not silently lost. They are written to `import_row_log` with
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
