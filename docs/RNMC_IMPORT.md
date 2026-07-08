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
   task number, sheet name, header row, and rejected counts. No catalog rows are
   imported.
4. **Import rows into catalog** — validates and writes accepted rows to
   `catalog_items`, updates `imported_files`, and writes rejected-row details to
   `import_row_log`.

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
- extracts one task number per workbook from `№ задачи 1Ф` / `№ задачи`;
- reads the first worksheet that has a matching header and accepted rows;
- stops after 3 consecutive blank rows across number/name/unit/quantity;
- ignores temporary Excel lock files starting with `~$`;
- `.xlsx` and `.xlsm` are supported for preview/import;
- `.xls` is discovered as an Excel file, but detailed parsing is deferred.

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

- automatic extraction of `lsr_quarter`, planned start, and planned finish from
  real RNMC workbooks;
- `.xls` reader support;
- durable storage of source files for one-click retry;
- richer duplicate-name review UI;
- richer rejected-row diagnostics and exports.
