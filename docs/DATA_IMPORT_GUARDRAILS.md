# Data provenance and high-risk import rules — RNMC, TKP, Excel

Updated 2026-10-28. Use this along with `DOMAIN_RULES.md`, `RNMC_IMPORT.md`,
`core/tkp_folder_ingest.py`, `core/storage/tkp.py`, tests and original source
workbooks. This is a guardrail document, **not** a newly approved algorithm.

## 1. What is a price source?

| Source | Tables / files | How to interpret it |
| --- | --- | --- |
| RNMC | `catalog_sources`, `catalog_items` | Historical RNMC catalog; original/ZLVL prices, total, labor, coefficient and provenance. |
| TKP winner | `tkp_sources`, `tkp_items` | Commercial proposal of the **selected winner**, with task/file/winner metadata and source row. |
| Other KL participant | Original KL workbook only unless specifically persisted | **Not necessarily in the main TKP database.** Can have different price from winner; requires reading its own participant group. |
| Reserve winner | Original KL workbook unless explicitly represented | Never assume `winner_*` means reserve winner or vice versa. |
| Estimate/reference columns | KL/RNMC reference values | **Not commercial price offered by participant** unless proven. |

SQLite schema reference: `core/storage/schema.py` (`SCHEMA_VERSION=12` in
this snapshot). The `tkp_items` table stores `winner_unit_price_no_vat`,
`winner_line_total_no_vat`, `qty`, `qty_source_text`, `winner_group_index`,
`winner_start_col` and `winner_method`/block audit fields. `tkp_sources` stores
file name, task, procedure, selected winner and import status. Check schema
with PRAGMA rather than assuming fields exist in all older DB versions.

## 2. RNMC filename-only dedup — binding user decision

- A file is processed by its normalized **basename** (case-insensitive
  according to implementation), independent of its ZIP directory/region.
- Duplicates in different directories are a **data quality violation**, not
  separate price sources. Surface `duplicate_name`.
- `File_Log.xlsx` records count as processed even with legacy status
  `нет данных`, `новая РНМЦ` or zero rows; `новая РНМЦ` means *previously
  handled*. Preserve legacy status as `legacy_note`.
- Preview must not commit accepted rows; only confirmed import writes to
  `catalog_items` and tracks `imported_files`/`import_row_log`.
- `.xlsx`/`.xlsm` detailed parsing is supported. `.xls` is detected but
  detailed support is deferred. Averages (`Цена средняя`, `Итого стоимость
  средняя`) are not valid substitute source prices.
- Region: manual override wins; else workbook metadata may win; else parent
  ZIP folder. Region is not an RNMC matching filter.
- Rejected/excluded rows and no-data statuses are meaningful audit records.
  Retry `failed/no_data` is a distinct workflow, not automatic reprocessing.

Current detail: `docs/RNMC_IMPORT.md`, parser `app/services/rnmc_excel.py`,
tests `test_rnmc_*`, `test_admin_imports.py`.

## 3. TKP ingestion — why original KL review is essential

`core/tkp_folder_ingest.py` parses workbook structure, participant groups,
recommended-winner block, task metadata and WOR rows. `core/storage/tkp.py`
performs persistence and versioned upsert. `core/tkp_ingest.py` supports an
aggregated CatalogBuilder source. Do not confuse those import paths.

A reliable record must establish:

1. **Document identity:** source filename, content/revision, task number,
   request date, sheet and source row; avoid task/filename duplicates.
2. **Winner identity:** determine winner from explicit recommended-winner
   fields, reconcile name/UIN/INN with the participant group. If missing or
   ambiguous, flag for manual review; don't choose a low price blindly.
3. **Participant group layout:** variants use different column spacing and
   headers. Verify `winner_start_col`, `winner_group_index`, unit-price and
   line-total columns **against original cells**. A standard offset is not
   safe for all input workbooks.
4. **Value meaning:** distinguish RNMC baseline, commercial winner price,
   commercial competitor price, reserve winner, total with VAT and total
   without VAT. Don't replace missing winner price with RNMC baseline.
5. **Arithmetic:** where quantity/price/total are all valid, validate
   `unit price × quantity ≈ line total`; handle legitimate rounding/formulas
   carefully. Unexplained mismatches must be held from automatic pricing.
6. **Unit & specification:** verify distance/dimension (`м`, `м2`, `м3`,
   `100 м2`), diameters, pressure, materials, technology, number of layers
   and presence of demolition; do not promote a textually similar but
   economically different job as an exact analog.
7. **Provenance:** keep cell coordinates, names, source file and task in
   the audit/export/report. No invented prices or hidden extrapolation.

### October 2026 original-KL supplement (historical)

From 156 KL files dated approximately 15.08–08.10, a separate checked import
added **97 source files / 2,838 item rows** to an existing 536/31,925 DB,
giving 633/34,763. The production copy was installed 27.10 after staged
merge, SQLite checks and backup. **59** documents were not included for
specified validation/duplicate reasons; **240** disputed/arithmetically
inconsistent pricing rows were kept out of automated pricing.

This supplement was prepared with additional original-cell checks because
some source KL structures were not handled correctly by the ordinary parser.
**The user's deployed database may be richer than what rerunning the current
UI folder parser would produce.** Do not bulk-reimport or rebuild without
reconciliation. A future parser change must be tested on the original files;
regressions for tasks `6940345` and `6941030` should cover winner/group
column selection, source provenance, and price arithmetic.

The source ZIP being delivered contains no live `.db` and no full set of real
KL files. To reproduce the import, obtain originals through an approved
channel and construct anonymized fixtures.

## 4. Strict commercial analog comparisons (ad hoc analyses)

The October "FRIZ" comparison work was a **separate analytical spreadsheet
workflow**, stricter than the built-in live TKP lexical matcher:

- Match names and technical qualifiers tightly, preserve dimensions and
  installation/demolition action; strict normalized units.
- Compare vendor-specific historical proposal only against actual proposal
  cells for that vendor; absence from winner-only DB does not prove absence
  from original KL documents.
- For other participants/reserve winners, use original KL and show their
  exact status, not an inferred winner label.
- Record source file, task, exact row/cell, vendor, unit, price **without VAT**
  and ambiguity status. Multiple different prices for the same apparent job
  within one task are ambiguous; don't silently pick one.
- Separate historically earlier source dates from later bids where temporal
  comparison matters. Rows from after the source RFQ date should not be
  silently treated as older historical prices.
- Retain rejected/ambiguous candidates on a separate explanation sheet when
  the user requests a comparison; fewer verified matches beat many false ones.

**Do not claim** that an externally compiled strict Excel shortlist is the
same behavior as live `core/tkp_matching.py`. The latter uses lexical ranking,
unit match bonus and a score threshold, whereas the strict shadow path uses
`core/tkp_candidate_rules.py`.

## 5. EKR section codes: document-safe editing

The application determines a section from GESN codes using
`core/sections.py::ResolveSectionCode(code, is_demolition)` and overrides,
where applicable, from the approved/project-specific correction workflow.
`ResolveSectionCode` may return `""` when there is no known mapping.

An October workbook job (November 2026 data) required **only** filling
column `K` from GESN in `J`: 7,206 rows inspected, 7,189 mapped,
17 unresolved in that run, including GESN19/GESN29/GESNm21 examples; 9
historical manual corrections were noted during analysis. These counts are
not a guarantee for a future workbook. **Do not change source values,
formulas, styles, sheets, hidden metadata or other columns** when asked
for `K`-only; compare all untouched cells against the original and do not
create a «Проверка ЕКР» sheet unless specifically requested.

Section `01` must remain text `"01"`, not numeric `1`. A blank uncertain
code is better than a fabricated one. Record unresolved cases separately
in your response if the file itself must remain otherwise identical.

## 6. Spreadsheet preservation checklist

For `.xlsx`/`.xlsm` outputs:

- Use the original workbook as the template, not `DataFrame.to_excel` when
  full structural preservation is required.
- Preserve cell values and formulas outside specified target, merged cells,
  sheets, named ranges, print areas, dimensions, macros and drawings.
- The project already has special OOXML-preservation logic
  (`core/ooxml_preservation.py`, `core/excel_writer.py`); do not casually
  replace it with a generic `openpyxl.save` round trip when it would discard
  unsupported parts.
- Compare original and result at cell and relevant package-part levels;
  changing just a value in one column is a different task from redesign.
- Use short ASCII deliverable names (`EKR_Nov_2026.xlsx`) to avoid client
  issues with Cyrillic/long download filenames.

## 7. Minimum regression matrix

| Change | At least verify |
| --- | --- |
| RNMC importer | filename-only duplicates incl. cross-folder; legacy statuses; metadata; no-data/rejected/excluded rows; source lineage. |
| TKP parser | winner identity, participant column offsets, cached/formula cell values, VAT/price per unit/line total, task/source row, non-standard headers. |
| TKP matching | unit mismatch, scale, diameter/material, demolition vs installation, layers, selected price/score, shadow remains read-only. |
| Database migration | schema version, existing rows unchanged, FK/integrity, other tables, backup/restore, idempotent re-run. |
| EKR/Excel | GESN prefixes/demolition, unknown codes, leading zeros, untouched cells/formulas/OOXML. |
| Full application | `python -m pytest -q` before/after and a representative end-to-end WA output. |

Above all: **an importer can pass its unit tests and still be unsafe for a
new original workbook shape.** Establish a verified real-file fixture first.
