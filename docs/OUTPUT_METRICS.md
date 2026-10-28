# WA export: dynamic price metrics (2026-10-28)

Status: ready for testing on a laptop; production rollout is pending manual acceptance.

Immediately after the **Код раздела** column, the generated WA workbook contains:

1. **Отклонение средней от исходной, %**: `(average - original) / original`, displayed with its sign; blank for zero/missing/non-numeric original.
2. **Мин. цена аналога, руб. без НДС**: smallest positive price using Excel AGGREGATE formula.
3. **Макс. цена аналога, руб. без НДС**: largest positive price using COUNTIF/MAX formula.

The current web exporter places RNMC prices and optional paired TKP winner/reserve participant prices in the same analog block. The dynamic extrema include all of these candidate prices, but not the original, average, or new summary columns. Deleting or editing a candidate price recalculates the extrema in Excel. When no positive prices exist, extrema are blank.

The legacy average formula and deterministic matching logic remain unchanged. Workbook formulas are recalculated upon opening with Microsoft Excel (AGGREGATE requires Excel 2010+); openpyxl does not evaluate cached values.

Before releasing to production: run `python -m pytest -q`, check GitHub CI, export files with and without paired TKP prices, delete minimum/maximum values and confirm formulas update, verify reprocessing and nonstandard templates. Do not alter production SQLite data for this feature.
