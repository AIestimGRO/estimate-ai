# WA export: dynamic price metrics (2026-10-28)

Status: ready for testing on a laptop; production rollout is pending manual acceptance.

Immediately after the **Код раздела** column, the generated WA workbook contains:

1. **Отклонение средней от исходной, %**: `(average - original) / original`, displayed with its sign; blank for zero/missing/non-numeric original.
2. **Мин. цена аналога, руб. без НДС**: smallest positive price using Excel AGGREGATE formula.
3. **Макс. цена аналога, руб. без НДС**: largest positive price using COUNTIF/MAX formula.

The current web exporter places RNMC prices and optional paired TKP winner/reserve participant prices in the same analog block. The dynamic extrema include all of these candidate prices, but not the original, average, or new summary columns. Deleting or editing a candidate price recalculates the extrema in Excel. When no positive prices exist, extrema are blank.

The legacy average formula and deterministic matching logic remain unchanged. Workbook formulas are recalculated upon opening with Microsoft Excel (AGGREGATE requires Excel 2010+); openpyxl does not evaluate cached values.

Before releasing to production: run `python -m pytest -q`, check GitHub CI, export files with and without paired TKP prices, delete minimum/maximum values and confirm formulas update, verify reprocessing and nonstandard templates. Do not alter production SQLite data for this feature.

## 2026-10-29: findings from reviewed WA workbook and new total

- In the reviewed workbook, the average unit price is in G and quantity is in E; the *Стоимость средняя* column I was blank. The exporter now writes a live Excel formula `I = G * E` for every processed work row, so existing I totals update automatically.
- The quantity and total columns are found by their headers; other templates without these headers are left untouched.
- Output metric headers wrap and have a reasonable width without changing the width of the analog columns.
- Formula cells have no cached values until recalculated by Excel. The workbook requests an automatic full recalc.
- The existing mean-price rule `MAX(original, AVERAGE(original, candidates))` is intentionally unchanged. Because the average is floored at the original, the signed percentage may be 0 when every found candidate is cheaper.
