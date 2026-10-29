# WA Excel export: price comparison columns

Status: development version ready for manual acceptance on user's laptop; production remains unchanged.

Immediately after Код раздела, output deviation %, minimum analog price and maximum analog price.
Deviation = (average unit price - original unit price) / original unit price; zero/missing original -> blank.

Both extrema are computed by regular Excel formulas on the same RNMC and paired TKP price columns:

    =IF(COUNTIF(T36:BM36,">0")=0,"",MIN(T36:BM36))
    =IF(COUNTIF(T36:BM36,">0")=0,"",MAX(T36:BM36))

Matching/import rules ensure positive candidate prices. MIN/MAX ignore blank and text cells.
Both extrema remain blank if no positive analog exists. As agreed, do not use AGGREGATE or `_xlfn.` functions.
The RNMC plus TKP winner/reserve columns contain numeric prices, not task identifiers or names.
Original and average prices and calculated summary columns are excluded.

The 2026-10-29 request added average work total as a *new feature*: detected quantity multiplied by average unit price (e.g. I = E * G).
Legacy average unit-price rule MAX(original, AVERAGE(original,candidates)) is unchanged.
Excel recalculates the generated formulas upon opening; libraries may show no cached values.

Regression: check standard formula serialization, both extrema use identical source ranges, blank/no candidates,
and deletion of candidate prices. Then manually process the original WA workbook on the laptop.
For row 36 verify R36 shows the lowest price, and deleting that candidate selects the next lowest.
Do not deploy any code to the ololosh production server before explicit manual acceptance.
