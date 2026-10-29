"""Read-only browser evaluation of Estimate AI-generated Excel price metrics.

The workbook and stored processing rows retain original formulas. Only known
writer formulas are evaluated; this module does not execute arbitrary Excel.
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from math import isfinite

from openpyxl.utils import column_index_from_string

_CELL = r"\$?([A-Z]{1,3})\$?(\d+)"
_PRODUCT = re.compile(rf"\b({_CELL})\s*\*\s*({_CELL})", re.IGNORECASE)
_DEVIATION = re.compile(rf"\(({_CELL})\s*-\s*({_CELL})\)\s*/\s*({_CELL})", re.IGNORECASE)
_EXTREMA = re.compile(rf"\b(MIN|MAX)\(\s*({_CELL})\s*:\s*({_CELL})\s*\)", re.IGNORECASE)


def _number(value: object) -> float | None:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    result = float(value)
    return result if isfinite(result) else None


def _cell_index(letter: str, row: str, excel_row: int) -> int | None:
    return column_index_from_string(letter.upper()) - 1 if int(row) == excel_row else None


def _get(values: Sequence[object], letter: str, row: str, excel_row: int) -> float | None:
    index = _cell_index(letter, row, excel_row)
    return _number(values[index]) if index is not None and index < len(values) else None


def _range(formula: str, excel_row: int) -> tuple[int, int] | None:
    match = _EXTREMA.search(formula)
    if match is None:
        return None
    a = _cell_index(match.group(3), match.group(4), excel_row)
    b = _cell_index(match.group(6), match.group(7), excel_row)
    return (a, b) if a is not None and b is not None and a <= b else None


def _derived(formula: str, values: Sequence[object], excel_row: int) -> float | None:
    """Evaluate only guarded quantity, deviation, minimum and maximum formulas."""
    if not formula.startswith("="):
        return None
    if "ISNUMBER(" in formula.upper() and "*" in formula:
        m = _PRODUCT.search(formula)
        if m:
            a = _get(values, m.group(2), m.group(3), excel_row)
            b = _get(values, m.group(5), m.group(6), excel_row)
            return a * b if a is not None and b is not None else None
    if "ISNUMBER(" in formula.upper() and "/" in formula:
        m = _DEVIATION.search(formula)
        if m and m.group(5).upper() == m.group(8).upper() and m.group(6) == m.group(9):
            avg = _get(values, m.group(2), m.group(3), excel_row)
            orig = _get(values, m.group(5), m.group(6), excel_row)
            if avg is not None and orig is not None and orig != 0:
                return (avg - orig) / orig
    if "COUNTIF(" in formula.upper():
        match = _EXTREMA.search(formula)
        columns = _range(formula, excel_row)
        if match and columns:
            prices = [n for raw in values[columns[0]:columns[1] + 1]
                      if (n := _number(raw)) is not None and n > 0]
            if prices:
                return min(prices) if match.group(1).upper() == "MIN" else max(prices)
    return None


def resolve_generated_preview_values(
    columns: Sequence[Mapping[str, object]],
    rows: Sequence[object],
    values_by_id: Mapping[int, list[object]],
    overrides: Sequence[object] = (),
) -> dict[int, list[object]]:
    """Render derived prices as numbers, without overwriting stored formulas.

    Saved overrides take precedence. Editing analog prices also refreshes the
    average unit price before dependent I (total) and Q (deviation) metrics.
    """
    resolved = {int(row_id): list(values) for row_id, values in values_by_id.items()}
    original = {int(row.id): list(row.values) for row in rows}
    manual: dict[int, set[int]] = {}
    for override in overrides:
        row_id, index = int(override.row_id), int(override.column_index) - 1
        if row_id in resolved and 0 <= index < len(resolved[row_id]):
            resolved[row_id][index] = override.current_value
            manual.setdefault(row_id, set()).add(index)

    average_index = next((int(c["index"]) - 1 for c in columns if c.get("kind") == "average"), None)
    metric_indices = {
        int(c["index"]) - 1
        for c in columns
        if any(term in str(c.get("label", "")).lower() for term in (
            "стоимость средняя", "отклонение средней", "мин. цена аналога", "макс. цена аналога"
        ))
    }
    for row in rows:
        row_id = int(row.id)
        values = resolved.get(row_id)
        formulas = original.get(row_id)
        if values is None or formulas is None:
            continue
        excel_row = int(row.excel_row_number)
        edited = manual.get(row_id, set())
        if average_index is not None and average_index not in edited:
            for raw in formulas:
                if not isinstance(raw, str) or "COUNTIF(" not in raw.upper():
                    continue
                span = _range(raw, excel_row)
                if span is None or not any(span[0] <= changed <= span[1] for changed in edited):
                    continue
                if 0 <= average_index < len(values):
                    base_index = next((int(c["index"]) - 1 for c in columns
                        if "цена единицы работ" in str(c.get("label", "")).lower()), None)
                    base = _number(values[base_index]) if base_index is not None else None
                    if base is not None:
                        prices = [_number(x) for x in values[span[0]:span[1] + 1]]
                        numeric = [x for x in [base, *prices] if x is not None]
                        values[average_index] = max(base, sum(numeric) / len(numeric))
                break
        for col, formula in enumerate(formulas):
            if col not in metric_indices or col in edited or not isinstance(formula, str) or not formula.startswith("="):
                continue
            if ("COUNTIF(" in formula.upper() and _EXTREMA.search(formula)) or (
                "ISNUMBER(" in formula.upper() and (_PRODUCT.search(formula) or _DEVIATION.search(formula))
            ):
                values[col] = _derived(formula, values, excel_row)
    return resolved
