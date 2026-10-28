"""Regression tests for price metrics in generated WA files."""

from types import SimpleNamespace

from openpyxl import Workbook

from core.excel_writer import (
    GlobalAnalogPlan,
    WriterColumns,
    _plan_output_columns,
    _write_metric_formulas,
    _write_metric_headers,
)


def _columns():
    layout = WriterColumns(
        base_price=6, average=7, code=14, code_kr=15,
        section=16, analog_start=20, header_row=7,
    )
    return _plan_output_columns(layout, SimpleNamespace(column=7), 20)


def _plan(last):
    return GlobalAnalogPlan(
        by_key={}, tkp_winner_by_key={}, tkp_reserve_by_key={},
        columns=(), last_column=last,
    )


def test_metrics_after_section_and_before_analogs():
    cols = _columns()
    assert (cols.deviation, cols.min_price, cols.max_price, cols.analog_start) == (17,18,19,20)
    sheet = Workbook().active
    _write_metric_headers(sheet, 7, cols)
    assert "Отклонение" in sheet["Q7"].value
    assert "Мин." in sheet["R7"].value
    assert "Макс." in sheet["S7"].value


def test_dynamic_formulas_cover_rnmc_and_paired_tkp_columns():
    sheet = Workbook().active
    _write_metric_formulas(sheet, 9, _columns(), _plan(23))
    assert sheet["Q9"].value == '=IFERROR(IF(OR(NOT(ISNUMBER(F9)),F9=0),"",(G9-F9)/F9),"")'
    assert sheet["Q9"].number_format == "+0.00%;-0.00%;0.00%"
    assert sheet["R9"].value == '=IFERROR(AGGREGATE(15,6,T9:W9/(T9:W9>0),1),"")'
    assert sheet["S9"].value == '=IF(COUNTIF(T9:W9,">0")=0,"",MAX(T9:W9))'


def test_price_metrics_blank_without_candidates():
    sheet = Workbook().active
    _write_metric_formulas(sheet, 9, _columns(), _plan(19))
    assert sheet["R9"].value is None
    assert sheet["S9"].value is None
