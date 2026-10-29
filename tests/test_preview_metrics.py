from types import SimpleNamespace

from app.services.preview_metrics import resolve_generated_preview_values

COLUMNS = [
    {"index": 5, "label": "Кол-во", "kind": "source"},
    {"index": 6, "label": "Цена единицы работ, руб. без НДС", "kind": "source"},
    {"index": 7, "label": "Цена средняя", "kind": "average"},
    {"index": 9, "label": "Стоимость средняя", "kind": "source"},
    {"index": 17, "label": "Отклонение средней от исходной, %", "kind": "source"},
    {"index": 18, "label": "Мин. цена аналога", "kind": "source"},
    {"index": 19, "label": "Макс. цена аналога", "kind": "source"},
]


def make_row():
    cells = [None] * 23
    cells[4:7] = [3, 100, 125]
    cells[8] = '=IFERROR(IF(AND(ISNUMBER(E36),ISNUMBER(G36)),E36*G36,""),"")'
    cells[16] = '=IFERROR(IF(OR(NOT(ISNUMBER(F36)),F36=0),"",(G36-F36)/F36),"")'
    cells[17] = '=IF(COUNTIF(T36:W36,">0")=0,"",MIN(T36:W36))'
    cells[18] = '=IF(COUNTIF(T36:W36,">0")=0,"",MAX(T36:W36))'
    cells[19:23] = [100, 150, 180, None]
    return SimpleNamespace(id=42, excel_row_number=36, values=cells)


def calculated(overrides=(), row=None):
    row = row or make_row()
    display = resolve_generated_preview_values(COLUMNS, [row], {42: list(row.values)}, overrides)
    return row, display[42]


def test_browser_displays_numbers_not_formula_text():
    original, cells = calculated()
    assert cells[8] == 375
    assert cells[16] == .25
    assert cells[17] == 100
    assert cells[18] == 180
    assert original.values[17].startswith('=IF(')


def test_deleting_min_and_max_updates_preview_without_modifying_formula():
    row, cells = calculated([
        SimpleNamespace(row_id=42, column_index=20, current_value=None),
        SimpleNamespace(row_id=42, column_index=22, current_value=None),
        SimpleNamespace(row_id=42, column_index=21, current_value=220),
    ])
    assert cells[17] == 220 and cells[18] == 220
    assert cells[6] == 160
    assert cells[8] == 480
    assert cells[16] == .6
    assert row.values[17].startswith('=IF(')


def test_no_analog_prices_displays_empty_extrema():
    row = make_row()
    row.values[19:23] = [None, None, None, None]
    _, cells = calculated(row=row)
    assert cells[17] is None and cells[18] is None
    assert cells[6] == 125


def test_manual_metric_override_is_respected():
    _, cells = calculated([SimpleNamespace(row_id=42, column_index=18, current_value=45)])
    assert cells[17] == 45 and cells[18] == 180


def test_malformed_other_excel_formula_is_untouched():
    row = make_row()
    row.values[9] = '=SOMETHING(1)'
    _, cells = calculated(row=row)
    assert cells[9] == '=SOMETHING(1)'
