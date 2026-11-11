"""GSFX normalization uses analogs only, never TKP, and preserves totals."""
from decimal import Decimal
import sqlite3
from pathlib import Path
from app.services.grand_normalization import load_catalog_index, normalize_rows, parse_unit

def _catalog_db(path: Path):
    db = sqlite3.connect(path)
    db.executescript('''
        CREATE TABLE catalog_sources(id INTEGER PRIMARY KEY, name TEXT);
        CREATE TABLE catalog_items(source_id INTEGER, code TEXT, work_name TEXT, unit TEXT, price REAL);
        CREATE TABLE tkp_items(item_code TEXT, item_name TEXT, unit TEXT);
        INSERT INTO catalog_sources VALUES (1, 'main'), (2, 'other');
        INSERT INTO catalog_items VALUES
          (1, 'G-10', 'Прокладка кабеля диаметром 12 мм', 'м', 100),
          (1, 'G-20', 'Установка оборудования', 'шт', 20),
          (1, 'G-30', 'Монтаж трубы', '100 м', 2000),
          (1, 'G-40', 'Монтаж кабельных конструкций', '100 м', 200),
          (1, 'G-40', 'Монтаж кабельных конструкций', 'м', 2),
          (1, 'G-50', 'Погрузка в автотранспортное средство', 'т', 3),
          (2, 'G-10', 'Монтаж подстанции', 'компл', 100000);
        INSERT INTO tkp_items VALUES ('G-10','Случайная работа','км');
    ''')
    db.commit()
    db.close()

def _row(code, name, unit, qty, price=''):
    return dict(code=code, name=name, unit=unit, quantity=qty, price=price)

def test_units_from_analogs_preserve_quantity_price_and_total(tmp_path):
    path = tmp_path / 'catalog.db'
    _catalog_db(path)
    before = path.read_bytes()
    index = load_catalog_index(path)
    assert index.count == 6
    rows = [_row('G-10', 'Прокладка кабеля диаметром 12 мм', '100 м', '2.5', '12000'),
            _row('G-20', 'Установка оборудования', '1 шт.', '5', ''),
            _row('G-50', 'Погрузка в автотранспортное средство', '1т груза', '1225', '')]
    total_before = Decimal(rows[0]['quantity']) * Decimal(rows[0]['price'])
    summary = normalize_rows(rows, index)
    assert summary.changed == 3
    assert (rows[0]['unit'], rows[0]['quantity'], rows[0]['price']) == ('м', '250', '120')
    assert Decimal(rows[0]['quantity']) * Decimal(rows[0]['price']) == total_before
    assert (rows[1]['unit'], rows[1]['quantity']) == ('шт', '5')
    assert (rows[2]['unit'], rows[2]['quantity']) == ('т', '1225')
    assert rows[0]['original_unit'] == '100 м'
    assert rows[0]['original_quantity'] == '2.5'
    assert path.read_bytes() == before

def test_scaled_analog_is_retained_and_mixed_scale_is_reviewed(tmp_path):
    path = tmp_path / 'catalog.db'
    _catalog_db(path)
    index = load_catalog_index(path)
    rows = [_row('G-30', 'Монтаж трубы', '100 м', '2', '3000'),
            _row('G-40', 'Монтаж кабельных конструкций', '100 м', '2', '3000')]
    summary = normalize_rows(rows, index)
    assert summary.review == 1
    assert rows[0]['unit'] == '100 м' and rows[0]['price'] == '3000'
    assert rows[1]['normalization_status'] == 'review'
    assert rows[1]['unit'] == '100 м' and rows[1]['price'] == '3000'
    assert 'разные масштабы' in rows[1]['normalization_note']

def test_no_inferred_catalog_price_or_unsupported_conversion(tmp_path):
    path = tmp_path / 'catalog.db'
    _catalog_db(path)
    index = load_catalog_index(path)
    rows = [_row('G-10', 'Прокладка кабеля диаметром 99 мм', '100 м', '2'),
            _row('MISSING', 'Погрузка', '100 т', '3'),
            _row('G-10', 'Прокладка кабеля диаметром 12 мм', '100 м2', '2')]
    normalize_rows(rows, index)
    assert rows[0]['name'].endswith('99 мм')
    assert rows[0]['price'] == ''
    assert rows[1]['unit'] == '100 т' and rows[1]['price'] == ''
    assert rows[2]['normalization_status'] == 'review'
    assert rows[2]['unit'] == '100 м2'

def test_normalization_shortens_only_near_identical_names(tmp_path):
    path = tmp_path / 'catalog.db'
    _catalog_db(path)
    index = load_catalog_index(path)
    rows = [_row('G-10', 'Прокладка кабеля диаметром 12 мм.', 'м', '2'),
            _row('G-10', 'Прокладка кабеля диаметром 13 мм', 'м', '3')]
    normalize_rows(rows, index)
    assert rows[0]['name'] == 'Прокладка кабеля диаметром 12 мм'
    assert rows[1]['name'] == 'Прокладка кабеля диаметром 13 мм'

def test_unit_parser_known_scale_and_rejects_unfamiliar():
    assert parse_unit('100 м²').scale == Decimal(100)
    assert parse_unit('1000 м³').base == 'м3'
    assert parse_unit('1т груза').base == 'т'
    assert parse_unit('м3').scale == Decimal(1)
    assert parse_unit('100 пар') is None
    assert parse_unit('0.5 м').scale == Decimal('0.5')
