"""Conservative GrandSmeta -> OS normalization using only the analog catalog.

The source is catalog_items of catalog_sources.name='main'. TKP tables and
catalog prices are never consulted. The database is opened read-only.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from pathlib import Path
import re
import sqlite3
from urllib.parse import quote

from core.normalize import HasDemontazh, NormCode


@dataclass(frozen=True)
class Unit:
    text: str
    base: str
    scale: Decimal


@dataclass(frozen=True)
class CatalogExample:
    name: str
    unit: str
    parsed_unit: Unit


@dataclass(frozen=True)
class CatalogIndex:
    by_code: dict[str, tuple[CatalogExample, ...]]
    count: int


@dataclass(frozen=True)
class NormalizationSummary:
    available: bool
    catalog_items: int
    changed: int
    review: int
    untouched: int
    reason: str = ""


_UNIT_ALIASES = {
    'шт': 'шт', 'штука': 'шт', 'штук': 'шт', 'штуки': 'шт',
    'м': 'м', 'погм': 'м', 'погонныйм': 'м', 'погонныхм': 'м',
    'м2': 'м2', 'квм': 'м2', 'квадратныйм': 'м2', 'квадратныхм': 'м2',
    'м3': 'м3', 'кубм': 'м3', 'кубическийм': 'м3', 'кубическихм': 'м3',
    'т': 'т', 'тн': 'т', 'тонна': 'т', 'тонны': 'т', 'тонн': 'т',
    'тгруза': 'т', 'тоннагруза': 'т', 'тоннгруза': 'т',
    'кг': 'кг', 'компл': 'компл', 'комплект': 'компл', 'комплекта': 'компл',
    'км': 'км', 'час': 'час', 'ч': 'час', 'машчас': 'машчас',
    'челчас': 'челчас', 'мес': 'мес', 'месяц': 'мес',
}


def parse_unit(value: str) -> Unit | None:
    """Recognize known multiplier without confusing m2/m3 with multipliers."""
    text = str(value or '').strip()
    compact = re.sub(r'\s+', '', text.casefold().replace('ё', 'е').replace('\xa0', ' '))
    compact = compact.replace('²', '2').replace('³', '3').replace('^', '')
    match = re.fullmatch(r'(?:(\d+(?:[.,]\d+)?)\s*)?(.+)', compact)
    if match is None:
        return None
    raw_multiplier, raw_base = match.groups()
    raw_base = raw_base.replace('.', '').replace(',', '')
    try:
        scale = Decimal(raw_multiplier.replace(',', '.')) if raw_multiplier else Decimal('1')
    except InvalidOperation:
        return None
    if scale <= 0 or scale > 1_000_000:
        return None
    base = _UNIT_ALIASES.get(raw_base)
    if not base:
        return None
    return Unit(text, base, scale)


def _name_words(text: str) -> tuple[str, ...]:
    return tuple(re.findall(r'[а-яa-z]+|\d+(?:[.,]\d+)?', text.casefold().replace('ё', 'е')))


def _name_score(source: str, candidate: str) -> float:
    a = _name_words(source)
    b = _name_words(candidate)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, ' '.join(a), ' '.join(b)).ratio()


def _same_technical_values(source: str, candidate: str) -> bool:
    """Never discard source numeric diameter, dimensions, mass or model numbers."""
    def numbers(text: str) -> set[str]:
        return set(re.findall(r'\d+(?:[.,]\d+)?', text))
    return numbers(source).issubset(numbers(candidate))


def _decimal(value: object) -> Decimal | None:
    try:
        raw = str(value).replace(' ', '').replace('\xa0', '').replace(',', '.').strip()
        number = Decimal(raw)
        return number if number.is_finite() else None
    except (InvalidOperation, ValueError):
        return None


def _decimal_text(value: Decimal) -> str:
    return format(value.normalize(), 'f')


def load_catalog_index(path: str | Path, *, source_name: str = 'main') -> CatalogIndex:
    """Read analog source only, excluding TKP and all catalog price fields."""
    file = Path(path)
    if not file.is_file():
        return CatalogIndex({}, 0)
    uri = 'file:' + quote(str(file.resolve()), safe='/:') + '?mode=ro'
    connection = sqlite3.connect(uri, uri=True)
    try:
        data = connection.execute(
            "SELECT ci.code, ci.work_name, ci.unit "
            "FROM catalog_items ci "
            "JOIN catalog_sources cs ON cs.id=ci.source_id "
            "WHERE cs.name=? AND ci.code<>'' AND ci.unit<>'' AND ci.work_name<>''",
            (source_name,),
        )
        by_code: dict[str, list[CatalogExample]] = defaultdict(list)
        count = 0
        for code, name, unit in data:
            count += 1
            parsed = parse_unit(str(unit))
            if parsed is not None:
                by_code[NormCode(code)].append(CatalogExample(str(name), str(unit), parsed))
        return CatalogIndex({key: tuple(value) for key, value in by_code.items()}, count)
    finally:
        connection.close()


def normalize_rows(rows: list[dict], catalog: CatalogIndex) -> NormalizationSummary:
    """Normalize using unambiguous code, name and unit matches only.

    If the catalog uses mixed unit price bases for one code, do not guess.
    When a verified scale differs, convert quantity and price together.
    """
    changed = review = untouched = 0
    for row in rows:
        for key in ('name', 'unit', 'quantity', 'price'):
            row.setdefault('original_' + key, row.get(key, ''))
        row['analog_suggestion'] = ''
        row['normalization_note'] = ''
        row['normalization_status'] = 'no_match'
        source_name = row.get('name', '')
        source_unit = parse_unit(row.get('unit', ''))
        code = NormCode(row.get('code', ''))
        candidates = catalog.by_code.get(code, ())
        if not candidates:
            row['normalization_note'] = 'Нет позиции с таким шифром в базе аналогов'
            untouched += 1
            continue
        if source_unit is None:
            row['normalization_status'] = 'review'
            row['normalization_note'] = 'Единица измерения неизвестна — проверьте вручную'
            review += 1
            continue
        comparable = [item for item in candidates
                      if item.parsed_unit.base == source_unit.base
                      and HasDemontazh(item.name) == HasDemontazh(source_name)]
        if not comparable:
            row['normalization_status'] = 'review'
            row['normalization_note'] = 'У аналогов с этим шифром другая единица или вид работ'
            review += 1
            continue
        ranked = sorted(((_name_score(source_name, item.name), item) for item in comparable),
                        key=lambda item: (-item[0], len(item[1].name)))
        top_score, best = ranked[0]
        if top_score < .72:
            row['normalization_status'] = 'review'
            row['normalization_note'] = 'Найден шифр, но наименование заметно отличается'
            row['analog_suggestion'] = best.name
            review += 1
            continue
        # Matching groups by normalized code/base unit rather than unit multiplier.
        # Hence ALL catalog rows with a compatible key must agree on price basis.
        scales = {item.parsed_unit.scale for item in comparable}
        row['analog_suggestion'] = best.name
        did_change = False
        notes: list[str] = []
        if len(scales) == 1:
            target_unit = best.unit
            target_scale = best.parsed_unit.scale
            factor = source_unit.scale / target_scale
            old_qty = _decimal(row.get('quantity', ''))
            old_price = _decimal(row.get('price', ''))
            if factor != 1 and old_qty is None:
                notes.append('Не удалось пересчитать количество; проверьте вручную')
            else:
                if row['unit'] != target_unit:
                    row['unit'] = target_unit
                    did_change = True
                    notes.append('Единица по базе аналогов')
                if factor != 1:
                    row['quantity'] = _decimal_text(old_qty * factor)
                    if old_price is not None:
                        row['price'] = _decimal_text(old_price / factor)
                    did_change = True
                    notes.append(f'Масштаб: ×{_decimal_text(factor)} к количеству, ÷{_decimal_text(factor)} к цене')
        else:
            notes.append('В базе аналогов разные масштабы единицы — требуется проверка')
        if (top_score >= .90 and _same_technical_values(source_name, best.name)
                and len(best.name) < len(source_name)):
            near_names = {item.name for score, item in ranked if score >= top_score - .02}
            if len(near_names) == 1:
                row['name'] = best.name
                did_change = True
                notes.append('Наименование по базе аналогов')
        if len(scales) > 1:
            row['normalization_status'] = 'review'
            review += 1
        elif notes and any('Не удалось' in item for item in notes):
            row['normalization_status'] = 'review'
            review += 1
        elif did_change:
            row['normalization_status'] = 'changed'
            changed += 1
        else:
            row['normalization_status'] = 'matched'
            untouched += 1
        row['normalization_note'] = '; '.join(notes) or 'Совпадение с базой аналогов'
    return NormalizationSummary(bool(catalog.count), catalog.count, changed, review, untouched)
