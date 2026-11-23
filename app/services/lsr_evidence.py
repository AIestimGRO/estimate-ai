"""Compare a specialist's OS with its GSFX before retaining training evidence.

This is an evidence-quality gate, not an analog-matching or price-learning
engine. In particular, an exact code/quantity match cannot prove a price.
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook

from app.services.grand_normalization import _name_score, _same_technical_values, parse_unit
from app.services.grand_rim_calculator import calculate_gsfx_prices
from app.services.read_estimate import load_estimate
from core.grand_smeta import read_gsfx
from core.normalize import NormCode


@dataclass(frozen=True)
class EvidenceReport:
    quality: str  # candidate or review; neither is approved training data
    os_count: int
    lsr_count: int
    matched: int
    checked_prices: int
    issues: tuple[str, ...]
    rows: tuple[dict, ...]

    @property
    def message(self) -> str:
        label = ('Структурная сверка пройдена; требуется утверждение администратором'
                 if self.quality == 'candidate' else
                 'Есть расхождения; материалы останутся на проверке')
        details = '; '.join(self.issues[:3])
        return (f'ЛСР ↔ ОС: {self.matched} из {self.os_count} строк сопоставлено; '
                f'проверено цен: {self.checked_prices}. {label}.'
                + (f' {details}' if details else ''))


def _num(value: object) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    try:
        amount = Decimal(str(value).strip().replace('\u00a0', '').replace(' ', '').replace(',', '.'))
        return amount if amount.is_finite() else None
    except (InvalidOperation, ValueError):
        return None


def _column_for(sheet, header: int, patterns: tuple[str, ...]) -> int | None:
    if header < 1:
        return None
    for col in range(1, min(sheet.max_column, 40) + 1):
        label = str(sheet.cell(header, col).value or '').casefold().replace('ё', 'е')
        label = ' '.join(label.replace('\n', ' ').split())
        if any(pattern in label for pattern in patterns):
            return col
    return None


def evaluate_lsr_pair(os_path: Path, gsfx_path: Path, *, sheet_title: str) -> EvidenceReport:
    """Fail closed on missing structural fields or unverifiable prices.

    The workbook is read-only; the original spreadsheet and analog database
    never receive writes during this comparison.
    """
    estimate = load_estimate(os_path, selected_sheet_title=sheet_title)
    source = read_gsfx(gsfx_path)
    calculated = calculate_gsfx_prices(gsfx_path)
    sources = {}
    duplicates: set[tuple[str, str]] = set()
    for position, price in zip(source.positions, calculated):
        key = (position.number.strip(), NormCode(position.code))
        if key in sources:
            duplicates.add(key)
        sources[key] = (position, price)
    issues: list[str] = []
    if duplicates:
        issues.append('В ЛСР повторяются номера позиций с одинаковыми шифрами')

    book = load_workbook(os_path, read_only=True, data_only=True)
    matched = checked_prices = 0
    used: set[tuple[str, str]] = set()
    snapshots: list[dict] = []
    try:
        sheet = book[sheet_title]
        number_col = _column_for(sheet, estimate.header_row, ('№ п/п в смете', '№ позиции в смете'))
        qty_col = _column_for(sheet, estimate.header_row, ('кол-во', 'количество', 'объем работ', 'объём работ'))
        worker_col = _column_for(sheet, estimate.header_row, ('тз на ед.', 'тз на единицу'))
        machine_col = _column_for(sheet, estimate.header_row, ('тзм на ед.', 'тзм на единицу'))
        if number_col is None or qty_col is None:
            issues.append('В ОС не найдены заголовки номера позиции или количества')
        for excel_row, os_item in estimate.positioned_rows:
            number = str(sheet.cell(excel_row, number_col).value or '').strip() if number_col else ''
            quantity = _num(sheet.cell(excel_row, qty_col).value) if qty_col else None
            code = NormCode(os_item.code)
            key = (number, code)
            item_issues = []
            original, calculated_price = sources.get(key, (None, None))
            if original is None or key in used or key in duplicates:
                item_issues.append('Нет однозначной позиции ЛСР с тем же номером и шифром')
            else:
                used.add(key)
                matched += 1
                left = parse_unit(original.unit)
                right = parse_unit(str(os_item.unit))
                if not left or not right or left.base != right.base:
                    item_issues.append('Несовместимые или неизвестные единицы измерения')
                elif original.quantity is None or quantity is None or quantity <= 0:
                    item_issues.append('Не удалось подтвердить количество')
                else:
                    in_source_units = original.quantity * left.scale
                    in_os_units = quantity * right.scale
                    if abs(in_source_units - in_os_units) > max(Decimal('0.00001'), abs(in_source_units)*Decimal('0.0001')):
                        item_issues.append('Не совпадает физическое количество после перевода единиц')
                    if calculated_price is not None:
                        for col, expected, title in (
                            (worker_col, calculated_price.labor_per_norm_unit, 'ТЗ'),
                            (machine_col, calculated_price.machinist_labor_per_norm_unit, 'ТЗм'),
                        ):
                            if col is None:
                                continue
                            actual = _num(sheet.cell(excel_row, col).value)
                            if actual is None:
                                continue
                            scaled = expected * right.scale / left.scale
                            if abs(actual - scaled) > max(Decimal('0.02'), abs(scaled) * Decimal('0.01')):
                                item_issues.append(f'{title} на единицу отличается от расчёта ЛСР')
                if (_name_score(original.name, str(os_item.work_name)) < 0.72 or
                        not _same_technical_values(original.name, str(os_item.work_name))):
                    item_issues.append('Название или технические параметры работы отличаются')
                source_total = (calculated_price.total if calculated_price else
                                original.unit_work_price * original.quantity
                                if original.unit_work_price is not None and original.quantity is not None
                                else None)
                os_price = _num(os_item.base_price)
                if source_total is not None and quantity is not None and os_price is not None and os_price > 0:
                    checked_prices += 1
                    if abs(quantity * os_price - source_total) > max(Decimal('1'), abs(source_total) * Decimal('0.005')):
                        item_issues.append('Цена × количество не соответствует проверенной стоимости ЛСР')
                elif os_price is None or os_price <= 0:
                    item_issues.append('Цена ОС отсутствует или некорректна')
            if item_issues:
                issues.append(f'Строка ОС {excel_row}: ' + ', '.join(item_issues))
            snapshots.append({
                'excel_row': excel_row, 'number': number, 'code': code,
                'name': str(os_item.work_name or ''), 'unit': str(os_item.unit or ''),
                'quantity': str(quantity) if quantity is not None else '',
                'price': str(os_item.base_price or ''), 'issues': item_issues,
            })
    finally:
        book.close()
    missing = len(source.positions) - len(used)
    if missing:
        issues.append(f'В ОС отсутствуют {missing} из {len(source.positions)} позиций ЛСР')
    if checked_prices < matched:
        issues.append(f'Не все цены подтверждены GSFX: {checked_prices} из {matched}')
    quality = ('candidate' if not issues and matched and checked_prices == matched else 'review')
    return EvidenceReport(quality, len(snapshots), len(source.positions),
                          matched, checked_prices, tuple(issues[:150]), tuple(snapshots))
