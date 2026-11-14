"""Verified prices from a filled GrandSmeta RIM report.

Calculated prices are produced by GrandSmeta's template engine. Never invent
a work price by summing GSFX resources or borrowing prices from analogs/TKP.
"""
from __future__ import annotations
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from zipfile import BadZipFile
from openpyxl import load_workbook
from core.normalize import NormCode

class RimReportError(ValueError):
    """The Excel file is not a completed RIM estimate report."""

@dataclass(frozen=True)
class RimWork:
    number: str
    code: str
    name: str
    unit: str
    quantity: Decimal
    unit_price: Decimal
    total: Decimal
    labor_unit: Decimal | None
    machinist_labor_unit: Decimal | None
    labor_total: Decimal | None
    machinist_labor_total: Decimal | None
    average_price: Decimal | None
    average_total: Decimal | None
    section_code: str
    sheet_row: int

def _decimal(value: object) -> Decimal | None:
    if value is None or isinstance(value, bool):
        return None
    try:
        parsed = Decimal(str(value).replace('\u00a0', '').replace(' ', '').replace(',', '.'))
    except (InvalidOperation, ValueError):
        return None
    return parsed if parsed.is_finite() else None

def read_filled_rim_report(path: str | Path) -> tuple[RimWork, ...]:
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except (OSError, ValueError, KeyError, TypeError, BadZipFile) as exc:
        raise RimReportError('Не удалось открыть заполненный Excel-отчёт РИМ') from exc
    try:
        worksheet = workbook.active
        header = next((i for i, cells in enumerate(worksheet.iter_rows(max_col=15, values_only=True), 1)
                       if str(cells[1] or '').strip() == '№ п/п в смете'
                       and 'Цена единицы' in str(cells[5] or '')), None)
        if header is None:
            raise RimReportError('Нет заголовка формы РИМ: № п/п в смете / Цена единицы работ')
        output: list[RimWork] = []
        seen: set[tuple[str, str]] = set()
        for i, cells in enumerate(worksheet.iter_rows(min_row=header+1, max_col=15, values_only=True), header+1):
            no, code = cells[1], cells[13]
            if no is None or code is None or not str(no).strip() or not str(code).strip():
                continue
            if str(code).strip() in ('12', 'Перечень ГЭСН/ФЕР/ТЕР/КР'):
                continue
            if not str(cells[2] or '').strip():
                continue
            price, quantity, total = _decimal(cells[5]), _decimal(cells[4]), _decimal(cells[7])
            if price is None or quantity is None or total is None or price <= 0 or quantity == 0:
                raise RimReportError(f'В строке РИМ {i} нет рассчитанной цены, количества или стоимости')
            if abs(quantity * price-total) > max(Decimal('0.06'), abs(total)*Decimal('0.000001')):
                raise RimReportError(f'В строке РИМ {i} нарушена проверка цена × количество = итог')
            key = (str(no).strip(), NormCode(str(code)))
            if key in seen:
                raise RimReportError(f'Дублирующаяся позиция № {no} ({code}) в РИМ')
            seen.add(key)
            output.append(RimWork(
                number=key[0], code=str(code).strip(), name=str(cells[2]).strip(),
                unit=str(cells[3] or '').strip(), quantity=quantity,
                unit_price=price, total=total,
                labor_unit=_decimal(cells[9]), machinist_labor_unit=_decimal(cells[10]),
                labor_total=_decimal(cells[11]), machinist_labor_total=_decimal(cells[12]),
                average_price=_decimal(cells[6]), average_total=_decimal(cells[8]),
                section_code=str(cells[14] or '').strip(), sheet_row=i))
        if not output:
            raise RimReportError('В отчёте РИМ нет рассчитанных строк работ')
        return tuple(output)
    finally:
        workbook.close()

def attach_verified_rim_prices(rows: list[dict], report: tuple[RimWork, ...]) -> tuple[int, int]:
    """Match by exact GSFX position number AND code; preserve unmatched rows."""
    report_by_key = {(r.number, NormCode(r.code)): r for r in report}
    seen: set[tuple[str, str]] = set()
    matched = 0
    for row in rows:
        key = (str(row.get('number', '')).strip(), NormCode(row.get('code', '')))
        if key in seen:
            raise RimReportError(f'Дублирующаяся позиция GSFX {key}')
        seen.add(key)
        rim = report_by_key.get(key)
        if rim is None:
            if key[1].startswith('ФСБЦ'):
                row['include'] = False
                row['price_origin'] = 'Ресурс ФСБЦ: исключён из ОС, остаётся в сверке'
            else:
                row['price_origin'] = 'Нет в отчёте РИМ: требуется проверка'
                row['requires_price_review'] = True
            continue
        matched += 1
        row.update({
            'name': rim.name, 'unit': rim.unit,
            'quantity': format(rim.quantity, 'f'),
            'price': format(rim.unit_price, 'f'),
            'rim_total': format(rim.total, 'f'), 'rim_row': rim.sheet_row,
            'labor_unit': str(rim.labor_unit) if rim.labor_unit is not None else '',
            'machinist_labor_unit': str(rim.machinist_labor_unit) if rim.machinist_labor_unit is not None else '',
            'labor_total': str(rim.labor_total) if rim.labor_total is not None else '',
            'machinist_labor_total': str(rim.machinist_labor_total) if rim.machinist_labor_total is not None else '',
            'average_price': str(rim.average_price) if rim.average_price is not None else '',
            'average_total': str(rim.average_total) if rim.average_total is not None else '',
            'section_code': rim.section_code,
            'price_origin': 'Рассчитано ГРАНД-Сметой в РИМ-отчёте',
        })
    unexpected = set(report_by_key) - seen
    if unexpected:
        key = sorted(unexpected)[0]
        raise RimReportError(f'Отчёт РИМ от другой сметы: № {key[0]}, шифр {key[1]}')
    return matched, len(rows) - matched
