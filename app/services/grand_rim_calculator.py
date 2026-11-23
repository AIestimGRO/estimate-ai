"""Conservative, auditable RIM work prices from a single GrandSmeta GSFX.

Only calculate positions whose normative category, northern rate and resource
prices are unambiguous. Unsupported estimates remain unpriced for review; the
RNMC catalog and TKP are never used to fabricate work costs.
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from functools import lru_cache
import json
from pathlib import Path
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile


ZERO = Decimal('0')
ONE = Decimal('1')
CENT = Decimal('0.01')
_REFERENCE_PATH = Path(__file__).resolve().parents[2] / 'core/data/nr_sp_fsnb2022_dop19.json'


@dataclass(frozen=True)
class CalculatedWork:
    total: Decimal
    unit_price: Decimal
    labor_per_norm_unit: Decimal
    machinist_labor_per_norm_unit: Decimal
    nr_sp_code: str
    unpriced_materials: int


def _num(value: str | None) -> Decimal | None:
    if value is None or not value.strip():
        return None
    try:
        result = Decimal(value.replace(' ', '').replace('\u00a0', '').replace(',', '.'))
        return result if result.is_finite() else None
    except InvalidOperation:
        return None


@lru_cache(maxsize=1)
def _reference() -> dict:
    return json.loads(_REFERENCE_PATH.read_text(encoding='utf-8'))['groups']


def _northern_class(root: ET.Element) -> str | None:
    """Only known, tested territorial classifications; never default north to 0."""
    region = root.find('RegionInfo')
    if region is None:
        return None
    if region.get('AdmRegionCode') == '14' and region.get('AdmRegionZoneCode') == '4':
        return 'overhead_far_north'
    name = (region.get('AdmRegionZone') or '').casefold().replace('ё', 'е')
    if ('москва' in name or 'московская область' in name or
            'санкт-петербург' in name or 'ленинградская область' in name):
        return 'overhead_ordinary'
    return None


def _group_for_code(code: str, groups: dict) -> str | None:
    """Resolve only code ranges with a single documented NR/SP assignment."""
    match = re.fullmatch(r'ГЭСН(\d{2})-(\d{2})-(\d{3})-(\d{2})', code.strip().upper())
    if not match:
        return None
    volume, section, _, _ = match.groups()
    if volume == '01':
        # Table 1, sections 1.1-1.7: mechanized excavation.
        return '001.1' if section == '01' else None
    if volume == '46':
        # Table 40.1: GESN46 section 3; section 2 has exceptions.
        return '040.1' if section == '03' else None
    candidates = []
    prefix = re.compile(r'ГЭСН\s+81-02-' + volume + r'\s*[-–]')
    for group, rule in groups.items():
        applicability = rule['applicability']
        if prefix.search(applicability) and 'за исключением' not in applicability.casefold():
            candidates.append(group)
    return candidates[0] if len(candidates) == 1 else None


def _calculate_position(position: ET.Element, groups: dict, northern: str) -> CalculatedWork | None:
    code = _group_for_code(position.get('Code', ''), groups)
    if code is None:
        return None
    norm_qty = position.find('Quantity')
    qty = _num(norm_qty.get('Result')) if norm_qty is not None else None
    if qty is None or qty <= 0:
        return None
    rates = groups[code]
    percent = Decimal(rates[northern] + rates['profit']) / Decimal(100)
    elements: dict[str, Decimal] = {'Tzr': ZERO, 'Mch': ZERO, 'mz': ZERO, 'Mat': ZERO, 'transport': ZERO}
    labor: dict[str, Decimal] = {'Tzr': ZERO, 'mz': ZERO}
    resources = position.find('Resources')
    if resources is None or not len(resources):
        return None
    declared_tzm = ZERO
    unpriced_materials = 0
    for res in resources:
        if res.tag not in ('Tzr', 'Tzm', 'Mch', 'Mat'):
            return None
        unit_qty = _num(res.get('Quantity'))
        if unit_qty is None or unit_qty < 0:
            return None
        deleted = 'Deleted' in res.get('Attribs', '').split()
        if res.tag == 'Tzm':
            declared_tzm += unit_qty
            continue
        if deleted:
            if res.tag == 'Mat':
                transportation = res.find('Transportation')
                if transportation is not None:
                    mass = _num(res.get('Mass'))
                    if mass is None or mass < 0:
                        return None
                    for item in transportation:
                        value = _num(item.get('PriceCurr'))
                        multiplier = _num(item.get('KC')) if item.get('KC') else ONE
                        if value is None or multiplier is None:
                            return None
                        elements['transport'] += unit_qty * mass / Decimal(1000) * value * multiplier
            continue
        price_node = res.find('PriceCurr')
        rate = _num(price_node.get('Value')) if price_node is not None else None
        if rate is None and res.tag == 'Mat':
            # In an LSR without materials, an unpriced material is not
            # included in GrandSmeta's priced work subtotal. Report its
            # omission separately, without borrowing a catalog price.
            unpriced_materials += 1
            continue
        if rate is None or rate <= 0:
            return None
        elements[res.tag] += unit_qty * rate
        if res.tag == 'Tzr':
            labor['Tzr'] += unit_qty
        if res.tag == 'Mch':
            wage = res.find('Tzm2022')
            if wage is None:
                # A machine without operators is valid when no Tzm wage is
                # declared. Never guess missing operator wages otherwise.
                continue
            hours = _num(wage.get('Quantity'))
            salary = _num(wage.get('Price'))
            if hours is None or salary is None or hours < 0 or salary <= 0:
                return None
            labor['mz'] += unit_qty * hours
            elements['mz'] += unit_qty * hours * salary
    if abs(declared_tzm - labor['mz']) > Decimal('0.02'):
        return None
    for coefficient in position.findall('Koefficients/K'):
        flags = set((coefficient.get('Options') or '').split())
        if 'Disabled' in flags:
            continue
        # Other coefficient families can affect NR/SP or resource prices and
        # require dedicated validation before they may be calculated.
        if any(key not in ('Value_OZ', 'Value_EM', 'Value_PZ', 'Value_MT')
               for key in coefficient.attrib if key.startswith('Value_')):
            return None
        common = _num(coefficient.get('Value_PZ')) if 'Value_PZ' in coefficient.attrib else ONE
        oz = _num(coefficient.get('Value_OZ')) if 'Value_OZ' in coefficient.attrib else ONE
        em = _num(coefficient.get('Value_EM')) if 'Value_EM' in coefficient.attrib else ONE
        mt = _num(coefficient.get('Value_MT')) if 'Value_MT' in coefficient.attrib else ONE
        if any(v is None for v in (common, oz, em, mt)) or min(common, oz, em, mt) <= 0:
            return None
        if common != ONE and 'PzAll' not in flags:
            return None
        if oz != ONE and 'OzpTz' not in flags:
            return None
        if em != ONE and 'EmAll' not in flags:
            return None
        if mt != ONE and 'MatQty' not in flags:
            return None
        for key in elements:
            factor = common
            if key == 'Tzr':
                factor *= oz
            elif key in ('Mch', 'mz'):
                factor *= em
            elif key in ('Mat', 'transport'):
                factor *= mt
            elements[key] *= factor
        labor['Tzr'] *= common * oz
        labor['mz'] *= common * em
    direct = sum(elements.values(), ZERO)
    wage = elements['Tzr'] + elements['mz']
    norm_total = direct + wage * percent
    if norm_total <= 0:
        return None
    total = (norm_total * qty).quantize(CENT, rounding=ROUND_HALF_UP)
    return CalculatedWork(
        total=total, unit_price=total / qty,
        labor_per_norm_unit=labor['Tzr'],
        machinist_labor_per_norm_unit=labor['mz'],
        nr_sp_code=code,
        unpriced_materials=unpriced_materials,
    )


def calculate_gsfx_prices(path: Path) -> tuple[CalculatedWork | None, ...]:
    """Calculate only verified FSNB-2022/dop19 RIM positions, in source order."""
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read('Data.xml'))
    chapters = root.find('Chapters')
    positions = [pos for chapter in chapters.iter('Chapter')
                 for pos in chapter.findall('Position')] if chapters is not None else []
    region = root.find('RegionInfo')
    frsn = root.find('FRSN_Info')
    northern = _northern_class(root)
    if (northern is None or region is None or frsn is None or
            not (region.get('RegionName') or '').startswith('ФСНБ-2022') or
            frsn.get('AddOnNumber') != '19'):
        return (None,) * len(positions)
    # Global active coefficients can change individual positions. Until such
    # coefficients are implemented, the estimate cannot be priced safely.
    if any('Disabled' not in (k.get('Options') or '').split()
           for k in root.findall('./Koefficients/K')):
        return (None,) * len(positions)
    groups = _reference()
    return tuple(_calculate_position(pos, groups, northern) for pos in positions)
