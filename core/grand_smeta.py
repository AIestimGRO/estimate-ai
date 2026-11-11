"""Read-only import of GrandSmeta GSFX source records.

This is an extraction boundary, not an estimate price calculator. Per-unit
work prices must be provided by a verified calculation/export before matching.
"""
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

class GrandSmetaFormatError(ValueError):
    """The upload is not a supported GSFX document."""

@dataclass(frozen=True)
class GrandResource:
    kind: str
    code: str
    name: str
    unit: str
    quantity_per_unit: Decimal | None
    current_unit_price: Decimal | None
    base_unit_price: Decimal | None
    deleted: bool

@dataclass(frozen=True)
class GrandPosition:
    source_id: str
    number: str
    chapter: str
    code: str
    name: str
    unit: str
    quantity: Decimal | None
    resources: tuple[GrandResource, ...]
    unit_work_price: None = None

@dataclass(frozen=True)
class GrandEstimate:
    local_number: str
    object_name: str
    description: str
    region: str
    positions: tuple[GrandPosition, ...]

def _decimal(value: str | None) -> Decimal | None:
    if value is None or not value.strip():
        return None
    try:
        value = Decimal(value.strip().replace(' ', '').replace('\u00a0', '').replace(',', '.'))
        return value if value.is_finite() else None
    except InvalidOperation:
        return None

def _price(node: ET.Element, kind: str) -> Decimal | None:
    price_node = node.find(kind)
    return _decimal(price_node.get('Value')) if price_node is not None else None

def read_gsfx(path: str | Path) -> GrandEstimate:
    """Extract all chapter positions and nested resources, without pricing guesses."""
    try:
        with ZipFile(path) as archive:
            names = archive.namelist()
            if 'Data.xml' not in names or 'Properties.txt' not in names:
                raise GrandSmetaFormatError('Required GSFX archive entries are missing')
            info = archive.getinfo('Data.xml')
            if info.file_size > 50_000_000:
                raise GrandSmetaFormatError('GSFX XML exceeds the permitted size')
            xml_data = archive.read(info)
    except (BadZipFile, OSError, KeyError) as exc:
        raise GrandSmetaFormatError('Cannot open GSFX archive') from exc
    try:
        root = ET.fromstring(xml_data)
    except ET.ParseError as exc:
        raise GrandSmetaFormatError('Invalid GSFX Data.xml') from exc
    if root.tag != 'Document' or root.get('Generator') != 'GrandSmeta':
        raise GrandSmetaFormatError('Unsupported GSFX document')
    properties = root.find('Properties')
    region_info = root.find('RegionInfo')
    chapters = root.find('Chapters')
    if chapters is None:
        raise GrandSmetaFormatError('GSFX chapters are missing')
    positions: list[GrandPosition] = []
    for chapter in chapters.iter('Chapter'):
        for pos in chapter.findall('Position'):
            quantity_node = pos.find('Quantity')
            quantity = _decimal(quantity_node.get('Result')) if quantity_node is not None else None
            resources_node = pos.find('Resources')
            resources: list[GrandResource] = []
            if resources_node is not None:
                for resource in resources_node:
                    if resource.tag not in ('Tzr', 'Tzm', 'Mch', 'Mat', 'Ob', 'Perevozka'):
                        continue
                    resources.append(GrandResource(
                        kind=resource.tag,
                        code=resource.get('Code', ''),
                        name=resource.get('Caption', ''),
                        unit=resource.get('Units', ''),
                        quantity_per_unit=_decimal(resource.get('Quantity')),
                        current_unit_price=_price(resource, 'PriceCurr'),
                        base_unit_price=_price(resource, 'PriceBase'),
                        deleted='Deleted' in resource.get('Attribs', '').split(),
                    ))
            positions.append(GrandPosition(
                source_id=pos.get('SysID', ''),
                number=pos.get('Number', ''),
                chapter=chapter.get('Caption', ''),
                code=pos.get('Code', ''),
                name=pos.get('Caption', ''),
                unit=pos.get('Units', ''),
                quantity=quantity,
                resources=tuple(resources),
            ))
    return GrandEstimate(
        local_number=properties.get('LocNum', '') if properties is not None else '',
        object_name=properties.get('Object', '') if properties is not None else '',
        description=properties.get('Description', '') if properties is not None else '',
        region=region_info.get('AdmRegionZone', '') if region_info is not None else '',
        positions=tuple(positions),
    )
