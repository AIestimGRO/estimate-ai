"""Catalog construction ported from the VBA matching macro."""

# Notes for future modules:
# - _parse_iso_date accepts only strict ISO strings for now. When core/excel_io.py
#   is built, Excel dates will likely arrive as datetime objects (already handled)
#   or Excel serial floats (not handled yet); extend date parsing there.
# - Dedup within one task_id is O(n^2), which is fine for tens of entries per task.
#   Revisit this if catalogs ever grow to hundreds+ of entries per task.

from dataclasses import dataclass
from datetime import date, datetime
from numbers import Real

from core.exclusions import NameExclusionRule, is_name_excluded
from core.normalize import AnalogSearchKey, HasDemontazh, NormCode, NormUnit


DEDUP_PCT = 0.04
CATALOG_SCOPE = "CATALOG"
VBA_DATE_BASE = date(1899, 12, 30)


@dataclass(frozen=True)
class CatalogRow:
    """Raw catalog input row for BuildCatalog, DOMAIN_RULES.md section 3."""

    task_id: object
    price: object
    code: object
    unit: object
    work_name: object
    region: object = ""
    added_date: object = None
    total_price: object = None
    labor_unit: object = None
    labor_total: object = None
    machine_labor_unit: object = None
    machine_labor_total: object = None


@dataclass(frozen=True)
class CatalogEntry:
    """Surviving catalog entry in the same order as the VBA array payload."""

    price: float
    region: str
    is_demolition: bool
    source_row_number: int
    original_row: CatalogRow
    task_id: str
    norm_code: str
    norm_unit: str
    added_date_serial: float


Catalog = dict[str, dict[str, list[CatalogEntry]]]


def BuildCatalog(
    rows: list[CatalogRow],
    name_exclusion_rules: list[NameExclusionRule] | None = None,
) -> Catalog:
    """Build the nested catalog structure from rows.

    Ports BuildCatalog from Module3, DOMAIN_RULES.md section 3.
    """
    rules = [] if name_exclusion_rules is None else name_exclusion_rules
    catalog: Catalog = {}

    for source_row_number, row in enumerate(rows, start=1):
        task_id = _trim_text(row.task_id)
        if task_id == "":
            continue

        price = _parse_positive_price(row.price)
        if price is None:
            continue

        norm_code = NormCode(row.code)
        if norm_code == "":
            continue

        norm_unit = NormUnit(row.unit)
        if norm_unit == "":
            continue

        matching_key = AnalogSearchKey(row.unit, row.code)
        if matching_key == "":
            continue

        if is_name_excluded(rules, CATALOG_SCOPE, row.work_name):
            continue

        entry = CatalogEntry(
            price=price,
            region=_trim_text(row.region),
            is_demolition=HasDemontazh(row.work_name),
            source_row_number=source_row_number,
            original_row=row,
            task_id=task_id,
            norm_code=norm_code,
            norm_unit=norm_unit,
            added_date_serial=_catalog_date_serial(row.added_date),
        )

        catalog.setdefault(matching_key, {}).setdefault(task_id, []).append(entry)

    return _deduplicate_catalog(catalog)


def _trim_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _parse_positive_price(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, Real):
        price = float(value)
    else:
        text = str(value).strip()
        if text == "":
            return None
        try:
            price = float(text)
        except ValueError:
            return None

    if price <= 0:
        return None

    return price


def _catalog_date_serial(value: object) -> float:
    if value is None:
        return 0

    if isinstance(value, datetime):
        return float((value.date() - VBA_DATE_BASE).days)

    if isinstance(value, date):
        return float((value - VBA_DATE_BASE).days)

    if isinstance(value, str):
        text = value.strip()
        if text == "":
            return 0

        parsed = _parse_iso_date(text)
        if parsed is not None:
            return float((parsed - VBA_DATE_BASE).days)

    return 0


def _parse_iso_date(value: str) -> date | None:
    try:
        return date.fromisoformat(value)
    except ValueError:
        return None


def _deduplicate_catalog(catalog: Catalog) -> Catalog:
    for task_groups in catalog.values():
        for task_id, entries in task_groups.items():
            kept_entries: list[CatalogEntry] = []

            for entry in entries:
                should_keep = True

                for kept_entry in kept_entries:
                    if entry.is_demolition != kept_entry.is_demolition:
                        continue

                    price_delta = abs(entry.price - kept_entry.price) / kept_entry.price
                    if price_delta <= DEDUP_PCT:
                        should_keep = False
                        break

                if should_keep:
                    kept_entries.append(entry)

            task_groups[task_id] = kept_entries

    return catalog
