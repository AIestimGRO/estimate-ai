"""Estimate-row matching against a built catalog."""

from dataclasses import dataclass
from numbers import Real

from core.catalog import Catalog, CatalogEntry
from core.exclusions import NameExclusionRule, is_name_excluded
from core.normalize import AnalogSearchKey, HasDemontazh, NormCode, NormUnit


SMETA_SCOPE = "SMETA"
REASON_MATCHED = "matched"
REASON_EXCLUDED_BY_NAME = "excluded_by_name"
REASON_INVALID_INPUT = "invalid_input"
REASON_NO_MATCH = "no_match"
REASON_FILTERED_BY_DEMOLITION = "filtered_by_demolition"


@dataclass(frozen=True)
class EstimateRow:
    """Raw estimate input row for matching, DOMAIN_RULES.md section 4."""

    code: object
    unit: object
    work_name: object
    base_price: object


@dataclass(frozen=True)
class AnalogColumn:
    """One output analog column for a task/price-position pair."""

    task_id: str
    price_position: int
    entry: CatalogEntry


@dataclass(frozen=True)
class MatchResult:
    """Analog matching result plus the reason for the outcome."""

    analogs: list[AnalogColumn]
    has_analogs: bool
    reason: str


def MatchEstimateRow(
    estimate_row: EstimateRow,
    catalog: Catalog,
    name_exclusion_rules: list[NameExclusionRule] | None = None,
    demontazh_filter_enabled: bool = True,
) -> MatchResult:
    """Match one estimate row to catalog analogs.

    Ports the matching portion of ProcessSmeta from Module4,
    DOMAIN_RULES.md section 4.
    """
    rules = [] if name_exclusion_rules is None else name_exclusion_rules

    if is_name_excluded(rules, SMETA_SCOPE, estimate_row.work_name):
        return _zero(REASON_EXCLUDED_BY_NAME)

    norm_code = NormCode(estimate_row.code)
    norm_unit = NormUnit(estimate_row.unit)
    base_price = _parse_positive_price(estimate_row.base_price)
    if norm_code == "" or norm_unit == "" or base_price is None:
        return _zero(REASON_INVALID_INPUT)

    matching_key = AnalogSearchKey(estimate_row.unit, estimate_row.code)
    if matching_key == "":
        return _zero(REASON_INVALID_INPUT)

    task_groups = catalog.get(matching_key)
    if not task_groups:
        return _zero(REASON_NO_MATCH)

    row_is_demolition = HasDemontazh(estimate_row.work_name)
    analogs = _build_analog_columns(
        task_groups,
        row_is_demolition,
        demontazh_filter_enabled,
    )

    if not analogs:
        return _zero(REASON_FILTERED_BY_DEMOLITION)

    return MatchResult(
        analogs=analogs,
        has_analogs=True,
        reason=REASON_MATCHED,
    )


def _zero(reason: str) -> MatchResult:
    return MatchResult(
        analogs=[],
        has_analogs=False,
        reason=reason,
    )


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


def _build_analog_columns(
    task_groups: dict[str, list[CatalogEntry]],
    row_is_demolition: bool,
    demontazh_filter_enabled: bool,
) -> list[AnalogColumn]:
    analogs: list[AnalogColumn] = []

    for task_id, entries in task_groups.items():
        filtered_entries = [
            entry
            for entry in entries
            if not demontazh_filter_enabled or entry.is_demolition == row_is_demolition
        ]

        for price_position, entry in enumerate(filtered_entries, start=1):
            analogs.append(
                AnalogColumn(
                    task_id=task_id,
                    price_position=price_position,
                    entry=entry,
                )
            )

    return analogs
