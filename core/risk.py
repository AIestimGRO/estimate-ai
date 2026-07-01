"""Price-spread risk checks for matched catalog entries."""

from dataclasses import dataclass

from core.catalog import CatalogEntry


DEFAULT_PRICE_SPREAD_LIMIT = 2.0
REASON_NONE = "none"
REASON_RATIO_EXCEEDED = "RATIO_EXCEEDED"
REASON_OUT_OF_APPROVED_RANGE = "OUT_OF_APPROVED_RANGE"

DEM_KEY_DEM = "DEM"
DEM_KEY_NO_DEM = "NO_DEM"
DEM_KEY_FILTER_OFF = "DEM_FILTER_OFF"


@dataclass(frozen=True)
class GesnException:
    """Approved price range for one matching key plus demolition dimension."""

    exception_key: str
    approved_min: float
    approved_max: float
    last_range_update_date: float


@dataclass(frozen=True)
class RiskResult:
    """Structured price risk result for one matched group."""

    is_flagged: bool
    reason: str
    flagged_entries: list[CatalogEntry]
    min_entry: CatalogEntry | None = None
    max_entry: CatalogEntry | None = None
    ratio: float = 0


def build_dem_key(source_has_demolition: bool, demontazh_filter_enabled: bool) -> str:
    """Build the demolition dimension of an exception key.

    Ports GetDemKey from Module6, DOMAIN_RULES.md section 5.2.
    """
    if not demontazh_filter_enabled:
        return DEM_KEY_FILTER_OFF
    return DEM_KEY_DEM if source_has_demolition else DEM_KEY_NO_DEM


def build_gesn_exception_key(unit_key: str, code_key: str, dem_key: str) -> str:
    """Build the approved-range lookup key `unit||code||DEM_FLAG`.

    Ports GetGesnExceptionKey from Module6, DOMAIN_RULES.md section 5.2.
    Expects an already-normalized unit and code (NormUnit/NormCode).
    """
    return f"{unit_key.strip()}||{code_key.strip()}||{dem_key.strip().upper()}"


def CheckPriceRisk(
    entries: list[CatalogEntry],
    gesn_exception: GesnException | None = None,
    price_spread_limit: float = DEFAULT_PRICE_SPREAD_LIMIT,
) -> RiskResult:
    """Evaluate ratio risk or approved-range override for matched entries."""
    if gesn_exception is not None:
        return _check_approved_range(entries, gesn_exception)

    return _check_ratio(entries, price_spread_limit)


def _check_ratio(
    entries: list[CatalogEntry],
    price_spread_limit: float,
) -> RiskResult:
    if price_spread_limit <= 0:
        return _not_flagged()

    priced_entries = [entry for entry in entries if entry.price > 0]
    if len(priced_entries) < 2:
        return _not_flagged()

    min_entry = min(priced_entries, key=lambda entry: entry.price)
    max_entry = max(priced_entries, key=lambda entry: entry.price)
    if min_entry.price <= 0:
        return _not_flagged()

    ratio = max_entry.price / min_entry.price
    if ratio >= price_spread_limit:
        return RiskResult(
            is_flagged=True,
            reason=REASON_RATIO_EXCEEDED,
            flagged_entries=[min_entry, max_entry],
            min_entry=min_entry,
            max_entry=max_entry,
            ratio=ratio,
        )

    return RiskResult(
        is_flagged=False,
        reason=REASON_NONE,
        flagged_entries=[],
        min_entry=min_entry,
        max_entry=max_entry,
        ratio=ratio,
    )


def _check_approved_range(
    entries: list[CatalogEntry],
    gesn_exception: GesnException,
) -> RiskResult:
    flagged_entries: list[CatalogEntry] = []

    for entry in entries:
        # Mirrors an explicit guard in the original VBA (Module6,
        # MarkOutOfApprovedRange): entries with no recorded added-date are
        # intentionally treated as not new enough to check, to avoid false
        # positives on data with missing dates. This is deliberate macro
        # behavior, not an accidental gap.
        if entry.added_date_serial <= 0:
            continue
        if entry.added_date_serial <= gesn_exception.last_range_update_date:
            continue
        if _is_outside_approved_range(entry.price, gesn_exception):
            flagged_entries.append(entry)

    if flagged_entries:
        return RiskResult(
            is_flagged=True,
            reason=REASON_OUT_OF_APPROVED_RANGE,
            flagged_entries=flagged_entries,
        )

    return _not_flagged()


def _is_outside_approved_range(price: float, gesn_exception: GesnException) -> bool:
    below_min = gesn_exception.approved_min > 0 and price < gesn_exception.approved_min
    above_max = gesn_exception.approved_max > 0 and price > gesn_exception.approved_max
    return below_min or above_max


def _not_flagged() -> RiskResult:
    return RiskResult(
        is_flagged=False,
        reason=REASON_NONE,
        flagged_entries=[],
    )
