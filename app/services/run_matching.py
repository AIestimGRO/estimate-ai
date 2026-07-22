"""Application service for one end-to-end estimate matching run.

Orchestrates the already-tested core modules in the same order as the VBA
macro ProcessSmeta (Module4): build catalog -> match each estimate row ->
price-spread risk (with approved-range override) -> section code ->
recommended price -> `/KR` code suffix. This module contains no matching or
pricing math of its own; it only wires the pieces together.

Ports the orchestration of ProcessSmeta, Module4, DOMAIN_RULES.md sections
3-6.
"""

from dataclasses import dataclass, field
from numbers import Real
from pathlib import Path

from core.catalog import BuildCatalog, Catalog, CatalogRow
from core.excel_io import Settings, read_catalog_rows, read_estimate_rows
from core.exclusions import NameExclusionRule
from core.matching import (
    REASON_EXCLUDED_BY_NAME,
    REASON_INVALID_INPUT,
    REASON_MATCHED,
    AnalogColumn,
    EstimateRow,
    MatchEstimateRow,
    MatchResult,
)
from core.normalize import HasDemontazh, NormCode, NormUnit
from core.pricing import CalculateAveragePrice
from core.risk import (
    DEFAULT_PRICE_SPREAD_LIMIT,
    CheckPriceRisk,
    GesnException,
    RiskResult,
    build_dem_key,
    build_gesn_exception_key,
)
from core.sections import ResolveSectionCode
from core.tkp_matching import (
    TkpCatalogEntry,
    TkpMatch,
    find_best_tkp_matches,
)


# Cyrillic "/KR" suffix appended to the code column, DOMAIN_RULES.md section 6.
KR = "\u041a\u0420"
KR_END = f"/{KR}"
KR_SUFFIX = f" /{KR}"


@dataclass(frozen=True)
class EstimateRowResult:
    """Structured matching result for one estimate row."""

    row_index: int
    estimate_row: EstimateRow
    norm_code: str
    norm_unit: str
    is_demolition: bool
    match_result: MatchResult
    risk_result: RiskResult
    section_code: str
    recommended_price: float | None
    kr_code: str | None
    exception_key: str
    status: str
    tkp_match: TkpMatch | None = None

    @property
    def analogs(self) -> list[AnalogColumn]:
        return self.match_result.analogs

    @property
    def has_analogs(self) -> bool:
        return self.match_result.has_analogs

    @property
    def has_tkp_analog(self) -> bool:
        return self.tkp_match is not None


@dataclass(frozen=True)
class MatchingRunResult:
    """Aggregate result of a single matching run over all estimate rows."""

    rows: list[EstimateRowResult] = field(default_factory=list)
    catalog_key_count: int = 0
    matched_row_count: int = 0
    flagged_row_count: int = 0
    tkp_matched_row_count: int = 0


def run_matching(
    catalog_rows: list[CatalogRow],
    estimate_rows: list[EstimateRow],
    *,
    name_exclusion_rules: list[NameExclusionRule] | None = None,
    gesn_exceptions: dict[str, GesnException] | None = None,
    demontazh_filter_enabled: bool = True,
    price_spread_limit: float = DEFAULT_PRICE_SPREAD_LIMIT,
    regional_coefficient: float = 1.0,
    tkp_catalog_index: list[TkpCatalogEntry] | None = None,
    use_tkp_analogs: bool = False,
) -> MatchingRunResult:
    """Run matching for pre-read structured rows (no Excel I/O here)."""
    rules = [] if name_exclusion_rules is None else name_exclusion_rules
    exceptions = {} if gesn_exceptions is None else gesn_exceptions

    catalog: Catalog = BuildCatalog(catalog_rows, rules)
    tkp_index = _priced_tkp_entries(tkp_catalog_index) if use_tkp_analogs else []

    row_results: list[EstimateRowResult] = []
    matched_row_count = 0
    flagged_row_count = 0
    tkp_matched_row_count = 0

    for row_index, estimate_row in enumerate(estimate_rows, start=1):
        row_result = _match_one_row(
            row_index,
            estimate_row,
            catalog,
            rules,
            exceptions,
            demontazh_filter_enabled,
            price_spread_limit,
            regional_coefficient,
            tkp_index,
            use_tkp_analogs,
        )
        row_results.append(row_result)

        if row_result.has_analogs:
            matched_row_count += 1
        if row_result.risk_result.is_flagged:
            flagged_row_count += 1
        if row_result.has_tkp_analog:
            tkp_matched_row_count += 1

    return MatchingRunResult(
        rows=row_results,
        catalog_key_count=len(catalog),
        matched_row_count=matched_row_count,
        flagged_row_count=flagged_row_count,
        tkp_matched_row_count=tkp_matched_row_count,
    )


def run_matching_from_files(
    catalog_path: str | Path,
    estimate_path: str | Path,
    *,
    settings: Settings | None = None,
    name_exclusion_rules: list[NameExclusionRule] | None = None,
    gesn_exceptions: dict[str, GesnException] | None = None,
    demontazh_filter_enabled: bool = True,
    price_spread_limit: float = DEFAULT_PRICE_SPREAD_LIMIT,
    regional_coefficient: float = 1.0,
    tkp_catalog_index: list[TkpCatalogEntry] | None = None,
    use_tkp_analogs: bool = False,
) -> MatchingRunResult:
    """Read catalog/estimate workbooks, then run matching over their rows."""
    catalog_rows = read_catalog_rows(catalog_path, settings)
    estimate_rows = read_estimate_rows(estimate_path, settings)

    return run_matching(
        catalog_rows,
        estimate_rows,
        name_exclusion_rules=name_exclusion_rules,
        gesn_exceptions=gesn_exceptions,
        demontazh_filter_enabled=demontazh_filter_enabled,
        price_spread_limit=price_spread_limit,
        regional_coefficient=regional_coefficient,
        tkp_catalog_index=tkp_catalog_index,
        use_tkp_analogs=use_tkp_analogs,
    )


def _match_one_row(
    row_index: int,
    estimate_row: EstimateRow,
    catalog: Catalog,
    rules: list[NameExclusionRule],
    exceptions: dict[str, GesnException],
    demontazh_filter_enabled: bool,
    price_spread_limit: float,
    regional_coefficient: float,
    tkp_catalog_index: list[TkpCatalogEntry],
    use_tkp_analogs: bool,
) -> EstimateRowResult:
    norm_code = NormCode(estimate_row.code)
    norm_unit = NormUnit(estimate_row.unit)
    is_demolition = HasDemontazh(estimate_row.work_name)

    match_result = MatchEstimateRow(
        estimate_row,
        catalog,
        rules,
        demontazh_filter_enabled,
    )

    exception_key = ""
    gesn_exception: GesnException | None = None
    if norm_code != "" and norm_unit != "":
        dem_key = build_dem_key(is_demolition, demontazh_filter_enabled)
        exception_key = build_gesn_exception_key(norm_unit, norm_code, dem_key)
        gesn_exception = exceptions.get(exception_key)

    matched_entries = [analog.entry for analog in match_result.analogs]
    risk_result = CheckPriceRisk(
        matched_entries,
        gesn_exception,
        price_spread_limit,
    )

    section_code = ""
    if norm_code != "":
        section_code = ResolveSectionCode(estimate_row.code, is_demolition)

    tkp_match: TkpMatch | None = None
    if use_tkp_analogs and tkp_catalog_index:
        matches = find_best_tkp_matches(
            estimate_row.work_name,
            estimate_row.unit,
            tkp_catalog_index,
            limit=1,
        )
        if matches:
            tkp_match = matches[0]

    recommended_price = _recommended_price(
        estimate_row.base_price,
        match_result,
        regional_coefficient,
        tkp_match,
    )

    if match_result.has_analogs:
        kr_code = _append_kr_suffix(estimate_row.code)
    else:
        # No analog found: still fill the /КР column, but with the plain
        # ГЭСН code as-is -- no "/КР" suffix is added (2026-07 rule).
        plain_code = _normalize_code_text(estimate_row.code)
        kr_code = plain_code if plain_code != "" else None

    return EstimateRowResult(
        row_index=row_index,
        estimate_row=estimate_row,
        norm_code=norm_code,
        norm_unit=norm_unit,
        is_demolition=is_demolition,
        match_result=match_result,
        risk_result=risk_result,
        section_code=section_code,
        recommended_price=recommended_price,
        kr_code=kr_code,
        exception_key=exception_key,
        status=match_result.reason,
        tkp_match=tkp_match,
    )


def _recommended_price(
    base_price: object,
    match_result: MatchResult,
    regional_coefficient: float,
    tkp_match: TkpMatch | None = None,
) -> float | None:
    """Value port of the average-price formula, DOMAIN_RULES.md section 6.

    The regional coefficient is applied to analog prices only; the base price
    is never adjusted, mirroring ProcessSmeta (analog cells * coef, average
    formula references the raw base cell).
    """
    base = _parse_positive_number(base_price)
    if base is None:
        return None

    analog_prices = [
        analog.entry.price * regional_coefficient for analog in match_result.analogs
    ]
    if tkp_match is not None:
        tkp_price = _parse_positive_number(tkp_match.entry.winner_unit_price_no_vat)
        if tkp_price is not None:
            # TKP winner prices are already stored at their source price level;
            # unlike RNMC ZLVL prices, they are not scaled by the estimate's
            # regional coefficient.
            analog_prices.append(tkp_price)
    return CalculateAveragePrice(base, analog_prices)


def _priced_tkp_entries(
    index: list[TkpCatalogEntry] | None,
) -> list[TkpCatalogEntry]:
    """Keep only candidates that can be written and included in the average."""
    if not index:
        return []
    return [
        entry
        for entry in index
        if _parse_positive_number(entry.winner_unit_price_no_vat) is not None
    ]


def _normalize_code_text(code: object) -> str:
    text = "" if code is None else str(code)
    text = text.replace("\r", " ")
    text = text.replace("\n", " ")
    text = text.replace("\t", " ")
    text = text.replace("\u00a0", " ")
    while "  " in text:
        text = text.replace("  ", " ")
    return text.strip()


def _append_kr_suffix(code: object) -> str:
    """Append the `/KR` suffix to a code, mirroring step 8 of ProcessSmeta."""
    text = _normalize_code_text(code)

    if text == "":
        return KR_END
    if text.upper().endswith(KR_END):
        return text
    return text + KR_SUFFIX


def _parse_positive_number(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, Real):
        number = float(value)
    else:
        text = str(value).strip()
        if text == "":
            return None
        try:
            number = float(text)
        except ValueError:
            return None

    if number <= 0:
        return None
    return number
