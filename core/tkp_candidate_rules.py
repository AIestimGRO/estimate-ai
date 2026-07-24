"""Deterministic safety rules for TKP shadow candidates."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from core.multiplicity import multiplicity_is_compatible
from core.normalize import BaseUnit, NormUnit
from core.tkp_matching import TkpCatalogEntry


_CONFIG_PATH = (
    Path(__file__).resolve().parents[1]
    / "data"
    / "config"
    / "tkp_shadow_rules.json"
)
_UNIT_SCALE_PATTERN = re.compile(r"^(\d+)(.+)$")

REASON_UNIT_CONFLICT = "unit_conflict"
REASON_WORK_TYPE_CONFLICT = "work_type_conflict"
REASON_MULTIPLICITY_CONFLICT = "multiplicity_conflict"
REASON_PRICE_MISSING = "price_missing"


@dataclass(frozen=True)
class UnitConversion:
    query_unit: str
    candidate_unit: str
    base_unit: str
    query_scale: float
    candidate_scale: float

    @property
    def price_factor(self) -> float:
        return self.query_scale / self.candidate_scale


@dataclass(frozen=True)
class CandidateRuleResult:
    accepted: bool
    reason: str
    normalized_unit_price: float | None
    unit_conversion: UnitConversion | None


def evaluate_tkp_candidate(
    query_name: object,
    query_unit: object,
    entry: TkpCatalogEntry,
) -> CandidateRuleResult:
    conversion = compatible_unit_conversion(query_unit, entry.unit)
    if conversion is None:
        return CandidateRuleResult(False, REASON_UNIT_CONFLICT, None, None)

    query_type = detect_work_type(query_name)
    candidate_type = detect_work_type(
        " ".join(
            value
            for value in (
                entry.section_name,
                entry.subsection_name,
                entry.item_name,
            )
            if value
        )
    )
    if work_types_conflict(query_type, candidate_type):
        return CandidateRuleResult(
            False,
            REASON_WORK_TYPE_CONFLICT,
            None,
            conversion,
        )

    if not multiplicity_is_compatible(query_name, entry.item_name):
        return CandidateRuleResult(
            False,
            REASON_MULTIPLICITY_CONFLICT,
            None,
            conversion,
        )

    price = _positive_float(entry.winner_unit_price_no_vat)
    if price is None:
        return CandidateRuleResult(
            False,
            REASON_PRICE_MISSING,
            None,
            conversion,
        )
    return CandidateRuleResult(
        True,
        "",
        price * conversion.price_factor,
        conversion,
    )


def compatible_unit_conversion(
    query_unit: object,
    candidate_unit: object,
) -> UnitConversion | None:
    query_base, query_scale = split_unit_scale(query_unit)
    candidate_base, candidate_scale = split_unit_scale(candidate_unit)
    if not query_base or not candidate_base or query_base != candidate_base:
        return None
    return UnitConversion(
        query_unit=NormUnit(query_unit),
        candidate_unit=NormUnit(candidate_unit),
        base_unit=query_base,
        query_scale=query_scale,
        candidate_scale=candidate_scale,
    )


def split_unit_scale(value: object) -> tuple[str, float]:
    normalized = NormUnit(value)
    if not normalized:
        return "", 1.0
    base = BaseUnit(normalized)
    match = _UNIT_SCALE_PATTERN.fullmatch(normalized)
    if match is None:
        return base, 1.0
    scale = int(match.group(1))
    if scale <= 0:
        return "", 1.0
    return base, float(scale)


def detect_work_type(value: object) -> str:
    text = " ".join(str(value or "").casefold().split())
    if not text:
        return ""
    for work_type in ("demolition", "restoration", "installation"):
        if any(root in text for root in _work_type_roots().get(work_type, ())):
            return work_type
    return ""


def work_types_conflict(query_type: str, candidate_type: str) -> bool:
    if candidate_type == "demolition":
        return query_type != "demolition"
    if query_type == "demolition":
        return candidate_type != "demolition"
    return False


@lru_cache(maxsize=1)
def _work_type_roots() -> dict[str, tuple[str, ...]]:
    try:
        payload = json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, TypeError, ValueError):
        return {}
    groups = payload.get("work_types", {})
    return {
        str(key): tuple(
            str(value).casefold().strip()
            for value in values
            if str(value).strip()
        )
        for key, values in groups.items()
        if isinstance(values, list)
    }


def _positive_float(value: object) -> float | None:
    try:
        parsed = float(value)
    except (TypeError, ValueError):
        return None
    return parsed if parsed > 0 else None
