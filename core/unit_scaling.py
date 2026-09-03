"""Deterministic unit-scale conversion for TKP quantity and unit-price checks."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from functools import lru_cache
from numbers import Real
from pathlib import Path

from core.normalize import BaseUnit, NormUnit

_CONFIG_PATH = Path(__file__).resolve().parents[1] / "data" / "config" / "unit_scaling.json"
_LEADING_SCALE_RE = re.compile(r"^\s*(\d+)\s*(.+?)\s*$")
_TRAILING_PUNCT_RE = re.compile(r"[.,;:]+$")


@dataclass(frozen=True)
class UnitConversion:
    query_unit: str
    candidate_unit: str
    base_unit: str
    query_scale: float
    candidate_scale: float

    @property
    def price_factor(self) -> float:
        """Multiply candidate unit price to express it in the query unit."""
        return self.query_scale / self.candidate_scale


def split_unit_scale(value: object) -> tuple[str, float]:
    """Return canonical base unit and numeric prefix used as a unit multiplier."""
    raw = _unit_text(value)
    if not raw:
        return "", 1.0

    scale = 1.0
    remainder = raw
    match = _LEADING_SCALE_RE.fullmatch(raw)
    if match is not None:
        parsed_scale = int(match.group(1))
        if parsed_scale <= 0:
            return "", 1.0
        scale = float(parsed_scale)
        remainder = match.group(2).strip()

    canonical = _canonical_base_unit(remainder)
    if canonical:
        return canonical, scale

    normalized = NormUnit(value)
    if not normalized:
        return "", 1.0
    return BaseUnit(normalized), scale


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


def normalized_quantity(value: object, unit: object) -> float | None:
    """Express a quantity in the base unit (for example 3 x 100m2 -> 300m2)."""
    number = positive_float(value)
    if number is None:
        return None
    base, scale = split_unit_scale(unit)
    if not base:
        return None
    return number * scale


def normalized_unit_price(value: object, unit: object) -> float | None:
    """Express a unit price per one base unit (12000/100m2 -> 120/m2)."""
    number = positive_float(value)
    if number is None:
        return None
    base, scale = split_unit_scale(unit)
    if not base or scale <= 0:
        return None
    return number / scale


def convert_unit_price(
    value: object,
    source_unit: object,
    target_unit: object,
) -> float | None:
    conversion = compatible_unit_conversion(target_unit, source_unit)
    number = positive_float(value)
    if conversion is None or number is None:
        return None
    return number * conversion.price_factor


def positive_float(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, Real):
        number = float(value)
    else:
        text = str(value).strip().replace("\xa0", "").replace(" ", "")
        if not text:
            return None
        if "," in text and "." not in text:
            text = text.replace(",", ".")
        try:
            number = float(text)
        except ValueError:
            return None
    return number if number > 0 else None


def _unit_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).casefold().replace("\u0451", "\u0435").replace("\xa0", " ")
    text = text.replace("\u00b2", "2").replace("\u00b3", "3")
    return " ".join(text.split()).strip()


def _canonical_base_unit(remainder: str) -> str:
    token = remainder.split(" ", 1)[0]
    token = _TRAILING_PUNCT_RE.sub("", token)
    token_norm = NormUnit(token)
    if not token_norm:
        return ""
    return _unit_aliases().get(token_norm, "")


@lru_cache(maxsize=1)
def _unit_aliases() -> dict[str, str]:
    try:
        payload = json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, TypeError, ValueError):
        return {}
    result: dict[str, str] = {}
    for item in payload.get("units", []):
        if not isinstance(item, dict):
            continue
        base = str(item.get("base") or "").strip()
        aliases = item.get("aliases", [])
        if not base or not isinstance(aliases, list):
            continue
        for alias in aliases:
            normalized = NormUnit(alias)
            if normalized:
                result[normalized] = base
    return result
