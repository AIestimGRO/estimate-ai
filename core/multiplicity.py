"""Deterministic work multiplicity extraction from configurable patterns."""

from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path


_CONFIG_PATH = (
    Path(__file__).resolve().parents[1] / "data" / "config" / "multiplicity.json"
)


def extract_layer_count(value: object) -> int | None:
    """Return an explicit layer count, or None when the name does not state it."""
    text = "" if value is None else " ".join(str(value).casefold().split())
    if text == "":
        return None
    for pattern in _layer_patterns():
        match = pattern.search(text)
        if match is None:
            continue
        count = int(match.group(1))
        if count > 0:
            return count
    return None


def multiplicity_is_compatible(estimate_name: object, analog_name: object) -> bool:
    """Reject only an explicit mismatch; unspecified multiplicity stays compatible."""
    estimate_count = extract_layer_count(estimate_name)
    analog_count = extract_layer_count(analog_name)
    if estimate_count is None or analog_count is None:
        return True
    return estimate_count == analog_count


@lru_cache(maxsize=1)
def _layer_patterns() -> tuple[re.Pattern[str], ...]:
    try:
        payload = json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError, TypeError):
        return ()
    patterns = payload.get("patterns", [])
    return tuple(
        re.compile(str(pattern), flags=re.IGNORECASE)
        for pattern in patterns
        if str(pattern).strip()
    )
