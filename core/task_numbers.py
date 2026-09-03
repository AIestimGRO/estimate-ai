"""Task-number normalization shared by RNMC and TKP matching/import."""

from __future__ import annotations

import re

_TASK_NUMBER_RE = re.compile(r"(?<!\d)(\d{7})(?!\d)")


def extract_task_numbers(value: object) -> tuple[str, ...]:
    """Return distinct seven-digit task numbers in first-seen order."""
    text = "" if value is None else str(value)
    seen: set[str] = set()
    result: list[str] = []
    for match in _TASK_NUMBER_RE.finditer(text):
        number = match.group(1)
        if number in seen:
            continue
        seen.add(number)
        result.append(number)
    return tuple(result)


def normalize_single_task_number(value: object) -> str:
    """Return one normalized task number, or an empty string when ambiguous/invalid."""
    numbers = extract_task_numbers(value)
    return numbers[0] if len(numbers) == 1 else ""
