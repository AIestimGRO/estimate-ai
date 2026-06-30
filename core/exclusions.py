"""Name exclusion rules ported from the VBA macro."""

from dataclasses import dataclass


NBSP = "\u00a0"
VALID_SCOPES = {"SMETA", "CATALOG", "BOTH"}
VALID_MATCH_MODES = {"CONTAINS", "ALL_WORDS"}


@dataclass(frozen=True)
class NameExclusionRule:
    """Config row from Name_Exclusions, DOMAIN_RULES.md section 7."""

    enabled: bool
    scope: str
    match_mode: str
    pattern: str
    group: str = ""
    comment: str = ""


@dataclass(frozen=True)
class TaskColorEntry:
    """Task highlight metadata; it does not exclude or block matching."""

    enabled: bool
    task_number: str
    reason: str = ""
    comment: str = ""


def normalize_name_for_rules(value: object) -> str:
    """Port of NormalizeNameForRules from Module7."""
    text = "" if value is None else str(value)
    text = text.lower()
    text = text.replace("\r", " ")
    text = text.replace("\n", " ")
    text = text.replace("\t", " ")
    text = text.replace(NBSP, " ")

    while "  " in text:
        text = text.replace("  ", " ")

    return text.strip()


def normalize_task_key(value: object) -> str:
    """Port of NormalizeTaskKey from Module7."""
    text = "" if value is None else str(value)
    text = text.replace("\r", "")
    text = text.replace("\n", "")
    text = text.replace("\t", "")
    text = text.replace(NBSP, "")
    text = text.replace(" ", "")
    return text.strip()


def is_name_excluded(
    rules: list[NameExclusionRule],
    scope: str,
    work_name: object,
) -> bool:
    """Return whether work_name is excluded for the given scope."""
    text_key = normalize_name_for_rules(work_name)
    if text_key == "":
        return False

    scope_key = _normalize_scope(scope)

    for rule in rules:
        if not rule.enabled:
            continue

        rule_scope = _normalize_scope(rule.scope)
        if rule_scope != "BOTH" and rule_scope != scope_key:
            continue

        pattern = normalize_name_for_rules(rule.pattern)
        if pattern == "":
            continue

        mode_key = _normalize_match_mode(rule.match_mode)
        if mode_key not in VALID_MATCH_MODES:
            continue

        if _rule_matches(text_key, mode_key, pattern):
            return True

    return False


def is_task_marked(
    color_entries: list[TaskColorEntry],
    task_number: object,
) -> bool:
    """Return highlight metadata state; this never affects exclusions."""
    task_key = normalize_task_key(task_number)
    if task_key == "":
        return False

    for entry in color_entries:
        if entry.enabled and normalize_task_key(entry.task_number) == task_key:
            return True

    return False


def _normalize_scope(scope: str) -> str:
    scope_key = str(scope).strip().upper()
    if scope_key == "":
        return "BOTH"
    if scope_key in VALID_SCOPES:
        return scope_key
    return scope_key


def _normalize_match_mode(match_mode: str) -> str:
    mode_key = str(match_mode).strip().upper()
    if mode_key == "":
        return "ALL_WORDS"
    if mode_key in VALID_MATCH_MODES:
        return mode_key
    return mode_key


def _rule_matches(text_key: str, mode_key: str, pattern: str) -> bool:
    if mode_key == "CONTAINS":
        return pattern in text_key
    if mode_key == "ALL_WORDS":
        return _match_all_tokens(text_key, pattern)
    return False


def _match_all_tokens(text_key: str, pattern: str) -> bool:
    for token in pattern.split("|"):
        token_key = token.strip()
        if token_key != "" and token_key not in text_key:
            return False
    return True
