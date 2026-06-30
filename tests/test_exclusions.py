from core.exclusions import (
    NameExclusionRule,
    TaskColorEntry,
    is_name_excluded,
    is_task_marked,
)


SMETA = "SMETA"
CATALOG = "CATALOG"
BOTH = "BOTH"
ALL_WORDS = "ALL_WORDS"
CONTAINS = "CONTAINS"

CM = "\u0441\u043c"
CHANGED = "\u0438\u0437\u043c\u0435\u043d\u0435\u043d"
WIDTH = "\u0448\u0438\u0440\u0438\u043d\u0430"
EXTRA = "\u0434\u043e\u043f\u043e\u043b\u043d\u0438\u0442\u0435\u043b\u044c\u043d"
SHIELD = "\u0449\u0438\u0442"


def rule(
    *,
    enabled: bool = True,
    scope: str = BOTH,
    match_mode: str = ALL_WORDS,
    pattern: str = "",
) -> NameExclusionRule:
    return NameExclusionRule(
        enabled=enabled,
        scope=scope,
        match_mode=match_mode,
        pattern=pattern,
        group="test",
        comment="test rule",
    )


def test_all_words_matches_regardless_of_token_order() -> None:
    rules = [rule(pattern=f"{CM}|{CHANGED}")]
    work_name = f"{CHANGED} depth 20 {CM}"

    assert is_name_excluded(rules, SMETA, work_name)


def test_all_words_does_not_match_when_one_token_is_missing() -> None:
    rules = [rule(pattern=f"{CM}|{CHANGED}")]
    work_name = f"depth 20 {CM}"

    assert not is_name_excluded(rules, SMETA, work_name)


def test_contains_matches_substring() -> None:
    rules = [rule(match_mode=CONTAINS, pattern=EXTRA)]
    work_name = f"{EXTRA} item for {SHIELD}"

    assert is_name_excluded(rules, CATALOG, work_name)


def test_disabled_rule_never_excludes() -> None:
    rules = [rule(enabled=False, match_mode=CONTAINS, pattern=EXTRA)]
    work_name = f"{EXTRA} item"

    assert not is_name_excluded(rules, SMETA, work_name)


def test_scope_filtering() -> None:
    catalog_rule = rule(scope=CATALOG, match_mode=CONTAINS, pattern=EXTRA)
    smeta_rule = rule(scope=SMETA, match_mode=CONTAINS, pattern=WIDTH)
    both_rule = rule(scope=BOTH, match_mode=CONTAINS, pattern=SHIELD)

    assert not is_name_excluded([catalog_rule], SMETA, EXTRA)
    assert not is_name_excluded([smeta_rule], CATALOG, WIDTH)
    assert is_name_excluded([both_rule], SMETA, SHIELD)
    assert is_name_excluded([both_rule], CATALOG, SHIELD)


def test_task_color_list_does_not_affect_exclusion_results() -> None:
    color_entries = [
        TaskColorEntry(
            enabled=True,
            task_number=" 12 34 ",
            reason="highlight",
            comment="metadata only",
        )
    ]

    assert is_task_marked(color_entries, "1234")
    assert not is_name_excluded([], SMETA, "1234")


def test_invalid_scope_silently_does_not_exclude() -> None:
    rules = [rule(scope="SMETAA", match_mode=CONTAINS, pattern=EXTRA)]

    assert not is_name_excluded(rules, SMETA, EXTRA)


def test_invalid_match_mode_silently_does_not_exclude() -> None:
    rules = [rule(scope=SMETA, match_mode="CONTAINSS", pattern=EXTRA)]

    assert not is_name_excluded(rules, SMETA, EXTRA)
