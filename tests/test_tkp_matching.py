from dataclasses import dataclass

from core.tkp_matching import (
    DEFAULT_MIN_SCORE,
    build_tkp_catalog_index,
    find_best_tkp_matches,
    leading_action_word,
    normalize_for_matching,
    same_action_group,
    score_names,
)


@dataclass(frozen=True)
class _FakeTkpItem:
    id: int
    item_name: str
    unit: str
    winner_unit_price_no_vat: float | None
    winner_name: str
    source_file_name: str
    task_no: str


ITEMS = [
    _FakeTkpItem(
        id=1,
        item_name="\u0420\u0430\u0437\u0431\u043e\u0440\u043a\u0430 \u0442\u0435\u043f\u043b\u043e\u0432\u043e\u0439 \u0438\u0437\u043e\u043b\u044f\u0446\u0438\u0438: \u0438\u0437 \u0432\u0430\u0442\u044b \u043c\u0438\u043d\u0435\u0440\u0430\u043b\u044c\u043d\u043e\u0439",
        unit="100 \u043c2",
        winner_unit_price_no_vat=26547.77,
        winner_name="\u041e\u041e\u041e \u041f\u0440\u0438\u043c\u0435\u0440",
        source_file_name="file1.xlsx",
        task_no="111",
    ),
    _FakeTkpItem(
        id=2,
        item_name="\u041f\u043e\u043a\u0440\u044b\u0442\u0438\u0435 \u043f\u043e\u0432\u0435\u0440\u0445\u043d\u043e\u0441\u0442\u0438 \u0438\u0437\u043e\u043b\u044f\u0446\u0438\u0438 \u0442\u0440\u0443\u0431\u043e\u043f\u0440\u043e\u0432\u043e\u0434\u043e\u0432: \u0441\u0442\u0430\u043b\u044c\u044e \u043e\u0446\u0438\u043d\u043a\u043e\u0432\u0430\u043d\u043d\u043e\u0439",
        unit="100 \u043c2",
        winner_unit_price_no_vat=217954.85,
        winner_name="\u041e\u041e\u041e \u041f\u0440\u0438\u043c\u0435\u0440",
        source_file_name="file1.xlsx",
        task_no="111",
    ),
    _FakeTkpItem(
        id=3,
        item_name="\u0423\u0441\u0442\u0440\u043e\u0439\u0441\u0442\u0432\u043e \u0442\u0435\u043f\u043b\u043e\u0432\u043e\u0439 \u0438\u0437\u043e\u043b\u044f\u0446\u0438\u0438 \u0442\u0440\u0443\u0431\u043e\u043f\u0440\u043e\u0432\u043e\u0434\u043e\u0432",
        unit="\u043c3",
        winner_unit_price_no_vat=23691.16,
        winner_name="\u041e\u041e\u041e \u0412\u0442\u043e\u0440\u043e\u0439",
        source_file_name="file2.xlsx",
        task_no="112",
    ),
    _FakeTkpItem(
        id=4,
        item_name="\u041c\u043e\u043d\u0442\u0430\u0436 \u0441\u0438\u043b\u043e\u0432\u043e\u0433\u043e \u044d\u043b\u0435\u043a\u0442\u0440\u043e\u043e\u0431\u043e\u0440\u0443\u0434\u043e\u0432\u0430\u043d\u0438\u044f",
        unit="\u0448\u0442.",
        winner_unit_price_no_vat=20000.0,
        winner_name="\u041e\u041e\u041e \u0422\u0440\u0435\u0442\u0438\u0439",
        source_file_name="file3.xlsx",
        task_no="113",
    ),
]


def test_normalize_for_matching_strips_punctuation_and_stopwords() -> None:
    text, tokens = normalize_for_matching("\u0420\u0430\u0437\u0431\u043e\u0440\u043a\u0430: \u0438\u0437 \u0432\u0430\u0442\u044b, \u0438 \u0442\u043a\u0430\u043d\u0435\u0439")
    assert "\u0438\u0437" not in tokens  # stopword
    assert "\u0438" not in tokens  # stopword
    assert "\u0440\u0430\u0437\u0431\u043e\u0440\u043a\u0430" in tokens
    assert "\u0432\u0430\u0442\u044b" in tokens


def test_normalize_for_matching_handles_none_and_empty() -> None:
    assert normalize_for_matching(None) == ("", frozenset())
    assert normalize_for_matching("") == ("", frozenset())


def test_score_names_identical_text_is_100() -> None:
    text, tokens = normalize_for_matching("\u0420\u0430\u0437\u0431\u043e\u0440\u043a\u0430 \u0442\u0435\u043f\u043b\u043e\u0432\u043e\u0439 \u0438\u0437\u043e\u043b\u044f\u0446\u0438\u0438")
    assert score_names(text, tokens, text, tokens) == 100.0


def test_score_names_unrelated_text_is_low() -> None:
    text_a, tokens_a = normalize_for_matching("\u0420\u0430\u0437\u0431\u043e\u0440\u043a\u0430 \u0442\u0435\u043f\u043b\u043e\u0432\u043e\u0439 \u0438\u0437\u043e\u043b\u044f\u0446\u0438\u0438")
    text_b, tokens_b = normalize_for_matching("\u0414\u043e\u0441\u0442\u0430\u0432\u043a\u0430 \u043f\u0438\u0446\u0446\u044b \u043d\u0430 \u0434\u043e\u043c")
    assert score_names(text_a, tokens_a, text_b, tokens_b) < 30.0


def test_find_best_tkp_matches_ranks_exact_name_first() -> None:
    index = build_tkp_catalog_index(ITEMS)

    matches = find_best_tkp_matches(
        "\u0420\u0430\u0437\u0431\u043e\u0440\u043a\u0430 \u0442\u0435\u043f\u043b\u043e\u0432\u043e\u0439 \u0438\u0437\u043e\u043b\u044f\u0446\u0438\u0438: \u0438\u0437 \u0432\u0430\u0442\u044b \u043c\u0438\u043d\u0435\u0440\u0430\u043b\u044c\u043d\u043e\u0439",
        "100 \u043c2",
        index,
        min_score=0,
        limit=3,
    )

    assert matches[0].entry.item_id == 1
    assert matches[0].score == 100.0
    # the unrelated electrical-work item should not even be in the top 3
    assert all(match.entry.item_id != 4 for match in matches)


def test_find_best_tkp_matches_ranks_synonym_reasonably() -> None:
    index = build_tkp_catalog_index(ITEMS)

    # "\u0414\u0435\u043c\u043e\u043d\u0442\u0430\u0436" (demolition) instead of
    # "\u0420\u0430\u0437\u0431\u043e\u0440\u043a\u0430" (disassembly) - a
    # realistic estimate-vs-TKP wording difference for the same real work.
    matches = find_best_tkp_matches(
        "\u0414\u0435\u043c\u043e\u043d\u0442\u0430\u0436 \u0442\u0435\u043f\u043b\u043e\u0432\u043e\u0439 \u0438\u0437\u043e\u043b\u044f\u0446\u0438\u0438 \u0438\u0437 \u043c\u0438\u043d\u0435\u0440\u0430\u043b\u044c\u043d\u043e\u0439 \u0432\u0430\u0442\u044b",
        "100 \u043c2",
        index,
        min_score=0,
        limit=1,
    )

    assert matches[0].entry.item_id == 1
    assert matches[0].score >= DEFAULT_MIN_SCORE


def test_find_best_tkp_matches_respects_min_score(monkeypatch=None) -> None:
    index = build_tkp_catalog_index(ITEMS)

    matches = find_best_tkp_matches(
        "\u041f\u043e\u043b\u043d\u043e\u0441\u0442\u044c\u044e \u043d\u0435\u0441\u0432\u044f\u0437\u0430\u043d\u043d\u044b\u0439 \u0437\u0430\u043f\u0440\u043e\u0441 \u043f\u0440\u043e \u043a\u043e\u0442\u0438\u043a\u043e\u0432",
        "\u0448\u0442",
        index,
        min_score=DEFAULT_MIN_SCORE,
    )

    assert matches == []


def test_find_best_tkp_matches_unit_bonus_breaks_ties() -> None:
    # Two items with the same normalized name but different units; a query
    # matching one unit should score that one at least as high as the other.
    tied_items = [
        _FakeTkpItem(id=10, item_name="\u041c\u043e\u043d\u0442\u0430\u0436 \u043a\u0430\u0431\u0435\u043b\u044f", unit="\u043c", winner_unit_price_no_vat=10.0, winner_name="A", source_file_name="f.xlsx", task_no="1"),
        _FakeTkpItem(id=11, item_name="\u041c\u043e\u043d\u0442\u0430\u0436 \u043a\u0430\u0431\u0435\u043b\u044f", unit="\u0448\u0442", winner_unit_price_no_vat=20.0, winner_name="B", source_file_name="f.xlsx", task_no="1"),
    ]
    index = build_tkp_catalog_index(tied_items)

    matches = find_best_tkp_matches("\u041c\u043e\u043d\u0442\u0430\u0436 \u043a\u0430\u0431\u0435\u043b\u044f", "\u0448\u0442", index, min_score=0, limit=2)

    # both display as a capped 100.0, but the unit-matching item must still
    # rank first - the bonus must not get lost to the display-side clamp.
    assert [match.entry.item_id for match in matches] == [11, 10]
    assert matches[0].score == 100.0
    assert matches[1].score == 100.0


def test_build_tkp_catalog_index_skips_items_with_empty_name() -> None:
    blank_item = _FakeTkpItem(id=99, item_name="   ", unit="\u0448\u0442", winner_unit_price_no_vat=1.0, winner_name="A", source_file_name="f.xlsx", task_no="1")
    index = build_tkp_catalog_index([*ITEMS, blank_item])

    assert all(entry.item_id != 99 for entry in index)
    assert len(index) == len(ITEMS)


def test_leading_action_word_returns_first_content_word() -> None:
    text, _ = normalize_for_matching("\u0423\u043a\u043b\u0430\u0434\u043a\u0430 \u0442\u0440\u0443\u0431\u043e\u043f\u0440\u043e\u0432\u043e\u0434\u043e\u0432 \u0438\u0437 \u043f\u043e\u043b\u0438\u044d\u0442\u0438\u043b\u0435\u043d\u0430")
    assert leading_action_word(text) == "\u0443\u043a\u043b\u0430\u0434\u043a\u0430"


def test_same_action_group_recognizes_demontage_synonyms() -> None:
    assert same_action_group("\u0434\u0435\u043c\u043e\u043d\u0442\u0430\u0436", "\u0440\u0430\u0437\u0431\u043e\u0440\u043a\u0430") is True
    assert same_action_group("\u0443\u0441\u0442\u0440\u043e\u0439\u0441\u0442\u0432\u043e", "\u043c\u043e\u043d\u0442\u0430\u0436") is True


def test_same_action_group_rejects_unrelated_operations() -> None:
    assert same_action_group("\u0443\u043a\u043b\u0430\u0434\u043a\u0430", "\u043f\u0440\u043e\u0442\u0430\u0441\u043a\u0438\u0432\u0430\u043d\u0438\u0435") is False


def test_same_action_group_treats_missing_word_as_no_penalty() -> None:
    assert same_action_group("", "\u043c\u043e\u043d\u0442\u0430\u0436") is True
    assert same_action_group("\u043c\u043e\u043d\u0442\u0430\u0436", "") is True


def test_find_best_tkp_matches_penalizes_different_operation_same_object() -> None:
    # Real false-positive class found on real data: "Укладка" (laying, open
    # trench) vs "Протаскивание в футляр" (pulling through a casing) share
    # almost every noun (pipe material + diameter) but are different
    # construction operations with different pricing.
    items = [
        _FakeTkpItem(
            id=20,
            item_name="\u041f\u0440\u043e\u0442\u0430\u0441\u043a\u0438\u0432\u0430\u043d\u0438\u0435 \u0432 \u0444\u0443\u0442\u043b\u044f\u0440 \u043f\u043e\u043b\u0438\u044d\u0442\u0438\u043b\u0435\u043d\u043e\u0432\u044b\u0445 \u0442\u0440\u0443\u0431 \u0434\u0438\u0430\u043c\u0435\u0442\u0440\u043e\u043c: 110 \u043c\u043c",
            unit="\u043c",
            winner_unit_price_no_vat=500.0,
            winner_name="A",
            source_file_name="f.xlsx",
            task_no="1",
        ),
    ]
    index = build_tkp_catalog_index(items)

    matches = find_best_tkp_matches(
        "\u0423\u043a\u043b\u0430\u0434\u043a\u0430 \u0442\u0440\u0443\u0431\u043e\u043f\u0440\u043e\u0432\u043e\u0434\u043e\u0432 \u0438\u0437 \u043f\u043e\u043b\u0438\u044d\u0442\u0438\u043b\u0435\u043d\u043e\u0432\u044b\u0445 \u0442\u0440\u0443\u0431 \u0434\u0438\u0430\u043c\u0435\u0442\u0440\u043e\u043c: 110 \u043c\u043c",
        "\u043c",
        index,
        min_score=DEFAULT_MIN_SCORE,
    )

    assert matches == []  # below threshold once the penalty is applied


def test_find_best_tkp_matches_still_finds_same_group_synonym() -> None:
    items = [
        _FakeTkpItem(
            id=21,
            item_name="\u0420\u0430\u0437\u0431\u043e\u0440\u043a\u0430 \u0442\u0435\u043f\u043b\u043e\u0432\u043e\u0439 \u0438\u0437\u043e\u043b\u044f\u0446\u0438\u0438: \u0438\u0437 \u0432\u0430\u0442\u044b \u043c\u0438\u043d\u0435\u0440\u0430\u043b\u044c\u043d\u043e\u0439",
            unit="100 \u043c2",
            winner_unit_price_no_vat=26547.77,
            winner_name="A",
            source_file_name="f.xlsx",
            task_no="1",
        ),
    ]
    index = build_tkp_catalog_index(items)

    matches = find_best_tkp_matches(
        "\u0414\u0435\u043c\u043e\u043d\u0442\u0430\u0436 \u0442\u0435\u043f\u043b\u043e\u0432\u043e\u0439 \u0438\u0437\u043e\u043b\u044f\u0446\u0438\u0438 \u0438\u0437 \u043c\u0438\u043d\u0435\u0440\u0430\u043b\u044c\u043d\u043e\u0439 \u0432\u0430\u0442\u044b",
        "100 \u043c2",
        index,
        min_score=DEFAULT_MIN_SCORE,
    )

    assert len(matches) == 1
    assert matches[0].entry.item_id == 21
