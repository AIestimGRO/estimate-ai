from types import SimpleNamespace

from core.catalog import CatalogRow, BuildCatalog
from core.tkp_matching import (
    build_tkp_catalog_index,
    build_tkp_task_index,
    find_tkp_match_for_rnmc_analog,
)
from core.task_numbers import extract_task_numbers, normalize_single_task_number

CODE = "GESN01"
NAME = "installation of steel structures"
UNIT_M2 = "m2"


def _rnmc_entry(
    *,
    task_id="1234567/7654321",
    unit=UNIT_M2,
    quantity=300.0,
    price_original=120.0,
    work_name=NAME,
):
    row = CatalogRow(
        task_id=task_id,
        price=120.0,
        code=CODE,
        unit=unit,
        work_name=work_name,
        quantity=quantity,
        price_original=price_original,
    )
    catalog = BuildCatalog([row])
    return next(iter(next(iter(catalog.values())).values()))[0]


def _tkp_item(
    item_id,
    *,
    task_no="1234567",
    item_name=NAME,
    unit=UNIT_M2,
    qty=300.0,
    rnmc_price=120.0,
    winner=100.0,
    reserve=110.0,
):
    return SimpleNamespace(
        id=item_id,
        item_name=item_name,
        unit=unit,
        qty=qty,
        rnmc_unit_price_no_vat=rnmc_price,
        winner_unit_price_no_vat=winner,
        reserve_unit_price_no_vat=reserve,
        winner_name="winner",
        reserve_name="reserve",
        source_file_name=f"source-{item_id}.xlsx",
        task_no=task_no,
        section_name="",
        subsection_name="",
    )


def _match(rnmc, items):
    index = build_tkp_catalog_index(items)
    return find_tkp_match_for_rnmc_analog(rnmc, build_tkp_task_index(index))


def test_task_number_normalization_extracts_composite_and_hash_prefix() -> None:
    assert extract_task_numbers("1F4349964 / #6692713") == ("4349964", "6692713")
    assert normalize_single_task_number("#6692713") == "6692713"
    assert normalize_single_task_number("58420033") == ""
    assert normalize_single_task_number("3") == ""


def test_pair_match_uses_task_unit_and_scaled_quantity_before_name() -> None:
    rnmc = _rnmc_entry(unit="m2", quantity=300, price_original=120)
    items = [
        _tkp_item(1, task_no="9999999", unit="100 m2", qty=3, winner=1),
        _tkp_item(2, unit="100 m2", qty=4, winner=9000, reserve=10000),
        _tkp_item(3, unit="100 m2", qty=3, winner=10000, reserve=11000),
    ]

    match = _match(rnmc, items)

    assert match is not None
    assert match.entry.item_id == 3
    assert match.quantity_matched is True
    assert match.winner_price == 100.0
    assert match.reserve_price == 110.0


def test_pair_match_falls_back_to_task_unit_when_quantity_has_no_exact_match() -> None:
    rnmc = _rnmc_entry(quantity=300)
    match = _match(
        rnmc,
        [_tkp_item(1, qty=250, winner=90, reserve=None)],
    )

    assert match is not None
    assert match.quantity_matched is False
    assert match.winner_price == 90.0
    assert match.reserve_price is None


def test_pair_match_uses_reserve_when_winner_is_missing() -> None:
    rnmc = _rnmc_entry()
    match = _match(
        rnmc,
        [_tkp_item(1, winner=None, reserve=130)],
    )

    assert match is not None
    assert match.winner_price is None
    assert match.reserve_price == 130.0


def test_pair_match_uses_normalized_rnmc_price_to_break_equal_name_tie() -> None:
    rnmc = _rnmc_entry(unit="m2", quantity=300, price_original=120)
    items = [
        _tkp_item(
            1,
            unit="100 m2",
            qty=3,
            rnmc_price=15000,
            winner=10000,
            reserve=11000,
        ),
        _tkp_item(
            2,
            unit="100 m2",
            qty=3,
            rnmc_price=12000,
            winner=9000,
            reserve=10000,
        ),
    ]

    match = _match(rnmc, items)

    assert match is not None
    assert match.entry.item_id == 2
    assert match.rnmc_price_delta == 0.0
    assert match.winner_price == 90.0
    assert match.reserve_price == 100.0


def test_pair_match_leaves_ambiguous_different_prices_blank_without_tie_breaker() -> None:
    rnmc = _rnmc_entry(price_original=None)
    items = [
        _tkp_item(1, rnmc_price=None, winner=90, reserve=100),
        _tkp_item(2, rnmc_price=None, winner=190, reserve=200),
    ]

    assert _match(rnmc, items) is None


def test_pair_match_allows_equal_price_duplicates_without_random_price_choice() -> None:
    rnmc = _rnmc_entry(price_original=None)
    items = [
        _tkp_item(2, rnmc_price=None, winner=90, reserve=100),
        _tkp_item(1, rnmc_price=None, winner=90, reserve=100),
    ]

    match = _match(rnmc, items)

    assert match is not None
    assert match.entry.item_id == 1
    assert match.winner_price == 90.0
    assert match.reserve_price == 100.0


def test_unit_scaling_handles_kl_descriptive_units() -> None:
    from core.unit_scaling import compatible_unit_conversion, normalized_quantity, normalized_unit_price

    # Cyrillic aliases are configured in data/config/unit_scaling.json; these
    # values mirror real RNMC/KL unit spellings while code stays ASCII-only.
    conversion = compatible_unit_conversion("1 \u043c3 \u043e\u0441\u043d\u043e\u0432\u0430\u043d\u0438\u044f", "100 \u043c3 \u0433\u0440\u0443\u043d\u0442\u0430")
    assert conversion is not None
    assert conversion.base_unit == "\u043c3"
    assert normalized_quantity(3, "100 \u043c3 \u0433\u0440\u0443\u043d\u0442\u0430") == 300
    assert normalized_unit_price(12000, "100 \u043c3 \u0433\u0440\u0443\u043d\u0442\u0430") == 120


def test_pair_match_uses_semantic_scorer_as_primary_name_signal() -> None:
    rnmc = _rnmc_entry(work_name="installation of steel structures")
    items = [
        _tkp_item(1, item_name="installation of steel structures", winner=100, reserve=None),
        _tkp_item(2, item_name="assembly of structural steel frames", winner=200, reserve=None),
    ]
    index = build_tkp_catalog_index(items)

    def fake_qwen(query: str, candidates: list[str]) -> list[float]:
        assert query == "installation of steel structures"
        return [0.60, 0.94]

    match = find_tkp_match_for_rnmc_analog(
        rnmc,
        build_tkp_task_index(index),
        semantic_scorer=fake_qwen,
        semantic_model_name="Fake Qwen",
    )

    assert match is not None
    assert match.entry.item_id == 2
    assert match.score == 94.0
    assert match.semantic_score == 94.0
    assert match.match_method == "Fake Qwen"
    assert match.winner_price == 200.0
    assert match.reserve_price is None


def test_pair_match_does_not_silently_fallback_when_semantic_scorer_fails() -> None:
    from core.tkp_matching import TkpSemanticScoringError

    rnmc = _rnmc_entry()
    index = build_tkp_catalog_index([_tkp_item(1)])

    def broken_qwen(query: str, candidates: list[str]) -> list[float]:
        raise RuntimeError("model failure")

    import pytest

    with pytest.raises(TkpSemanticScoringError, match="model failure"):
        find_tkp_match_for_rnmc_analog(
            rnmc,
            build_tkp_task_index(index),
            semantic_scorer=broken_qwen,
            semantic_model_name="Fake Qwen",
        )
