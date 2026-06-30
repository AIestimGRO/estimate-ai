import pytest

from core.catalog import BuildCatalog, CatalogRow
from core.exclusions import NameExclusionRule
from core.matching import (
    REASON_EXCLUDED_BY_NAME,
    REASON_FILTERED_BY_DEMOLITION,
    REASON_INVALID_INPUT,
    REASON_MATCHED,
    REASON_NO_MATCH,
    EstimateRow,
    MatchEstimateRow,
    MatchResult,
)
from core.normalize import AnalogSearchKey


METER = "\u043c"
DEMOLITION = "\u0434\u0435\u043c\u043e\u043d\u0442\u0430\u0436"
INSTALLATION = "\u043c\u043e\u043d\u0442\u0430\u0436"
EXCLUDED = "\u0438\u0441\u043a\u043b\u044e\u0447\u0438\u0442\u044c"


class TrackingCatalog(dict):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)
        self.lookup_count = 0

    def get(self, key: object, default: object = None) -> object:
        self.lookup_count += 1
        return super().get(key, default)


def catalog_row(
    *,
    task_id: object = "task-1",
    price: object = 100.0,
    code: object = "gesn01-01-001-01",
    unit: object = METER,
    work_name: object = INSTALLATION,
    region: object = "region-1",
) -> CatalogRow:
    return CatalogRow(
        task_id=task_id,
        price=price,
        code=code,
        unit=unit,
        work_name=work_name,
        region=region,
    )


def estimate_row(
    *,
    code: object = "gesn01-01-001-01",
    unit: object = METER,
    work_name: object = INSTALLATION,
    base_price: object = 50.0,
) -> EstimateRow:
    return EstimateRow(
        code=code,
        unit=unit,
        work_name=work_name,
        base_price=base_price,
    )


def exclusion_rule() -> NameExclusionRule:
    return NameExclusionRule(
        enabled=True,
        scope="SMETA",
        match_mode="CONTAINS",
        pattern=EXCLUDED,
        group="test",
        comment="test exclusion",
    )


def prices(result: MatchResult) -> list[float]:
    return [analog.entry.price for analog in result.analogs]


def test_name_exclusion_returns_zero_and_skips_catalog_lookup() -> None:
    matching_catalog = TrackingCatalog(BuildCatalog([catalog_row()]))

    result = MatchEstimateRow(
        estimate_row(work_name=f"{EXCLUDED} item"),
        matching_catalog,
        [exclusion_rule()],
    )

    assert not result.has_analogs
    assert result.reason == REASON_EXCLUDED_BY_NAME
    assert result.analogs == []
    assert matching_catalog.lookup_count == 0


@pytest.mark.parametrize(
    "estimate",
    [
        estimate_row(code=""),
        estimate_row(unit=""),
        estimate_row(base_price=0),
        estimate_row(base_price=-1),
        estimate_row(base_price="not numeric"),
    ],
)
def test_invalid_input_returns_zero_with_invalid_reason(estimate: EstimateRow) -> None:
    result = MatchEstimateRow(estimate, BuildCatalog([catalog_row()]))

    assert not result.has_analogs
    assert result.reason == REASON_INVALID_INPUT
    assert result.analogs == []


def test_no_matching_key_returns_zero_with_no_match_reason() -> None:
    result = MatchEstimateRow(
        estimate_row(code="gesn99-99-999-99"),
        BuildCatalog([catalog_row()]),
    )

    assert not result.has_analogs
    assert result.reason == REASON_NO_MATCH


def test_demolition_row_only_gets_demolition_analogs() -> None:
    catalog = BuildCatalog(
        [
            catalog_row(price=100, work_name=INSTALLATION),
            catalog_row(price=200, work_name=DEMOLITION),
        ]
    )

    result = MatchEstimateRow(estimate_row(work_name=DEMOLITION), catalog)

    assert result.reason == REASON_MATCHED
    assert prices(result) == [200]
    assert result.analogs[0].entry.is_demolition


def test_non_demolition_row_only_gets_non_demolition_analogs() -> None:
    catalog = BuildCatalog(
        [
            catalog_row(price=100, work_name=INSTALLATION),
            catalog_row(price=200, work_name=DEMOLITION),
        ]
    )

    result = MatchEstimateRow(estimate_row(work_name=INSTALLATION), catalog)

    assert result.reason == REASON_MATCHED
    assert prices(result) == [100]
    assert not result.analogs[0].entry.is_demolition


def test_demolition_filter_can_be_disabled() -> None:
    catalog = BuildCatalog(
        [
            catalog_row(price=100, work_name=INSTALLATION),
            catalog_row(price=200, work_name=DEMOLITION),
        ]
    )

    result = MatchEstimateRow(
        estimate_row(work_name=INSTALLATION),
        catalog,
        demontazh_filter_enabled=False,
    )

    assert result.reason == REASON_MATCHED
    assert prices(result) == [100, 200]


def test_filtered_out_by_demolition_returns_reason() -> None:
    catalog = BuildCatalog([catalog_row(price=200, work_name=DEMOLITION)])

    result = MatchEstimateRow(estimate_row(work_name=INSTALLATION), catalog)

    assert not result.has_analogs
    assert result.reason == REASON_FILTERED_BY_DEMOLITION
    assert result.analogs == []


def test_column_order_follows_catalog_insertion_order_not_price_order() -> None:
    catalog = BuildCatalog(
        [
            catalog_row(task_id="task-1", price=300),
            catalog_row(task_id="task-1", price=100),
            catalog_row(task_id="task-2", price=200),
        ]
    )

    result = MatchEstimateRow(estimate_row(), catalog)

    assert result.reason == REASON_MATCHED
    assert [
        (analog.task_id, analog.price_position, analog.entry.price)
        for analog in result.analogs
    ] == [
        ("task-1", 1, 300),
        ("task-1", 2, 100),
        ("task-2", 1, 200),
    ]
    assert prices(result) != sorted(prices(result))


def test_matching_key_present_in_catalog_fixture() -> None:
    catalog = BuildCatalog([catalog_row()])
    key = AnalogSearchKey(METER, "gesn01-01-001-01")

    assert key in catalog
