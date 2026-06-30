from core.catalog import CatalogEntry, CatalogRow
from core.risk import (
    REASON_NONE,
    REASON_OUT_OF_APPROVED_RANGE,
    REASON_RATIO_EXCEEDED,
    CheckPriceRisk,
    GesnException,
)


def entry(
    price: float,
    *,
    added_date_serial: float = 0,
    source_row_number: int = 1,
) -> CatalogEntry:
    row = CatalogRow(
        task_id="task-1",
        price=price,
        code="gesn01-01-001-01",
        unit="m",
        work_name="work",
    )
    return CatalogEntry(
        price=price,
        region="region-1",
        is_demolition=False,
        source_row_number=source_row_number,
        original_row=row,
        task_id="task-1",
        norm_code="GESN01-01-001-01",
        norm_unit="m",
        added_date_serial=added_date_serial,
    )


def exception(
    *,
    approved_min: float = 90,
    approved_max: float = 110,
    last_range_update_date: float = 10,
) -> GesnException:
    return GesnException(
        exception_key="m||GESN01-01-001-01||NO_DEM",
        approved_min=approved_min,
        approved_max=approved_max,
        last_range_update_date=last_range_update_date,
    )


def test_fewer_than_two_entries_never_flagged() -> None:
    result = CheckPriceRisk([entry(1000)], price_spread_limit=2.0)

    assert not result.is_flagged
    assert result.reason == REASON_NONE
    assert result.flagged_entries == []


def test_no_exception_ratio_exceeded_flags_min_and_max_entries() -> None:
    low = entry(100, source_row_number=1)
    high = entry(250, source_row_number=2)

    result = CheckPriceRisk([high, low], price_spread_limit=2.0)

    assert result.is_flagged
    assert result.reason == REASON_RATIO_EXCEEDED
    assert result.min_entry is low
    assert result.max_entry is high
    assert result.flagged_entries == [low, high]
    assert result.ratio == 2.5


def test_no_exception_ratio_below_threshold_not_flagged() -> None:
    result = CheckPriceRisk([entry(100), entry(150)], price_spread_limit=2.0)

    assert not result.is_flagged
    assert result.reason == REASON_NONE
    assert result.flagged_entries == []
    assert result.ratio == 1.5


def test_exception_new_entry_inside_approved_range_not_flagged() -> None:
    result = CheckPriceRisk(
        [entry(100, added_date_serial=11)],
        exception(),
    )

    assert not result.is_flagged
    assert result.reason == REASON_NONE
    assert result.flagged_entries == []


def test_exception_new_entry_outside_approved_range_flagged_individually() -> None:
    old_inside = entry(100, added_date_serial=11, source_row_number=1)
    new_outlier = entry(150, added_date_serial=11, source_row_number=2)

    result = CheckPriceRisk(
        [old_inside, new_outlier],
        exception(),
    )

    assert result.is_flagged
    assert result.reason == REASON_OUT_OF_APPROVED_RANGE
    assert result.flagged_entries == [new_outlier]


def test_exception_old_out_of_range_entry_is_not_reflagged() -> None:
    old_outlier = entry(150, added_date_serial=10)

    result = CheckPriceRisk(
        [old_outlier],
        exception(last_range_update_date=10),
    )

    assert not result.is_flagged
    assert result.reason == REASON_NONE
    assert result.flagged_entries == []


def test_exception_short_circuits_ratio_logic_even_when_ratio_exceeds_threshold() -> None:
    old_low = entry(10, added_date_serial=10)
    new_inside = entry(100, added_date_serial=11)

    result = CheckPriceRisk(
        [old_low, new_inside],
        exception(approved_min=90, approved_max=110, last_range_update_date=10),
        price_spread_limit=2.0,
    )

    assert not result.is_flagged
    assert result.reason == REASON_NONE
    assert result.flagged_entries == []
    assert result.ratio == 0
    assert result.min_entry is None
    assert result.max_entry is None
