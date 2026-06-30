from core.approval import (
    ApproveGesnExceptionBatch,
    ApproveGesnExceptionRange,
    ApprovedRiskLogRow,
)
from core.risk import GesnException


KEY = "m||GESN01-01-001-01||NO_DEM"
OTHER_KEY = "m||GESN02-01-001-01||NO_DEM"


def exception(
    *,
    key: str = KEY,
    approved_min: float = 100,
    approved_max: float = 200,
    last_range_update_date: float = 10,
) -> GesnException:
    return GesnException(
        exception_key=key,
        approved_min=approved_min,
        approved_max=approved_max,
        last_range_update_date=last_range_update_date,
    )


def approval_row(
    *,
    key: str = KEY,
    proposed_min: float = 100,
    proposed_max: float = 200,
    proposed_date_serial: float = 20,
) -> ApprovedRiskLogRow:
    return ApprovedRiskLogRow(
        exception_key=key,
        proposed_min=proposed_min,
        proposed_max=proposed_max,
        proposed_date_serial=proposed_date_serial,
    )


def test_first_approval_creates_new_exception_with_proposed_range() -> None:
    result = ApproveGesnExceptionRange(
        exception_key=KEY,
        proposed_min=90,
        proposed_max=150,
        proposed_date_serial=20,
    )

    assert result == GesnException(
        exception_key=KEY,
        approved_min=90,
        approved_max=150,
        last_range_update_date=20,
    )


def test_narrower_approval_does_not_shrink_existing_range() -> None:
    existing = exception(approved_min=90, approved_max=150, last_range_update_date=10)

    result = ApproveGesnExceptionRange(
        exception_key=KEY,
        proposed_min=100,
        proposed_max=140,
        proposed_date_serial=20,
        existing_exception=existing,
    )

    assert result.approved_min == 90
    assert result.approved_max == 150
    assert existing.approved_min == 90
    assert existing.approved_max == 150


def test_wider_approval_grows_both_min_and_max() -> None:
    result = ApproveGesnExceptionRange(
        exception_key=KEY,
        proposed_min=80,
        proposed_max=180,
        proposed_date_serial=20,
        existing_exception=exception(approved_min=90, approved_max=150),
    )

    assert result.approved_min == 80
    assert result.approved_max == 180


def test_last_range_update_date_always_updates_even_when_range_unchanged() -> None:
    result = ApproveGesnExceptionRange(
        exception_key=KEY,
        proposed_min=90,
        proposed_max=150,
        proposed_date_serial=30,
        existing_exception=exception(
            approved_min=90,
            approved_max=150,
            last_range_update_date=10,
        ),
    )

    assert result.approved_min == 90
    assert result.approved_max == 150
    assert result.last_range_update_date == 30


def test_batch_folds_two_approvals_for_same_key() -> None:
    result = ApproveGesnExceptionBatch(
        [
            approval_row(proposed_min=90, proposed_max=150, proposed_date_serial=20),
            approval_row(proposed_min=80, proposed_max=170, proposed_date_serial=21),
        ],
        {},
    )

    assert result[KEY] == GesnException(
        exception_key=KEY,
        approved_min=80,
        approved_max=170,
        last_range_update_date=21,
    )


def test_batch_keeps_different_keys_independent() -> None:
    result = ApproveGesnExceptionBatch(
        [
            approval_row(
                key=KEY,
                proposed_min=90,
                proposed_max=150,
                proposed_date_serial=20,
            ),
            approval_row(
                key=OTHER_KEY,
                proposed_min=10,
                proposed_max=20,
                proposed_date_serial=21,
            ),
        ],
        {KEY: exception(approved_min=100, approved_max=120, last_range_update_date=10)},
    )

    assert result[KEY] == GesnException(
        exception_key=KEY,
        approved_min=90,
        approved_max=150,
        last_range_update_date=20,
    )
    assert result[OTHER_KEY] == GesnException(
        exception_key=OTHER_KEY,
        approved_min=10,
        approved_max=20,
        last_range_update_date=21,
    )
