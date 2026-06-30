"""Pure approval math for GESN approved price ranges."""

from dataclasses import dataclass

from core.risk import GesnException


@dataclass(frozen=True)
class ApprovedRiskLogRow:
    """Approved Price_Check_Log row data needed to widen an exception."""

    exception_key: str
    proposed_min: float
    proposed_max: float
    proposed_date_serial: float


def ApproveGesnExceptionRange(
    exception_key: str,
    proposed_min: float,
    proposed_max: float,
    proposed_date_serial: float,
    existing_exception: GesnException | None = None,
) -> GesnException:
    """Create or widen a GesnException without mutating persisted state."""
    if existing_exception is None:
        return GesnException(
            exception_key=exception_key,
            approved_min=proposed_min,
            approved_max=proposed_max,
            last_range_update_date=proposed_date_serial,
        )

    return GesnException(
        exception_key=exception_key,
        approved_min=min(existing_exception.approved_min, proposed_min),
        approved_max=max(existing_exception.approved_max, proposed_max),
        last_range_update_date=proposed_date_serial,
    )


def ApproveGesnExceptionBatch(
    approved_rows: list[ApprovedRiskLogRow],
    existing_exceptions: dict[str, GesnException] | None = None,
) -> dict[str, GesnException]:
    """Apply approved rows, folding multiple approvals for the same key."""
    result = {} if existing_exceptions is None else dict(existing_exceptions)

    for row in approved_rows:
        result[row.exception_key] = ApproveGesnExceptionRange(
            exception_key=row.exception_key,
            proposed_min=row.proposed_min,
            proposed_max=row.proposed_max,
            proposed_date_serial=row.proposed_date_serial,
            existing_exception=result.get(row.exception_key),
        )

    return result
