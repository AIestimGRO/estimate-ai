"""Human-readable inspection of one matching run on real files.

Not the production CLI (that comes later) - this is a quick way to eyeball
the structured result of app.services.run_matching against the original
macro output, one estimate row per line.

Usage:
    python scripts/inspect_run.py <catalog.xlsx> <estimate.xlsx>
        [--coef 1.0] [--spread 2.0] [--no-demontazh-filter]
"""

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.services.run_matching import EstimateRowResult, run_matching_from_files
from core.risk import DEFAULT_PRICE_SPREAD_LIMIT


def _format_price(value: float | None) -> str:
    return "-" if value is None else f"{value:.2f}"


def _format_row(row: EstimateRowResult) -> str:
    prices = ", ".join(f"{analog.entry.price:.2f}" for analog in row.analogs)
    risk = row.risk_result.reason if row.risk_result.is_flagged else "-"
    return (
        f"[{row.row_index:>4}] "
        f"status={row.status:<22} "
        f"analogs={len(row.analogs):>2} "
        f"section={row.section_code or '-':<3} "
        f"recommended={_format_price(row.recommended_price):>12} "
        f"risk={risk:<20} "
        f"prices=[{prices}]"
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Inspect one matching run.")
    parser.add_argument("catalog", type=Path, help="path to the catalog .xlsx")
    parser.add_argument("estimate", type=Path, help="path to the estimate .xlsx")
    parser.add_argument("--coef", type=float, default=1.0, help="regional coefficient")
    parser.add_argument(
        "--spread",
        type=float,
        default=DEFAULT_PRICE_SPREAD_LIMIT,
        help="price spread limit for ratio risk",
    )
    parser.add_argument(
        "--no-demontazh-filter",
        action="store_true",
        help="disable the demolition filter",
    )
    args = parser.parse_args(argv)

    result = run_matching_from_files(
        args.catalog,
        args.estimate,
        demontazh_filter_enabled=not args.no_demontazh_filter,
        price_spread_limit=args.spread,
        regional_coefficient=args.coef,
    )

    print("=== Matching run summary ===")
    print(f"catalog keys      : {result.catalog_key_count}")
    print(f"estimate rows     : {len(result.rows)}")
    print(f"rows with analogs : {result.matched_row_count}")
    print(f"rows flagged risk : {result.flagged_row_count}")
    print()
    print("=== Per-row result ===")
    for row in result.rows:
        print(_format_row(row))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
