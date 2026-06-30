"""Pricing helpers ported from the VBA output logic."""

from numbers import Real

NBSP = "\u00a0"


def CalculateAveragePrice(base_price: object, analog_prices: list[object]) -> float:
    """Compute value equivalent of MAX(base, IFERROR(AVERAGE(base, analogs), base))."""
    base = _parse_number(base_price)
    if base is None:
        raise ValueError("base_price must be numeric")

    numeric_values = [base]
    for analog_price in analog_prices:
        parsed = _parse_number(analog_price)
        if parsed is not None:
            numeric_values.append(parsed)

    if not numeric_values:
        average_price = base
    else:
        average_price = sum(numeric_values) / len(numeric_values)

    return max(base, average_price)


def ApplyRegionalCoefficient(price: object, coefficient_cell_value: object) -> float:
    """Apply a positive regional coefficient; invalid values behave as 1."""
    parsed_price = _parse_number(price)
    if parsed_price is None:
        raise ValueError("price must be numeric")

    coefficient = _parse_number(coefficient_cell_value)
    if coefficient is None or coefficient <= 0:
        coefficient = 1.0

    return parsed_price * coefficient


def _parse_number(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, Real):
        return float(value)

    text = str(value).strip()
    if text == "":
        return None

    text = text.replace(NBSP, "")
    text = text.replace(" ", "")
    if "," in text and "." not in text:
        text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return None
