import pytest

from core.pricing import ApplyRegionalCoefficient, CalculateAveragePrice


def test_calculate_average_price_includes_base_price_in_average() -> None:
    result = CalculateAveragePrice(100, [200, 300])

    assert result == 200


def test_empty_analog_prices_falls_back_to_base_price() -> None:
    assert CalculateAveragePrice(100, []) == 100


def test_non_numeric_analog_prices_fall_back_to_base_price() -> None:
    assert CalculateAveragePrice(100, [None, "", "not numeric"]) == 100


def test_average_price_is_never_below_base_price() -> None:
    # Because the VBA formula wraps AVERAGE(base, analogs) in MAX(base, ...),
    # lower analog prices cannot pull the final value below the base price.
    assert CalculateAveragePrice(100, [10, 20]) == 100


def test_apply_regional_coefficient_valid_value_scales_price() -> None:
    assert ApplyRegionalCoefficient(100, 1.15) == pytest.approx(115)


@pytest.mark.parametrize("coefficient", [None, "", 0, -1, "not numeric"])
def test_apply_regional_coefficient_invalid_values_leave_price_unchanged(
    coefficient: object,
) -> None:
    assert ApplyRegionalCoefficient(100, coefficient) == 100


def test_regional_coefficient_applies_to_logged_min_max_and_new_values() -> None:
    coefficient = 1.15

    assert ApplyRegionalCoefficient(100, coefficient) == pytest.approx(115)
    assert ApplyRegionalCoefficient(250, coefficient) == pytest.approx(287.5)
    assert ApplyRegionalCoefficient(300, coefficient) == pytest.approx(345)
