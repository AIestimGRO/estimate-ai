from core.normalize import AnalogSearchKey, HasDemontazh, NormCode, NormUnit


CYRILLIC_KR = "\u041a\u0420"
CYRILLIC_M = "\u043c"
CYRILLIC_YO_UNIT = "\u043c\u00b2 \u0451\u043c\u043a."
CYRILLIC_YE_UNIT = "\u043c2\u0435\u043c\u043a"
DEMOLITION = "\u0434\u0435\u043c\u043e\u043d\u0442\u0430\u0436"
DEMOLITION_ADJ = "\u0434\u0435\u043c\u043e\u043d\u0442\u0430\u0436\u043d\u044b\u0435"
INSTALLATION = "\u043c\u043e\u043d\u0442\u0430\u0436"


def test_norm_code_strips_kr_suffix() -> None:
    assert NormCode(f"gesn01-01-001-01/{CYRILLIC_KR}") == "GESN01-01-001-01"


def test_norm_code_cleans_whitespace_tabs_line_breaks_and_nbsp() -> None:
    raw = "\t gesn01-01-001-01 \r\n / \u00a0 01 "

    assert NormCode(raw) == "GESN01-01-001-01/01"


def test_norm_unit_100_m_and_m_share_matching_key() -> None:
    assert NormUnit(f"100 {CYRILLIC_M}") == f"100{CYRILLIC_M}"
    assert NormUnit(CYRILLIC_M) == CYRILLIC_M
    assert AnalogSearchKey(f"100 {CYRILLIC_M}", "gesn01") == AnalogSearchKey(
        CYRILLIC_M,
        "gesn01",
    )


def test_base_unit_strips_leading_quantity_prefix() -> None:
    from core.normalize import BaseUnit

    assert BaseUnit(f"100 {CYRILLIC_M}") == CYRILLIC_M
    assert BaseUnit("100\u043c2") == "\u043c2"
    assert BaseUnit(CYRILLIC_M) == CYRILLIC_M


def test_norm_unit_handles_yo_superscript_and_punctuation() -> None:
    assert NormUnit(CYRILLIC_YO_UNIT) == CYRILLIC_YE_UNIT
    assert NormUnit(f" {CYRILLIC_M} . , ^ ") == CYRILLIC_M


def test_has_demontazh_matches_demolition_root_only() -> None:
    assert HasDemontazh(DEMOLITION)
    assert HasDemontazh(DEMOLITION_ADJ)
    assert not HasDemontazh(INSTALLATION)


def test_analog_search_key_returns_empty_when_code_or_unit_is_empty() -> None:
    assert AnalogSearchKey("", "gesn01") == ""
    assert AnalogSearchKey(CYRILLIC_M, "") == ""
    assert AnalogSearchKey(None, "gesn01") == ""
    assert AnalogSearchKey(CYRILLIC_M, None) == ""
