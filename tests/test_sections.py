from core.sections import (
    DEMOLITION_PRIORITY_PREFIXES,
    GESN,
    GESNM,
    GESNP,
    GESNR,
    BuildSectionDict,
    GESnPrefix,
    ResolveSectionCode,
)


def test_non_special_prefix_returns_table_value_regardless_of_demolition_flag() -> None:
    code = f"{GESN}26-01-001-01"

    assert ResolveSectionCode(code, is_demolition=False) == "07"
    assert ResolveSectionCode(code, is_demolition=True) == "07"


def test_demolition_priority_prefix_with_demolition_returns_08() -> None:
    assert ResolveSectionCode(f"{GESN}09-01-001-01", is_demolition=True) == "08"


def test_demolition_priority_prefix_without_demolition_prefers_non_08_table_value() -> None:
    assert ResolveSectionCode(f"{GESN}09-01-001-01", is_demolition=False) == "04"


def test_unknown_prefix_returns_empty_string() -> None:
    assert ResolveSectionCode(f"{GESN}99-01-001-01", is_demolition=False) == ""


def test_real_lookup_table_entries_resolve_correctly() -> None:
    assert ResolveSectionCode(f"{GESN}01-01-001-01", is_demolition=False) == "01"
    assert ResolveSectionCode(f"{GESNM}38-01-001-01", is_demolition=False) == "04"
    assert ResolveSectionCode(f"{GESNP}03-01-001-01", is_demolition=False) == "12"
    assert ResolveSectionCode(f"{GESNR}51-01-001-01", is_demolition=False) == "09"


def test_table_contains_vba_demolition_priority_prefixes() -> None:
    assert DEMOLITION_PRIORITY_PREFIXES == {
        f"{GESN}09",
        f"{GESN}27",
        f"{GESN}28",
        f"{GESN}46",
        f"{GESNR}67",
    }


def test_build_section_dict_exposes_lookup_table_as_data() -> None:
    section_dict = BuildSectionDict()

    assert section_dict[f"{GESN}09"] == "04"
    assert section_dict[f"{GESN}27"] == "05"
    assert section_dict[f"{GESN}28"] == "05"
    assert section_dict[f"{GESN}46"] == "08"
    assert section_dict[f"{GESNR}67"] == "08"


def test_gesn_prefix_extracts_letter_suffix() -> None:
    assert GESnPrefix(f"{GESNM}38-01-001-01") == f"{GESNM}38"


def test_gesn_prefix_extracts_plain_prefix_without_letter_suffix() -> None:
    assert GESnPrefix(f"{GESN}26-01-001-01") == f"{GESN}26"


def test_gesn_prefix_returns_partial_prefix_when_digits_are_missing() -> None:
    assert GESnPrefix(f"{GESN}M") == GESN
    assert GESnPrefix(GESNM) == GESNM
