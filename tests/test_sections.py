from core.sections import (
    DEMOLITION_PRIORITY_PREFIXES,
    GESN,
    GESNM,
    GESNP,
    GESNR,
    BuildSectionDict,
    GESnPrefix,
    CanonicalGesnCode,
    HasExactThirdLevelSection,
    NormalizeSectionValue,
    ResolveExactThirdLevelSection,
    ResolveSectionCode,
    ResolveSectionCodeWithSource,
    SECTION_SOURCE_MANUAL,
)


def test_manual_mapping_has_top_priority() -> None:
    code = f"{GESN}20-06-021-01"
    manual_map = {CanonicalGesnCode(code): "11"}

    section, source = ResolveSectionCodeWithSource(
        code,
        is_demolition=False,
        manual_section_mappings=manual_map,
    )

    assert section == "11"
    assert source == SECTION_SOURCE_MANUAL
    assert ResolveSectionCode(
        code,
        is_demolition=False,
        manual_section_mappings=manual_map,
    ) == "11"


def test_normalize_section_value_returns_two_digits() -> None:
    assert NormalizeSectionValue("1") == "01"
    assert NormalizeSectionValue("11") == "11"
    assert NormalizeSectionValue("0") == ""



def test_exact_third_level_source_can_be_detected() -> None:
    code = f"{GESNP}01-11-011-01"

    assert ResolveExactThirdLevelSection(code) == "17"
    assert HasExactThirdLevelSection(code) is True
    assert HasExactThirdLevelSection(f"{GESN}09-99-999-99") is False


def test_exact_third_level_mapping_has_priority() -> None:
    assert ResolveSectionCode(f"{GESNP}01-11-011-01", is_demolition=False) == "17"


def test_exact_third_level_mapping_overrides_demolition_fallback() -> None:
    assert ResolveSectionCode(f"{GESNR}67-01-003-01", is_demolition=True) == "09"


def test_exact_mapping_accepts_fer_and_ter_prefixes() -> None:
    assert ResolveSectionCode("\u0424\u0415\u0420\u043f01-11-011-01", is_demolition=False) == "17"
    assert ResolveSectionCode("\u0422\u0415\u042001-01-003-07", is_demolition=False) == "01"


def test_commissioning_prefix_uses_section_17_between_exact_and_fallback() -> None:
    assert ResolveExactThirdLevelSection(f"{GESNP}03-01-001-01") == ""
    assert ResolveSectionCode(f"{GESNP}03-01-001-01", is_demolition=False) == "17"
    assert ResolveSectionCode(f"{GESNP}02-99-999-99", is_demolition=True) == "17"


def test_commissioning_prefix_accepts_fer_and_ter_prefixes() -> None:
    assert ResolveSectionCode("\u0424\u0415\u0420\u043f02-99-999-99", is_demolition=False) == "17"
    assert ResolveSectionCode("\u0422\u0415\u0420\u043f02-99-999-99", is_demolition=False) == "17"


def test_non_special_prefix_returns_table_value_regardless_of_demolition_flag() -> None:
    code = f"{GESN}26-99-999-99"

    assert ResolveSectionCode(code, is_demolition=False) == "07"
    assert ResolveSectionCode(code, is_demolition=True) == "07"


def test_ambiguous_fallback_prefix_with_demolition_returns_demolition_section() -> None:
    assert ResolveSectionCode(f"{GESN}09-99-999-99", is_demolition=True) == "08"


def test_ambiguous_fallback_prefix_without_demolition_prefers_non_demolition_section() -> None:
    assert ResolveSectionCode(f"{GESN}09-99-999-99", is_demolition=False) == "04"
    assert ResolveSectionCode(f"{GESN}46-99-999-99", is_demolition=False) == "09"


def test_unknown_prefix_returns_empty_string() -> None:
    assert ResolveSectionCode(f"{GESN}99-01-001-01", is_demolition=False) == ""


def test_real_lookup_table_entries_resolve_correctly() -> None:
    assert ResolveSectionCode(f"{GESN}01-01-001-01", is_demolition=False) == "01"
    assert ResolveSectionCode(f"{GESNM}38-01-001-01", is_demolition=False) == "04"
    assert ResolveSectionCode(f"{GESNP}03-01-001-01", is_demolition=False) == "17"
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
    assert section_dict[f"{GESN}46"] == "09"
    assert section_dict[f"{GESNR}67"] == "08"


def test_gesn_prefix_extracts_letter_suffix() -> None:
    assert GESnPrefix(f"{GESNM}38-01-001-01") == f"{GESNM}38"


def test_gesn_prefix_extracts_plain_prefix_without_letter_suffix() -> None:
    assert GESnPrefix(f"{GESN}26-01-001-01") == f"{GESN}26"


def test_gesn_prefix_accepts_fer_and_ter_prefixes() -> None:
    assert GESnPrefix("\u0424\u0415\u0420\u043c08-02-412-02") == f"{GESNM}08"
    assert GESnPrefix("\u0422\u0415\u042001-01-003-07") == f"{GESN}01"


def test_gesn_prefix_returns_partial_prefix_when_digits_are_missing() -> None:
    assert GESnPrefix(f"{GESN}M") == GESN
    assert GESnPrefix(GESNM) == GESNM
