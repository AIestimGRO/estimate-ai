"""Section-code resolution for estimate rows."""

from functools import lru_cache
import json
from pathlib import Path

from core.normalize import NormCode


GESN = "\u0413\u042d\u0421\u041d"
FER = "\u0424\u0415\u0420"
TER = "\u0422\u0415\u0420"
GESNM = f"{GESN}\u041c"
GESNP = f"{GESN}\u041f"
GESNR = f"{GESN}\u0420"

SECTION_SOURCE_MANUAL = "manual"
SECTION_SOURCE_THIRD_LEVEL = "third_level"
SECTION_SOURCE_COMMISSIONING = "commissioning"
SECTION_SOURCE_FALLBACK = "fallback"

DATA_DIR = Path(__file__).with_name("data")
THIRD_LEVEL_SECTION_MAP = DATA_DIR / "ekr_third_level_sections.json"

DEMOLITION_PRIORITY_PREFIXES = {
    f"{GESN}09",
    f"{GESN}27",
    f"{GESN}28",
    f"{GESN}46",
    f"{GESNR}67",
}

FALLBACK_SECTION_CHOICES: dict[str, tuple[tuple[str, bool], ...]] = {
    f"{GESN}01": (("01", False),),
    f"{GESN}04": (("01", False),),
    f"{GESN}05": (("02", False),),
    f"{GESN}06": (("02", False),),
    f"{GESN}07": (("02", False),),
    f"{GESN}08": (("03", False),),
    f"{GESN}10": (("03", False),),
    f"{GESN}11": (("03", False),),
    f"{GESN}12": (("03", False),),
    f"{GESN}15": (("03", False),),
    f"{GESN}09": (("04", False), ("08", True)),
    f"{GESN}39": (("04", False),),
    f"{GESNM}38": (("04", False),),
    f"{GESN}27": (("05", False), ("08", True)),
    f"{GESN}28": (("05", False), ("08", True)),
    f"{GESN}47": (("05", False),),
    f"{GESNM}20": (("05", False),),
    f"{GESNM}03": (("06", False),),
    f"{GESNM}06": (("06", False),),
    f"{GESNM}07": (("06", False),),
    f"{GESNM}13": (("06", False),),
    f"{GESNM}18": (("06", False),),
    f"{GESNM}19": (("06", False),),
    f"{GESNM}22": (("06", False),),
    f"{GESNM}37": (("06", False),),
    f"{GESNP}07": (("06", False),),
    f"{GESN}13": (("07", False),),
    f"{GESN}26": (("07", False),),
    f"{GESN}45": (("07", False),),
    f"{GESN}46": (("09", False), ("08", True)),
    f"{GESNR}67": (("08", True),),
    f"{GESNR}51": (("09", False),),
    f"{GESNR}52": (("09", False),),
    f"{GESNR}53": (("09", False),),
    f"{GESNR}55": (("09", False),),
    f"{GESNR}61": (("09", False),),
    f"{GESNR}63": (("09", False),),
    f"{GESNR}65": (("09", False),),
    f"{GESNR}66": (("09", False), ("11", False)),
    f"{GESNR}68": (("09", False),),
    f"{GESNR}69": (("09", False),),
    f"{GESN}16": (("10", False),),
    f"{GESN}17": (("10", False),),
    f"{GESN}18": (("10", False),),
    f"{GESN}22": (("11", False), ("13", False)),
    f"{GESN}23": (("11", False),),
    f"{GESN}24": (("11", False), ("13", False)),
    f"{GESN}25": (("11", False),),
    f"{GESNM}12": (("11", False), ("13", False)),
    f"{GESN}20": (("12", False),),
    f"{GESNP}03": (("12", False),),
    f"{GESNM}39": (("13", False),),
    f"{GESN}34": (("14", False),),
    f"{GESNM}10": (("14", False),),
    f"{GESNM}11": (("15", False),),
    f"{GESN}33": (("16", False),),
    f"{GESNM}08": (("16", False),),
    f"{GESNP}01": (("16", False),),
}


def BuildSectionDict() -> dict[str, str]:
    """Return the primary fallback GESN-prefix-to-section table."""
    return {
        prefix: choices[0][0]
        for prefix, choices in FALLBACK_SECTION_CHOICES.items()
        if choices
    }


def ResolveSectionCode(
    code: object,
    is_demolition: bool,
    manual_section_mappings: dict[str, str] | None = None,
) -> str:
    """Resolve section code using the current EKR priority rules."""
    section_code, _source = ResolveSectionCodeWithSource(
        code,
        is_demolition,
        manual_section_mappings=manual_section_mappings,
    )
    return section_code


def ResolveSectionCodeWithSource(
    code: object,
    is_demolition: bool,
    manual_section_mappings: dict[str, str] | None = None,
) -> tuple[str, str]:
    """Resolve section code and return the rule source key."""
    canonical_code = CanonicalGesnCode(code)
    if canonical_code == "":
        return "", ""

    manual_section = _manual_section(canonical_code, manual_section_mappings)
    if manual_section != "":
        return manual_section, SECTION_SOURCE_MANUAL

    exact_section = _load_third_level_sections().get(canonical_code, "")
    if exact_section != "":
        return exact_section, SECTION_SOURCE_THIRD_LEVEL

    prefix = GESnPrefix(canonical_code)
    if prefix == "":
        return "", ""

    if _is_commissioning_prefix(prefix):
        return "17", SECTION_SOURCE_COMMISSIONING

    fallback_section = _resolve_fallback_prefix(prefix, is_demolition)
    if fallback_section != "":
        return fallback_section, SECTION_SOURCE_FALLBACK
    return "", ""


def ResolveExactThirdLevelSection(code: object) -> str:
    """Return the exact third-level EKR section, or empty when not mapped."""
    canonical_code = CanonicalGesnCode(code)
    if canonical_code == "":
        return ""
    return _load_third_level_sections().get(canonical_code, "")


def HasExactThirdLevelSection(code: object) -> bool:
    """Return True when section came from the exact third-level mapping."""
    return ResolveExactThirdLevelSection(code) != ""


def CanonicalGesnCode(code: object) -> str:
    """Return normalized code with FER/TER prefixes converted to GESN."""
    return _canonical_norm_code(NormCode(code))


def NormalizeSectionValue(value: object) -> str:
    """Return a two-digit EKR section value, or empty when invalid."""
    text = str(value).strip()
    if text == "":
        return ""
    if text.isdigit():
        number = int(text)
        if 1 <= number <= 99:
            return str(number).zfill(2)
    return ""


def GESnPrefix(code: object) -> str:
    """Extract a canonical GESN prefix from GESN, FER, or TER codes."""
    text = CanonicalGesnCode(code)
    position = text.find(GESN)
    if position == -1:
        return ""

    tail = text[position:]
    prefix = GESN
    idx = len(GESN)

    if idx < len(tail) and _is_cyrillic_upper_letter(tail[idx]):
        prefix += tail[idx]
        idx += 1

    if idx + 1 < len(tail) and tail[idx].isdigit() and tail[idx + 1].isdigit():
        prefix += tail[idx] + tail[idx + 1]

    return prefix


@lru_cache(maxsize=1)
def _load_third_level_sections() -> dict[str, str]:
    try:
        data = json.loads(THIRD_LEVEL_SECTION_MAP.read_text(encoding="utf-8"))
    except OSError:
        return {}

    if not isinstance(data, dict):
        return {}

    result: dict[str, str] = {}
    for raw_code, raw_section in data.items():
        code = CanonicalGesnCode(str(raw_code))
        section = NormalizeSectionValue(raw_section)
        if code != "" and section != "":
            result[code] = section
    return result


def _manual_section(
    canonical_code: str,
    manual_section_mappings: dict[str, str] | None,
) -> str:
    if not manual_section_mappings:
        return ""
    section = manual_section_mappings.get(canonical_code, "")
    return NormalizeSectionValue(section)


def _is_commissioning_prefix(prefix: str) -> bool:
    return prefix.startswith(GESNP)


def _resolve_fallback_prefix(prefix: str, is_demolition: bool) -> str:
    choices = FALLBACK_SECTION_CHOICES.get(prefix, ())
    if not choices:
        return ""

    if len(choices) > 1:
        demolition_choices = [section for section, is_demo in choices if is_demo]
        non_demolition_choices = [section for section, is_demo in choices if not is_demo]
        if is_demolition and demolition_choices:
            return demolition_choices[0]
        if not is_demolition and non_demolition_choices:
            return non_demolition_choices[0]

    return choices[0][0]


def _canonical_norm_code(value: object) -> str:
    text = NormCode(value)
    if text == "":
        return ""

    fer_pos = text.find(FER)
    ter_pos = text.find(TER)
    if fer_pos == -1 and ter_pos == -1:
        return text

    positions = [pos for pos in (fer_pos, ter_pos) if pos != -1]
    position = min(positions)
    source_prefix = FER if position == fer_pos else TER
    return text[:position] + GESN + text[position + len(source_prefix) :]


def _is_cyrillic_upper_letter(value: str) -> bool:
    codepoint = ord(value)
    return 0x0410 <= codepoint <= 0x042F or codepoint == 0x0401
