"""Section-code resolution ported from the VBA matching macro."""

from core.normalize import NormCode


GESN = "\u0413\u042d\u0421\u041d"
GESNM = f"{GESN}\u041c"
GESNP = f"{GESN}\u041f"
GESNR = f"{GESN}\u0420"

DEMOLITION_PRIORITY_PREFIXES = {
    f"{GESN}09",
    f"{GESN}27",
    f"{GESN}28",
    f"{GESN}46",
    f"{GESNR}67",
}


def BuildSectionDict() -> dict[str, str]:
    """Return the GESN-prefix-to-section table from Module3."""
    return {
        f"{GESN}01": "01",
        f"{GESN}04": "01",
        f"{GESN}05": "02",
        f"{GESN}06": "02",
        f"{GESN}07": "02",
        f"{GESN}08": "03",
        f"{GESN}10": "03",
        f"{GESN}11": "03",
        f"{GESN}12": "03",
        f"{GESN}15": "03",
        f"{GESN}09": "04",
        f"{GESN}39": "04",
        f"{GESNM}38": "04",
        f"{GESN}27": "05",
        f"{GESN}28": "05",
        f"{GESN}47": "05",
        f"{GESNM}20": "05",
        f"{GESNM}03": "06",
        f"{GESNM}06": "06",
        f"{GESNM}07": "06",
        f"{GESNM}13": "06",
        f"{GESNM}18": "06",
        f"{GESNM}19": "06",
        f"{GESNM}22": "06",
        f"{GESNM}37": "06",
        f"{GESNP}07": "06",
        f"{GESN}13": "07",
        f"{GESN}26": "07",
        f"{GESN}45": "07",
        f"{GESN}46": "08",
        f"{GESNR}67": "08",
        f"{GESNR}51": "09",
        f"{GESNR}52": "09",
        f"{GESNR}53": "09",
        f"{GESNR}55": "09",
        f"{GESNR}61": "09",
        f"{GESNR}63": "09",
        f"{GESNR}65": "09",
        f"{GESNR}66": "09",
        f"{GESNR}68": "09",
        f"{GESNR}69": "09",
        f"{GESN}16": "10",
        f"{GESN}17": "10",
        f"{GESN}18": "10",
        f"{GESN}22": "11",
        f"{GESN}23": "11",
        f"{GESN}24": "11",
        f"{GESN}25": "11",
        f"{GESNM}12": "11",
        f"{GESN}20": "12",
        f"{GESNP}03": "12",
        f"{GESNM}39": "13",
        f"{GESN}34": "14",
        f"{GESNM}10": "14",
        f"{GESNM}11": "15",
        f"{GESN}33": "16",
        f"{GESNM}08": "16",
        f"{GESNP}01": "16",
    }


def ResolveSectionCode(code: object, is_demolition: bool) -> str:
    """Resolve section code from GESN prefix, matching Module3 behavior."""
    section_dict = BuildSectionDict()
    prefix = GESnPrefix(NormCode(code))
    if prefix == "":
        return ""

    if prefix in DEMOLITION_PRIORITY_PREFIXES:
        if is_demolition:
            return "08"

        section_code = section_dict.get(prefix, "").strip()
        if section_code != "" and section_code != "08":
            return section_code
        return "08"

    return section_dict.get(prefix, "")


def GESnPrefix(code: object) -> str:
    """Extract a GESN prefix using the same rules as Module3."""
    text = str(code).upper()
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


def _is_cyrillic_upper_letter(value: str) -> bool:
    codepoint = ord(value)
    return 0x0410 <= codepoint <= 0x042F or codepoint == 0x0401
