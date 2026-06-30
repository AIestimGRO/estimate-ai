"""Normalization helpers ported from the VBA matching macro."""

KR_SUFFIX = "/\u041a\u0420"
YO_LOWER = "\u0451"
YE_LOWER = "\u0435"
DEM_ROOT = "\u0434\u0435\u043c\u043e\u043d\u0442"
NBSP = "\u00a0"
SUPERSCRIPT_TWO = "\u00b2"
SUPERSCRIPT_THREE = "\u00b3"


def _collapse_spaces(value: str) -> str:
    while "  " in value:
        value = value.replace("  ", " ")
    return value


def NormCode(value: object) -> str:
    """Port of NormCode from Module3, DOMAIN_RULES.md section 2.1."""
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\r", " ")
    text = text.replace("\n", " ")
    text = text.replace("\t", " ")
    text = text.replace(NBSP, " ")
    text = text.strip().upper()

    if text == "":
        return ""

    text = _collapse_spaces(text)

    while " /" in text:
        text = text.replace(" /", "/")

    while "/ " in text:
        text = text.replace("/ ", "/")

    if text.endswith(KR_SUFFIX):
        text = text[: -len(KR_SUFFIX)]

    return text.strip()


def NormUnit(value: object) -> str:
    """Port of NormUnit from Module3, DOMAIN_RULES.md section 2.2."""
    if value is None or value == "":
        return ""

    text = str(value).strip().lower()
    text = text.replace(YO_LOWER, YE_LOWER)
    text = text.replace(NBSP, " ")
    text = text.replace(SUPERSCRIPT_TWO, "2")
    text = text.replace(SUPERSCRIPT_THREE, "3")

    text = _collapse_spaces(text)

    text = text.replace(" ", "")
    text = text.replace(".", "")
    text = text.replace(",", "")
    text = text.replace("^", "")

    return text


def HasDemontazh(value: object) -> bool:
    """Port of HasDemontazh from Module3, DOMAIN_RULES.md section 2.4."""
    if value is None:
        return False

    text = str(value).lower()
    text = text.replace(YO_LOWER, YE_LOWER)
    text = text.replace(NBSP, " ")

    for char in (
        ".",
        ",",
        ";",
        ":",
        "(",
        ")",
        "[",
        "]",
        "{",
        "}",
        "/",
        "\\",
        "-",
        "_",
        "+",
        "=",
        "\t",
        "\r",
        "\n",
    ):
        text = text.replace(char, " ")

    text = _collapse_spaces(text)

    for word in text.strip().split(" "):
        if word.startswith(DEM_ROOT):
            return True

    return False


def AnalogSearchKey(unit_value: object, code_value: object) -> str:
    """Port of AnalogSearchKey from Module3, DOMAIN_RULES.md section 2.3."""
    unit = NormUnit(unit_value)
    code = NormCode(code_value)

    if unit == "" or code == "":
        return ""

    return f"{unit}||{code}"
