"""Deterministic sheet-layout resolution for variable estimate/catalog files.

This module is a deliberate extension beyond the original VBA macros (which
read fixed, user-configured column numbers). It resolves where each logical
field lives on a worksheet by matching header text against a config-driven
synonym dictionary, so the pipeline can tolerate files that do not exactly
match the template.

Resolution priority per field (highest first):

1. explicit column pin (caller-provided override),
2. detection by header-text synonym (config `data/config/layout.json`),
3. template default column (caller-provided, optional fields only),
4. otherwise the field is reported missing.

Required fields never fall back to a template default: if they cannot be
pinned or detected, the layout is reported as not resolvable (the caller
then surfaces a "key data not found" result). All matching is exact/
substring against a fixed dictionary, so results are deterministic and
testable; there is no fuzzy or semantic matching here (see
DOMAIN_RULES.md section 8).

Business terms (Russian header wording) live in the JSON config, never as
string literals in this module (see AGENTS.md rule 3).
"""

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Protocol

NBSP = "\u00a0"

DEFAULT_MAX_BLANK_RUN = 5

FIELD_WORK_NAME = "work_name"
FIELD_UNIT = "unit"
FIELD_CODE = "code"
FIELD_BASE_PRICE = "base_price"

DEFAULT_REQUIRED_FIELDS: tuple[str, ...] = (FIELD_CODE, FIELD_UNIT, FIELD_BASE_PRICE)

MODE_EQUALS = "equals"
MODE_STARTSWITH = "startswith"
MODE_CONTAINS = "contains"

METHOD_EXPLICIT = "explicit"
METHOD_DETECTED = "detected"
METHOD_DEFAULT = "default"
METHOD_MISSING = "missing"

COEF_METHOD_EXPLICIT = "explicit"
COEF_METHOD_REGION = "labeled_region"
COEF_METHOD_LABEL = "labeled_coefficient"
COEF_METHOD_DEFAULT = "default"

DEFAULT_COEFFICIENT = 1.0

_CONFIG_RELATIVE_PATH = ("data", "config", "layout.json")


class CellSource(Protocol):
    """Minimal worksheet interface used by the resolver (openpyxl-compatible)."""

    def cell(self, row: int, column: int) -> Any: ...


class WorksheetSource(CellSource, Protocol):
    """Worksheet with a title, used for sheet ranking/selection (R1)."""

    title: str


class WorkbookSource(Protocol):
    """Workbook exposing its worksheets, used for sheet selection (R1)."""

    @property
    def worksheets(self) -> list: ...


@dataclass(frozen=True)
class FieldRule:
    """One field's header-matching rule."""

    mode: str
    patterns: tuple[str, ...]


@dataclass(frozen=True)
class LayoutConfig:
    """Loaded layout dictionary and header-scan bounds."""

    fields: dict[str, FieldRule]
    field_priority: tuple[str, ...]
    max_rows: int
    max_cols: int
    min_matched_fields: int
    max_blank_run: int = DEFAULT_MAX_BLANK_RUN
    region_label: FieldRule | None = None
    coefficient_label: FieldRule | None = None


@dataclass(frozen=True)
class ColumnResolution:
    """How one field's column was resolved."""

    field: str
    column: int | None
    method: str
    header_text: str | None = None


@dataclass(frozen=True)
class LayoutResult:
    """Resolved layout for one worksheet."""

    header_row: int
    columns: dict[str, ColumnResolution]
    missing_required: list[str]

    @property
    def ok(self) -> bool:
        return not self.missing_required

    def column(self, field: str) -> int | None:
        resolution = self.columns.get(field)
        return None if resolution is None else resolution.column


@dataclass(frozen=True)
class CoefficientResolution:
    """Resolved regional coefficient and how it was found (R16)."""

    value: float
    method: str
    region: str | None = None
    label_cell: tuple[int, int] | None = None


@dataclass(frozen=True)
class AveragePlacement:
    """Target column for the average-price output relative to base price (R12)."""

    column: int
    needs_insert: bool


@dataclass(frozen=True)
class SheetCandidate:
    """One worksheet scored for its likelihood of holding the data (R1)."""

    title: str
    index: int
    layout: LayoutResult
    score: int

    @property
    def ok(self) -> bool:
        return self.layout.ok


@dataclass(frozen=True)
class SheetSelection:
    """Which sheet(s) to run, or a request for the user to choose (R1).

    Supports the intended web flow: a single sheet is used automatically; when
    several sheets qualify, `needs_user_choice` is set and `candidates` lists
    the options to present as buttons. The user (or a caller) can also force a
    selection by title via `selected_titles`.
    """

    chosen: list[SheetCandidate]
    needs_user_choice: bool
    candidates: list[SheetCandidate]


def load_layout_config(config_path: str | Path | None = None) -> LayoutConfig:
    """Load the layout dictionary from JSON (defaults to data/config/layout.json)."""
    path = _default_config_path() if config_path is None else Path(config_path)
    with path.open("r", encoding="utf-8") as config_file:
        raw = json.load(config_file)

    header_scan = raw.get("header_scan", {})
    data_scan = raw.get("data_scan", {})
    fields = {
        field_name: FieldRule(
            mode=rule["mode"],
            patterns=tuple(rule.get("patterns", [])),
        )
        for field_name, rule in raw.get("fields", {}).items()
    }
    field_priority = tuple(raw.get("field_priority", tuple(fields.keys())))

    coefficient = raw.get("coefficient", {})
    region_label = _load_rule(coefficient.get("region_label"))
    coefficient_label = _load_rule(coefficient.get("coefficient_label"))

    return LayoutConfig(
        fields=fields,
        field_priority=field_priority,
        max_rows=int(header_scan.get("max_rows", 50)),
        max_cols=int(header_scan.get("max_cols", 60)),
        min_matched_fields=int(header_scan.get("min_matched_fields", 2)),
        max_blank_run=int(data_scan.get("max_blank_run", DEFAULT_MAX_BLANK_RUN)),
        region_label=region_label,
        coefficient_label=coefficient_label,
    )


def _load_rule(raw_rule: dict | None) -> FieldRule | None:
    if not raw_rule:
        return None
    return FieldRule(
        mode=raw_rule["mode"],
        patterns=tuple(raw_rule.get("patterns", [])),
    )


def resolve_layout(
    worksheet: CellSource,
    config: LayoutConfig,
    *,
    default_columns: dict[str, int] | None = None,
    explicit_columns: dict[str, int] | None = None,
    required_fields: tuple[str, ...] = DEFAULT_REQUIRED_FIELDS,
) -> LayoutResult:
    """Resolve the header row and per-field columns for one worksheet."""
    defaults = {} if default_columns is None else dict(default_columns)
    explicit = {} if explicit_columns is None else dict(explicit_columns)

    header_row, detected = _detect_header_row(worksheet, config, explicit)

    all_fields = _all_field_names(config, defaults, explicit, required_fields)
    header_texts = (
        _row_header_texts(worksheet, header_row, config.max_cols)
        if header_row > 0
        else {}
    )

    columns: dict[str, ColumnResolution] = {}
    missing_required: list[str] = []

    for field in all_fields:
        is_required = field in required_fields
        resolution = _resolve_one_field(
            field,
            detected,
            explicit,
            defaults,
            header_texts,
            is_required,
        )
        columns[field] = resolution
        if is_required and resolution.method == METHOD_MISSING:
            missing_required.append(field)

    return LayoutResult(
        header_row=header_row,
        columns=columns,
        missing_required=missing_required,
    )


def format_layout_report(result: LayoutResult) -> str:
    """Human-readable summary of a resolved layout (R19 diagnostics)."""
    lines = [f"header_row: {result.header_row or 'not found'}"]
    for field in sorted(result.columns):
        resolution = result.columns[field]
        column = "-" if resolution.column is None else str(resolution.column)
        header = f" <- '{resolution.header_text}'" if resolution.header_text else ""
        lines.append(f"  {field:<12} col={column:<4} [{resolution.method}]{header}")
    if result.missing_required:
        lines.append(f"missing required: {', '.join(result.missing_required)}")
    return "\n".join(lines)


def resolve_regional_coefficient(
    worksheet: CellSource,
    config: LayoutConfig,
    *,
    explicit_value: object = None,
    max_rows: int | None = None,
    max_cols: int | None = None,
) -> CoefficientResolution:
    """Locate the regional coefficient by label (R16).

    Primary pattern (user rule): a "Region" label cell with a "Coefficient"
    label directly below it; the region name sits to the right of the region
    label and the numeric coefficient to the right of the coefficient label.
    Fallbacks: a standalone coefficient label with a number to its right; a
    caller-provided explicit value; otherwise the default coefficient (1.0).
    All label wording lives in data/config/layout.json.
    """
    if explicit_value is not None:
        value = _parse_number(explicit_value)
        if value is not None and value > 0:
            return CoefficientResolution(value=value, method=COEF_METHOD_EXPLICIT)

    rows = config.max_rows if max_rows is None else max_rows
    cols = config.max_cols if max_cols is None else max_cols

    stacked = _find_stacked_coefficient(worksheet, config, rows, cols)
    if stacked is not None:
        return stacked

    standalone = _find_standalone_coefficient(worksheet, config, rows, cols)
    if standalone is not None:
        return standalone

    return CoefficientResolution(value=DEFAULT_COEFFICIENT, method=COEF_METHOD_DEFAULT)


def resolve_average_placement(
    base_price_column: int,
    occupied_columns: set[int],
) -> AveragePlacement:
    """Pick the average-price output column next to the base price (R12).

    Writes into the column immediately to the right of the base price. If
    that neighbour already holds data, a new column must be inserted there
    (the writer performs the actual insert/shift); this function only decides
    the target column and whether an insert is required.
    """
    target = base_price_column + 1
    return AveragePlacement(column=target, needs_insert=target in occupied_columns)


def occupied_columns(
    worksheet: CellSource,
    row: int,
    max_cols: int,
) -> set[int]:
    """Columns holding a non-empty value in the given row."""
    return {
        column
        for column in range(1, max_cols + 1)
        if _normalize_header_text(worksheet.cell(row=row, column=column).value)
    }


def rank_sheets(
    workbook: WorkbookSource,
    config: LayoutConfig,
    *,
    default_columns: dict[str, int] | None = None,
    explicit_columns: dict[str, int] | None = None,
    required_fields: tuple[str, ...] = DEFAULT_REQUIRED_FIELDS,
) -> list[SheetCandidate]:
    """Score every worksheet by how well its layout resolves (R1).

    Sorted best-first: resolvable sheets first, then by number of fields
    resolved via detection/explicit pin, then by original sheet order.
    """
    candidates: list[SheetCandidate] = []
    for index, worksheet in enumerate(workbook.worksheets):
        layout = resolve_layout(
            worksheet,
            config,
            default_columns=default_columns,
            explicit_columns=explicit_columns,
            required_fields=required_fields,
        )
        score = sum(
            1
            for resolution in layout.columns.values()
            if resolution.method in (METHOD_DETECTED, METHOD_EXPLICIT)
        )
        candidates.append(
            SheetCandidate(
                title=getattr(worksheet, "title", str(index)),
                index=index,
                layout=layout,
                score=score,
            )
        )

    candidates.sort(key=lambda candidate: (not candidate.ok, -candidate.score, candidate.index))
    return candidates


def select_sheets(
    workbook: WorkbookSource,
    config: LayoutConfig,
    *,
    selected_titles: set[str] | None = None,
    default_columns: dict[str, int] | None = None,
    explicit_columns: dict[str, int] | None = None,
    required_fields: tuple[str, ...] = DEFAULT_REQUIRED_FIELDS,
) -> SheetSelection:
    """Decide which sheet(s) to run, or ask the user to choose (R1).

    - explicit `selected_titles` -> use exactly those sheets;
    - a single worksheet in the file -> use it automatically;
    - several worksheets but exactly one resolvable -> use it automatically;
    - several resolvable worksheets -> `needs_user_choice` with candidates;
    - none resolvable -> empty selection (caller surfaces "data not found").
    """
    ranked = rank_sheets(
        workbook,
        config,
        default_columns=default_columns,
        explicit_columns=explicit_columns,
        required_fields=required_fields,
    )

    if selected_titles is not None:
        chosen = [c for c in ranked if c.title in selected_titles]
        return SheetSelection(chosen=chosen, needs_user_choice=False, candidates=ranked)

    if len(ranked) <= 1:
        return SheetSelection(chosen=list(ranked), needs_user_choice=False, candidates=ranked)

    resolvable = [c for c in ranked if c.ok]
    if len(resolvable) == 1:
        return SheetSelection(chosen=resolvable, needs_user_choice=False, candidates=ranked)
    if len(resolvable) > 1:
        return SheetSelection(chosen=[], needs_user_choice=True, candidates=resolvable)

    return SheetSelection(chosen=[], needs_user_choice=False, candidates=ranked)


def data_row_numbers(
    worksheet: CellSource,
    start_row: int,
    key_columns: list[int],
    *,
    max_blank_run: int = DEFAULT_MAX_BLANK_RUN,
    max_row: int | None = None,
) -> list[int]:
    """Row numbers holding data, tolerating isolated blank rows (R5).

    Blank rows (all `key_columns` empty) are skipped, not returned, but the
    scan continues past them; it stops only after `max_blank_run` consecutive
    blank rows (end-of-table). Fully-empty columns in the body do not matter
    because only the resolved key columns are inspected. This prevents the
    read from breaking on one or two stray empty rows inside the table.
    """
    limit = _worksheet_max_row(worksheet) if max_row is None else max_row

    rows: list[int] = []
    blank_run = 0
    row = start_row
    while row <= limit:
        if _is_blank_row(worksheet, row, key_columns):
            blank_run += 1
            if blank_run >= max_blank_run:
                break
        else:
            blank_run = 0
            rows.append(row)
        row += 1

    return rows


def _is_blank_row(worksheet: CellSource, row: int, key_columns: list[int]) -> bool:
    for column in key_columns:
        if _normalize_header_text(worksheet.cell(row=row, column=column).value):
            return False
    return True


def _worksheet_max_row(worksheet: CellSource) -> int:
    return int(getattr(worksheet, "max_row", 0) or 0)


def _find_stacked_coefficient(
    worksheet: CellSource,
    config: LayoutConfig,
    rows: int,
    cols: int,
) -> CoefficientResolution | None:
    if config.region_label is None or config.coefficient_label is None:
        return None

    for row in range(1, rows):
        for column in range(1, cols + 1):
            region_text = _normalize_header_text(
                worksheet.cell(row=row, column=column).value
            )
            if not region_text or not _matches(region_text, config.region_label):
                continue

            below_text = _normalize_header_text(
                worksheet.cell(row=row + 1, column=column).value
            )
            if not below_text or not _matches(below_text, config.coefficient_label):
                continue

            value = _parse_number(
                worksheet.cell(row=row + 1, column=column + 1).value
            )
            if value is None or value <= 0:
                continue

            region_name = _trimmed_text(
                worksheet.cell(row=row, column=column + 1).value
            )
            return CoefficientResolution(
                value=value,
                method=COEF_METHOD_REGION,
                region=region_name,
                label_cell=(row + 1, column),
            )

    return None


def _find_standalone_coefficient(
    worksheet: CellSource,
    config: LayoutConfig,
    rows: int,
    cols: int,
) -> CoefficientResolution | None:
    if config.coefficient_label is None:
        return None

    for row in range(1, rows + 1):
        for column in range(1, cols + 1):
            label_text = _normalize_header_text(
                worksheet.cell(row=row, column=column).value
            )
            if not label_text or not _matches(label_text, config.coefficient_label):
                continue

            value = _parse_number(worksheet.cell(row=row, column=column + 1).value)
            if value is None or value <= 0:
                continue

            return CoefficientResolution(
                value=value,
                method=COEF_METHOD_LABEL,
                label_cell=(row, column),
            )

    return None


def _trimmed_text(value: object) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _parse_number(value: object) -> float | None:
    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if text == "":
        return None

    text = text.replace(NBSP, "").replace(" ", "")
    if "," in text and "." not in text:
        text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return None


def _resolve_one_field(
    field: str,
    detected: dict[str, int],
    explicit: dict[str, int],
    defaults: dict[str, int],
    header_texts: dict[int, str],
    is_required: bool,
) -> ColumnResolution:
    if field in explicit:
        column = explicit[field]
        return ColumnResolution(
            field=field,
            column=column,
            method=METHOD_EXPLICIT,
            header_text=header_texts.get(column),
        )

    if field in detected:
        column = detected[field]
        return ColumnResolution(
            field=field,
            column=column,
            method=METHOD_DETECTED,
            header_text=header_texts.get(column),
        )

    if not is_required and field in defaults:
        return ColumnResolution(
            field=field,
            column=defaults[field],
            method=METHOD_DEFAULT,
        )

    return ColumnResolution(field=field, column=None, method=METHOD_MISSING)


def _detect_header_row(
    worksheet: CellSource,
    config: LayoutConfig,
    explicit: dict[str, int],
) -> tuple[int, dict[str, int]]:
    """Pick the row with the most detected fields (min threshold applies).

    Ties keep the first (smallest-index) row. Explicit pins do not help
    locate the header row, so scoring counts detected fields only.
    """
    best_row = 0
    best_columns: dict[str, int] = {}
    best_count = -1

    for row in range(1, config.max_rows + 1):
        detected = _detect_columns_in_row(worksheet, row, config, explicit)
        count = len(detected)
        if count > best_count:
            best_count = count
            best_row = row
            best_columns = detected

    if best_count < config.min_matched_fields:
        return 0, {}

    return best_row, best_columns


def _detect_columns_in_row(
    worksheet: CellSource,
    row: int,
    config: LayoutConfig,
    explicit: dict[str, int],
) -> dict[str, int]:
    header_texts = _row_header_texts(worksheet, row, config.max_cols)

    claimed: set[int] = {explicit[field] for field in explicit if field in config.fields}
    detected: dict[str, int] = {}

    for field in config.field_priority:
        if field in explicit:
            continue
        rule = config.fields.get(field)
        if rule is None:
            continue

        for column in range(1, config.max_cols + 1):
            if column in claimed:
                continue
            text = header_texts.get(column)
            if text and _matches(text, rule):
                detected[field] = column
                claimed.add(column)
                break

    return detected


def _row_header_texts(
    worksheet: CellSource,
    row: int,
    max_cols: int,
) -> dict[int, str]:
    texts: dict[int, str] = {}
    for column in range(1, max_cols + 1):
        value = worksheet.cell(row=row, column=column).value
        normalized = _normalize_header_text(value)
        if normalized:
            texts[column] = normalized
    return texts


def _matches(header_text: str, rule: FieldRule) -> bool:
    for pattern in rule.patterns:
        if rule.mode == MODE_EQUALS and header_text == pattern:
            return True
        if rule.mode == MODE_STARTSWITH and header_text.startswith(pattern):
            return True
        if rule.mode == MODE_CONTAINS and pattern in header_text:
            return True
    return False


def _normalize_header_text(value: object) -> str:
    """Loose header normalization: lowercase, drop punctuation, collapse spaces.

    Intentionally looser than NormUnit/NormCode (see DOMAIN_RULES.md section
    9.2 note on separate header normalization) and not shared with them.
    """
    if value is None:
        return ""

    text = str(value).lower()
    text = text.replace(NBSP, " ")
    for char in ("\r", "\n", "\t", ".", ",", ";", ":"):
        text = text.replace(char, " ")

    while "  " in text:
        text = text.replace("  ", " ")

    return text.strip()


def _all_field_names(
    config: LayoutConfig,
    defaults: dict[str, int],
    explicit: dict[str, int],
    required_fields: tuple[str, ...],
) -> list[str]:
    seen: dict[str, None] = {}
    for field in (*config.field_priority, *defaults, *explicit, *required_fields):
        seen.setdefault(field, None)
    return list(seen)


def _default_config_path() -> Path:
    root = Path(__file__).resolve().parent.parent
    return root.joinpath(*_CONFIG_RELATIVE_PATH)
