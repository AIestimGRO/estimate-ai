"""TKP matching primitives for live pairing and admin diagnostics.

The legacy/global helper remains deterministic lexical matching. The paired
RNMC->TKP product path first applies deterministic task, unit, quantity,
multiplicity, and demolition filters and may then accept a caller-supplied
local semantic scorer (Qwen3 in the web application) for work-name ranking.
Semantic scoring never derives a price; winner/reserve price selection and unit
conversion remain deterministic.
"""

from collections.abc import Callable
from dataclasses import dataclass
from difflib import SequenceMatcher
from math import isclose

from core.multiplicity import multiplicity_is_compatible
from core.normalize import HasDemontazh, NormUnit
from core.task_numbers import extract_task_numbers, normalize_single_task_number
from core.unit_scaling import (
    compatible_unit_conversion,
    normalized_quantity,
    normalized_unit_price,
    positive_float,
)

# Common Russian function words that carry no discriminative weight for
# construction work-item names; stripping them focuses the token-overlap
# signal on content words instead of prepositions/conjunctions that both
# sides of almost any pair will share anyway.
STOPWORDS = frozenset(
    {
        "\u0438", "\u0432", "\u043d\u0430", "\u0441", "\u0438\u0437",
        "\u043f\u043e", "\u0434\u043b\u044f", "\u043e\u0442", "\u043a",
        "\u0443", "\u043e", "\u043e\u0431", "\u0437\u0430", "\u043d\u0435",
        "\u0434\u043e", "\u043f\u0440\u0438", "\u0438\u043b\u0438", "\u043a\u0430\u043a",
        "\u0447\u0442\u043e", "\u0442\u043e", "\u0442.\u043f", "\u0438 \u0442.\u043f",
        "\u043d\u0435\u0442", "\u0435\u0433\u043e", "\u0435\u0435", "\u0438\u0445",
    }
)

# Coarse equivalence groups for the leading action word (the first
# meaningful word of a GESN-style name, almost always the operation type:
# "Демонтаж ...", "Устройство ...", "Засыпка ..."). This is deliberately a
# small, hand-curated list of the operations that showed up as legitimate
# estimate-vs-TKP wording differences in real data (демонтаж/разборка), not
# a general synonym dictionary - unrecognized words are simply compared for
# exact equality, which is the safe default (no penalty is applied to
# established synonyms; an unrecognized different word is still penalized).
ACTION_SYNONYM_GROUPS: tuple[frozenset[str], ...] = (
    frozenset({"\u0434\u0435\u043c\u043e\u043d\u0442\u0430\u0436", "\u0440\u0430\u0437\u0431\u043e\u0440\u043a\u0430", "\u0441\u043d\u044f\u0442\u0438\u0435", "\u0432\u0441\u043a\u0440\u044b\u0442\u0438\u0435"}),
    frozenset({"\u0443\u0441\u0442\u0440\u043e\u0439\u0441\u0442\u0432\u043e", "\u043c\u043e\u043d\u0442\u0430\u0436", "\u0443\u0441\u0442\u0430\u043d\u043e\u0432\u043a\u0430", "\u043f\u0440\u043e\u043a\u043b\u0430\u0434\u043a\u0430", "\u0441\u0431\u043e\u0440\u043a\u0430"}),
    frozenset({"\u043f\u0435\u0440\u0435\u0432\u043e\u0437\u043a\u0430", "\u0442\u0440\u0430\u043d\u0441\u043f\u043e\u0440\u0442\u0438\u0440\u043e\u0432\u043a\u0430", "\u0434\u043e\u0441\u0442\u0430\u0432\u043a\u0430"}),
    frozenset({"\u043f\u043e\u043a\u0440\u044b\u0442\u0438\u0435", "\u043e\u0431\u043b\u0438\u0446\u043e\u0432\u043a\u0430", "\u043e\u0431\u0448\u0438\u0432\u043a\u0430"}),
    frozenset({"\u043e\u043a\u0440\u0430\u0441\u043a\u0430", "\u043f\u043e\u043a\u0440\u0430\u0441\u043a\u0430", "\u043e\u043a\u0440\u0430\u0448\u0438\u0432\u0430\u043d\u0438\u0435"}),
    frozenset({"\u043e\u0447\u0438\u0441\u0442\u043a\u0430", "\u0440\u0430\u0441\u0447\u0438\u0441\u0442\u043a\u0430", "\u0437\u0430\u0447\u0438\u0441\u0442\u043a\u0430"}),
    frozenset({"\u0437\u0430\u0441\u044b\u043f\u043a\u0430", "\u043f\u043e\u0434\u0441\u044b\u043f\u043a\u0430"}),
    frozenset({"\u0443\u043f\u043b\u043e\u0442\u043d\u0435\u043d\u0438\u0435", "\u0442\u0440\u0430\u043c\u0431\u043e\u0432\u0430\u043d\u0438\u0435", "\u0442\u0440\u0430\u043c\u0431\u043e\u0432\u043a\u0430"}),
)

_PUNCTUATION_TABLE = str.maketrans(
    {ch: " " for ch in ".,;:!?()[]{}\"'\u00ab\u00bb-/\\«»–—"}
)

# Blend weights: token overlap dominates (word-order-independent is more
# important for these names than exact character sequence), sequence ratio
# is a secondary tie-breaker/near-duplicate signal.
TOKEN_WEIGHT = 0.65
SEQUENCE_WEIGHT = 0.35
UNIT_MATCH_BONUS = 5.0
LEADING_WORD_MISMATCH_PENALTY = 20.0

DEFAULT_MIN_SCORE = 55.0
DEFAULT_LIMIT = 3
QUANTITY_REL_TOL = 1e-9
QUANTITY_ABS_TOL = 1e-8
PRICE_TIE_ABS_TOL = 1e-9
PRICE_SOURCE_WINNER = "winner"
PRICE_SOURCE_RESERVE = "reserve"
DEFAULT_SEMANTIC_MIN_SCORE = 45.0
SEMANTIC_TIE_ABS_TOL = 1e-6

SemanticScoreFunction = Callable[[str, list[str]], list[float]]


class TkpSemanticScoringError(RuntimeError):
    """Raised when the configured semantic model cannot score candidates."""



def leading_action_word(normalized_text: str) -> str:
    """First non-stopword token - the operation type in a GESN-style name."""
    for token in normalized_text.split(" "):
        if token and token not in STOPWORDS:
            return token
    return ""


def same_action_group(word_a: str, word_b: str) -> bool:
    """True if the two leading words name the same (or an equivalent) operation.

    Missing words on either side are treated as "nothing to penalize" (the
    caller has already filtered out empty names elsewhere).
    """
    if not word_a or not word_b:
        return True
    if word_a == word_b:
        return True
    return any(word_a in group and word_b in group for group in ACTION_SYNONYM_GROUPS)


def normalize_for_matching(value: object) -> tuple[str, frozenset[str]]:
    """Lowercase, strip punctuation/digits-only noise, and tokenize.

    Returns (normalized_text, content_tokens) - the text is used for the
    sequence-ratio signal, the token set for the Jaccard signal.
    """
    if value is None:
        return "", frozenset()
    text = str(value).lower().translate(_PUNCTUATION_TABLE)
    text = text.replace("\u00a0", " ")
    while "  " in text:
        text = text.replace("  ", " ")
    text = text.strip()

    tokens = frozenset(
        token for token in text.split(" ") if token and token not in STOPWORDS
    )
    return text, tokens


@dataclass(frozen=True)
class TkpCatalogEntry:
    """One catalog item pre-normalized for repeated matching queries."""

    item_id: int
    item_name: str
    unit: str
    winner_unit_price_no_vat: float | None
    winner_name: str
    source_file_name: str
    task_no: str
    _normalized_text: str
    _normalized_tokens: frozenset[str]
    _normalized_unit: str
    _leading_word: str
    section_name: str = ""
    subsection_name: str = ""
    qty: float | None = None
    rnmc_unit_price_no_vat: float | None = None
    reserve_unit_price_no_vat: float | None = None
    reserve_name: str = ""
    normalized_task_no: str = ""


@dataclass(frozen=True)
class TkpMatch:
    """One scored TKP candidate with participant prices in the RNMC unit."""

    entry: TkpCatalogEntry
    score: float
    winner_price: float | None = None
    reserve_price: float | None = None
    quantity_matched: bool = False
    rnmc_price_delta: float | None = None
    semantic_score: float | None = None
    match_method: str = "lexical"

    @property
    def has_winner_price(self) -> bool:
        return self.winner_price is not None

    @property
    def has_reserve_price(self) -> bool:
        return self.reserve_price is not None


@dataclass(frozen=True)
class TkpAnalogMatch:
    """TKP result paired with one concrete RNMC analog output column."""

    task_id: str
    price_position: int
    match: TkpMatch


def build_tkp_catalog_index(items: list) -> list[TkpCatalogEntry]:
    """Pre-normalize a list of TkpItemRecord-like objects for matching.

    Accepts anything with `.item_name`, `.unit`, `.winner_unit_price_no_vat`,
    `.winner_name`, `.source_file_name` (or `.file_name`), `.task_no`, and
    `.id` attributes - i.e. core.storage.tkp.TkpItemRecord.
    """
    index: list[TkpCatalogEntry] = []
    for item in items:
        normalized_text, tokens = normalize_for_matching(item.item_name)
        if not tokens:
            continue
        source_file_name = getattr(item, "source_file_name", None)
        if source_file_name is None:
            source_file_name = getattr(item, "file_name", "")
        index.append(
            TkpCatalogEntry(
                item_id=item.id,
                item_name=item.item_name,
                unit=item.unit,
                winner_unit_price_no_vat=item.winner_unit_price_no_vat,
                winner_name=item.winner_name,
                source_file_name=source_file_name,
                task_no=item.task_no,
                _normalized_text=normalized_text,
                _normalized_tokens=tokens,
                _normalized_unit=NormUnit(item.unit),
                _leading_word=leading_action_word(normalized_text),
                section_name=str(getattr(item, "section_name", "") or ""),
                subsection_name=str(getattr(item, "subsection_name", "") or ""),
                qty=positive_float(getattr(item, "qty", None)),
                rnmc_unit_price_no_vat=positive_float(
                    getattr(item, "rnmc_unit_price_no_vat", None)
                ),
                reserve_unit_price_no_vat=positive_float(
                    getattr(item, "reserve_unit_price_no_vat", None)
                ),
                reserve_name=str(getattr(item, "reserve_name", "") or ""),
                normalized_task_no=normalize_single_task_number(item.task_no),
            )
        )
    return index


def score_names(text_a: str, tokens_a: frozenset[str], text_b: str, tokens_b: frozenset[str]) -> float:
    """0-100 lexical similarity between two already-normalized names."""
    if not tokens_a or not tokens_b:
        return 0.0

    jaccard = _jaccard(tokens_a, tokens_b)
    sequence_ratio = SequenceMatcher(None, text_a, text_b).ratio()

    return 100.0 * (TOKEN_WEIGHT * jaccard + SEQUENCE_WEIGHT * sequence_ratio)


def _jaccard(tokens_a: frozenset[str], tokens_b: frozenset[str]) -> float:
    if not tokens_a or not tokens_b:
        return 0.0
    intersection = len(tokens_a & tokens_b)
    union = len(tokens_a) + len(tokens_b) - intersection
    return intersection / union if union else 0.0


def _max_possible_score(jaccard: float, unit_bonus: float) -> float:
    """Best score this pair could reach if sequence_ratio were a perfect 1.0.

    Used to skip the O(len(a)*len(b)) SequenceMatcher call entirely for
    candidates that cannot clear min_score no matter what the character-level
    comparison turns out to be - the token overlap (cheap: a frozenset
    intersection) already caps the outcome. This is what keeps a ~3000-row
    catalog scan per estimate line fast without changing which matches are
    returned for any given min_score.
    """
    return 100.0 * (TOKEN_WEIGHT * jaccard + SEQUENCE_WEIGHT * 1.0) + unit_bonus


def build_tkp_task_index(
    index: list[TkpCatalogEntry],
) -> dict[str, tuple[TkpCatalogEntry, ...]]:
    """Group live TKP entries by normalized task number for cheap pair lookup."""
    grouped: dict[str, list[TkpCatalogEntry]] = {}
    for entry in index:
        if not entry.normalized_task_no:
            continue
        grouped.setdefault(entry.normalized_task_no, []).append(entry)
    return {key: tuple(values) for key, values in grouped.items()}


def find_tkp_match_for_rnmc_analog(
    analog_entry,
    task_index: dict[str, tuple[TkpCatalogEntry, ...]],
    *,
    min_score: float = DEFAULT_MIN_SCORE,
    semantic_scorer: SemanticScoreFunction | None = None,
    semantic_model_name: str = "",
    semantic_min_score: float = DEFAULT_SEMANTIC_MIN_SCORE,
) -> TkpMatch | None:
    """Find one TKP work item paired to one RNMC catalog entry.

    Candidate narrowing follows the product rule: task number, compatible
    scaled unit, exact normalized quantity when available, then name score.
    When a semantic scorer is supplied, it is the primary name-ranking signal;
    lexical scoring remains only the non-semantic fallback for direct callers.
    RNMC unit price resolves truly tied name candidates after unit scaling.
    Winner and reserve unit prices are retained separately and converted into
    the concrete RNMC analog unit; they are never averaged at this stage.
    """
    task_numbers = set(extract_task_numbers(getattr(analog_entry, "task_id", "")))
    if not task_numbers:
        return None

    original_row = getattr(analog_entry, "original_row", None)
    if original_row is None:
        return None
    query_name = getattr(original_row, "work_name", "")
    query_unit = getattr(original_row, "unit", "")
    query_quantity = normalized_quantity(
        getattr(original_row, "quantity", None),
        query_unit,
    )
    query_rnmc_price = normalized_unit_price(
        getattr(original_row, "price_original", None),
        query_unit,
    )

    candidate_pool: list[TkpCatalogEntry] = []
    for task_no in task_numbers:
        candidate_pool.extend(task_index.get(task_no, ()))

    query_demolition = HasDemontazh(query_name)
    candidates: list[tuple[TkpCatalogEntry, float]] = []
    for entry in candidate_pool:
        conversion = compatible_unit_conversion(query_unit, entry.unit)
        if conversion is None:
            continue
        if not _entry_has_usable_participant_price(entry):
            continue
        if not multiplicity_is_compatible(query_name, entry.item_name):
            continue
        candidate_context = " ".join(
            part for part in (entry.section_name, entry.subsection_name, entry.item_name) if part
        )
        if HasDemontazh(candidate_context) != query_demolition:
            continue
        candidate_quantity = normalized_quantity(entry.qty, entry.unit)
        candidates.append((entry, candidate_quantity if candidate_quantity is not None else float("nan")))

    if not candidates:
        return None

    quantity_matched = False
    if query_quantity is not None:
        exact_quantity = [
            pair
            for pair in candidates
            if _same_quantity(query_quantity, pair[1])
        ]
        if exact_quantity:
            candidates = exact_quantity
            quantity_matched = True

    query_text, query_tokens = normalize_for_matching(query_name)
    if not query_tokens:
        return None

    semantic_score: float | None = None
    match_method = "lexical"
    if semantic_scorer is not None:
        scored = _semantic_name_scores(
            query_name,
            [entry for entry, _ in candidates],
            semantic_scorer,
            min_score=semantic_min_score,
        )
        match_method = semantic_model_name or "semantic"
        tie_tolerance = SEMANTIC_TIE_ABS_TOL
    else:
        query_leading_word = leading_action_word(query_text)
        scored = []
        for entry, _ in candidates:
            score = _name_score_for_entry(
                query_text,
                query_tokens,
                query_leading_word,
                entry,
                min_score=min_score,
            )
            if score >= min_score:
                scored.append((score, entry))
        tie_tolerance = PRICE_TIE_ABS_TOL

    if not scored:
        return None

    best_score = max(score for score, _ in scored)
    top = [
        entry
        for score, entry in scored
        if isclose(score, best_score, rel_tol=0.0, abs_tol=tie_tolerance)
    ]
    selected, rnmc_delta = _resolve_equal_name_candidates(
        top,
        query_rnmc_price,
    )
    if selected is None:
        return None

    winner_price, reserve_price = _paired_output_prices(selected, query_unit)
    if winner_price is None and reserve_price is None:
        return None
    if semantic_scorer is not None:
        semantic_score = max(0.0, min(best_score, 100.0))
    return TkpMatch(
        entry=selected,
        score=max(0.0, min(best_score, 100.0)),
        winner_price=winner_price,
        reserve_price=reserve_price,
        quantity_matched=quantity_matched,
        rnmc_price_delta=rnmc_delta,
        semantic_score=semantic_score,
        match_method=match_method,
    )



def _semantic_name_scores(
    query_name: object,
    entries: list[TkpCatalogEntry],
    semantic_scorer: SemanticScoreFunction,
    *,
    min_score: float,
) -> list[tuple[float, TkpCatalogEntry]]:
    if not entries:
        return []
    try:
        raw_scores = semantic_scorer(
            str(query_name or ""),
            [entry.item_name for entry in entries],
        )
    except Exception as exc:
        raise TkpSemanticScoringError(str(exc)) from exc
    if len(raw_scores) != len(entries):
        raise TkpSemanticScoringError(
            "semantic scorer returned an unexpected score count"
        )

    scored: list[tuple[float, TkpCatalogEntry]] = []
    for raw_score, entry in zip(raw_scores, entries):
        try:
            score = max(0.0, min(float(raw_score), 1.0)) * 100.0
        except (TypeError, ValueError) as exc:
            raise TkpSemanticScoringError(
                "semantic scorer returned a non-numeric score"
            ) from exc
        if score >= min_score:
            scored.append((score, entry))
    return scored

def _same_quantity(left: float, right: float) -> bool:
    if right != right:  # NaN marks missing candidate quantity.
        return False
    return isclose(
        left,
        right,
        rel_tol=QUANTITY_REL_TOL,
        abs_tol=QUANTITY_ABS_TOL,
    )


def _entry_has_usable_participant_price(entry: TkpCatalogEntry) -> bool:
    return (
        positive_float(entry.winner_unit_price_no_vat) is not None
        or positive_float(entry.reserve_unit_price_no_vat) is not None
    )


def _name_score_for_entry(
    query_text: str,
    query_tokens: frozenset[str],
    query_leading_word: str,
    entry: TkpCatalogEntry,
    *,
    min_score: float,
) -> float:
    jaccard = _jaccard(query_tokens, entry._normalized_tokens)
    leading_penalty = (
        0.0
        if same_action_group(query_leading_word, entry._leading_word)
        else LEADING_WORD_MISMATCH_PENALTY
    )
    if _max_possible_score(jaccard, 0.0) - leading_penalty < min_score:
        return 0.0
    return score_names(
        query_text,
        query_tokens,
        entry._normalized_text,
        entry._normalized_tokens,
    ) - leading_penalty


def _resolve_equal_name_candidates(
    candidates: list[TkpCatalogEntry],
    query_rnmc_price: float | None,
) -> tuple[TkpCatalogEntry | None, float | None]:
    if not candidates:
        return None, None
    if len(candidates) == 1:
        return candidates[0], _rnmc_price_delta(candidates[0], query_rnmc_price)

    if query_rnmc_price is not None:
        with_delta = [
            (delta, entry)
            for entry in candidates
            if (delta := _rnmc_price_delta(entry, query_rnmc_price)) is not None
        ]
        if with_delta:
            best_delta = min(delta for delta, _ in with_delta)
            closest = [
                entry
                for delta, entry in with_delta
                if isclose(delta, best_delta, rel_tol=0.0, abs_tol=PRICE_TIE_ABS_TOL)
            ]
            if len(closest) == 1:
                return closest[0], best_delta
            candidates = closest

    priced = [
        (entry, _paired_output_prices(entry, entry.unit))
        for entry in candidates
    ]
    signatures = {
        (
            round(float(result[0]), 9) if result[0] is not None else None,
            round(float(result[1]), 9) if result[1] is not None else None,
        )
        for _, result in priced
        if result[0] is not None or result[1] is not None
    }
    if len(signatures) == 1:
        return min(candidates, key=lambda entry: entry.item_id), None
    return None, None


def _rnmc_price_delta(
    entry: TkpCatalogEntry,
    query_rnmc_price: float | None,
) -> float | None:
    if query_rnmc_price is None:
        return None
    candidate_price = normalized_unit_price(entry.rnmc_unit_price_no_vat, entry.unit)
    if candidate_price is None:
        return None
    return abs(candidate_price - query_rnmc_price) / query_rnmc_price


def _paired_output_prices(
    entry: TkpCatalogEntry,
    target_unit: object,
) -> tuple[float | None, float | None]:
    """Return winner and reserve unit prices converted to the RNMC analog unit."""
    conversion = compatible_unit_conversion(target_unit, entry.unit)
    if conversion is None:
        return None, None
    winner = positive_float(entry.winner_unit_price_no_vat)
    reserve = positive_float(entry.reserve_unit_price_no_vat)
    winner_price = winner * conversion.price_factor if winner is not None else None
    reserve_price = reserve * conversion.price_factor if reserve is not None else None
    return winner_price, reserve_price


def find_best_tkp_matches(
    work_name: object,
    unit: object,
    index: list[TkpCatalogEntry],
    *,
    min_score: float = DEFAULT_MIN_SCORE,
    limit: int = DEFAULT_LIMIT,
) -> list[TkpMatch]:
    """Rank `index` against `work_name`, best first, score >= min_score only.

    `unit` is optional context: an exact NormUnit match adds a small fixed
    bonus (a tie-breaker between otherwise-similar names, not a hard
    filter - TKP unit text is sometimes formatted differently even for a
    genuinely matching item, e.g. "100 м2" vs "м2").
    """
    query_text, query_tokens = normalize_for_matching(work_name)
    if not query_tokens:
        return []
    query_unit = NormUnit(unit) if unit else ""
    query_leading_word = leading_action_word(query_text)

    ranked: list[tuple[float, TkpCatalogEntry]] = []
    for entry in index:
        jaccard = _jaccard(query_tokens, entry._normalized_tokens)
        unit_bonus = (
            UNIT_MATCH_BONUS
            if query_unit and entry._normalized_unit and query_unit == entry._normalized_unit
            else 0.0
        )
        leading_penalty = (
            0.0
            if same_action_group(query_leading_word, entry._leading_word)
            else LEADING_WORD_MISMATCH_PENALTY
        )
        if _max_possible_score(jaccard, unit_bonus) - leading_penalty < min_score:
            continue  # cannot clear the bar even with a perfect sequence_ratio
        raw_score = score_names(query_text, query_tokens, entry._normalized_text, entry._normalized_tokens)
        raw_score += unit_bonus - leading_penalty
        if raw_score >= min_score:
            ranked.append((raw_score, entry))

    # Sort by the *uncapped* score so the unit bonus can still break ties
    # between two otherwise-identical 100.0-base-score names; only the
    # score shown to the caller is clamped to a 0-100 range.
    ranked.sort(key=lambda pair: (-pair[0], pair[1].item_id))
    return [
        TkpMatch(entry=entry, score=max(0.0, min(raw_score, 100.0)))
        for raw_score, entry in ranked[:limit]
    ]
