"""Deterministic lexical matching against the TKP winner catalog.

Per docs/ROADMAP.md ("Future phase - multi-source analogs") and
docs/AGENTS.md rule 2 ("No LLM calls inside the matching/pricing functions
themselves"), this is intentionally NOT semantic/embedding search: it is a
pure token-overlap + character-sequence similarity score over normalized
Russian work-item names, isolated from core/matching.py's exact
GESN/FER/TER-code pipeline. Given the same two inputs it always returns the
same score - no external model, no network call, nothing to cache for
reproducibility because there is nothing non-deterministic to begin with.

Scoring is a blend of two signals so it tolerates both word-order
differences and minor wording/suffix differences:
  - token Jaccard similarity (order-independent, robust to reordering)
  - difflib.SequenceMatcher ratio over the normalized text (order-sensitive,
    catches near-duplicate phrasing that token overlap alone would miss,
    e.g. differing noun endings)
A small bonus is added when the units match (via the existing NormUnit),
since two identically-worded items with different units are rarely the
same real-world work.

A leading-action-word penalty (see ACTION_SYNONYM_GROUPS) subtracts points
when the first meaningful word of the two names names a different
operation (e.g. "укладка" vs "протаскивание"). This exists because token
overlap alone cannot tell "Укладка трубопроводов ... 110 мм" apart from
"Протаскивание в футляр ... 110 мм" - the pipe/diameter tokens dominate the
Jaccard score even though the two are different construction operations
with different pricing. Verified against a real 530-row estimate matched
against a real 2826-item TKP catalog (see tests/test_tkp_matching.py and
project history) before this penalty was added; it removed that specific
false-positive class without suppressing genuine synonyms like
"демонтаж" vs "разборка" (both are members of the same group below).
"""

from dataclasses import dataclass
from difflib import SequenceMatcher

from core.normalize import NormUnit

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
    """One catalog item pre-normalized for repeated matching queries.

    Build once per estimate run with `build_tkp_catalog_index` and reuse
    across every row being matched, instead of re-normalizing the whole
    catalog for every query.
    """

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


@dataclass(frozen=True)
class TkpMatch:
    """One scored candidate returned by `find_best_tkp_matches`."""

    entry: TkpCatalogEntry
    score: float


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
