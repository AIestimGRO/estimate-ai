"""Isolated TKP shadow comparison with optional local semantic models."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from functools import lru_cache
from pathlib import Path
from typing import Protocol

from core.tkp_candidate_rules import CandidateRuleResult, evaluate_tkp_candidate
from core.tkp_matching import (
    DEFAULT_LIMIT,
    DEFAULT_MIN_SCORE,
    TkpCatalogEntry,
    TkpMatch,
    find_best_tkp_matches,
)


MODEL_QWEN3 = "qwen3"
MODEL_BGE_M3 = "bge-m3"
STRICT_MIN_LEXICAL_SCORE = 15.0


class ShadowModelStatus(str, Enum):
    READY = "ready"
    UNAVAILABLE = "unavailable"
    ERROR = "error"


class SemanticBackend(Protocol):
    model_key: str
    display_name: str

    def score(self, query: str, candidates: list[str]) -> list[float]:
        """Return one cosine-like 0-1 score per candidate."""


@dataclass(frozen=True)
class ShadowCandidate:
    entry: TkpCatalogEntry
    lexical_score: float
    semantic_score: float | None
    normalized_unit_price: float
    unit_price_factor: float


@dataclass(frozen=True)
class RejectedShadowCandidate:
    entry: TkpCatalogEntry
    reason: str


@dataclass(frozen=True)
class ShadowModelResult:
    model_key: str
    display_name: str
    status: ShadowModelStatus
    message: str
    candidates: list[ShadowCandidate]


@dataclass(frozen=True)
class ShadowComparison:
    live_candidates: list[TkpMatch]
    strict_candidates: list[ShadowCandidate]
    rejected_candidates: list[RejectedShadowCandidate]
    semantic_models: list[ShadowModelResult]
    shadow_only: bool = True


def build_shadow_comparison(
    work_name: object,
    unit: object,
    index: list[TkpCatalogEntry],
    *,
    semantic_backends: list[SemanticBackend] | None = None,
    limit: int = DEFAULT_LIMIT,
) -> ShadowComparison:
    live = find_best_tkp_matches(
        work_name,
        unit,
        index,
        min_score=DEFAULT_MIN_SCORE,
        limit=limit,
    )
    lexical = find_best_tkp_matches(
        work_name,
        unit,
        index,
        min_score=0.0,
        limit=max(len(index), limit),
    )
    score_by_id = {match.entry.item_id: match.score for match in lexical}
    accepted: list[tuple[TkpCatalogEntry, CandidateRuleResult, float]] = []
    rejected: list[RejectedShadowCandidate] = []
    for entry in index:
        rule = evaluate_tkp_candidate(work_name, unit, entry)
        lexical_score = score_by_id.get(entry.item_id, 0.0)
        if not rule.accepted:
            rejected.append(RejectedShadowCandidate(entry, rule.reason))
            continue
        accepted.append((entry, rule, lexical_score))

    strict = sorted(
        (
            _shadow_candidate(entry, rule, lexical_score)
            for entry, rule, lexical_score in accepted
            if lexical_score >= STRICT_MIN_LEXICAL_SCORE
        ),
        key=lambda row: (-row.lexical_score, row.entry.item_id),
    )[:limit]
    semantic_results = [
        _run_semantic_backend(
            backend,
            str(work_name or ""),
            accepted,
            limit=limit,
        )
        for backend in (semantic_backends or [])
    ]
    rejected.sort(key=lambda row: row.entry.item_id)
    return ShadowComparison(
        live_candidates=live,
        strict_candidates=strict,
        rejected_candidates=rejected,
        semantic_models=semantic_results,
    )


def discover_local_semantic_backends(
    models_dir: str | Path,
) -> tuple[list[SemanticBackend], list[ShadowModelResult]]:
    root = Path(models_dir)
    definitions = (
        (MODEL_QWEN3, "Qwen3-Embedding-0.6B", root / "qwen3-embedding-0.6b"),
        (MODEL_BGE_M3, "BGE-M3", root / "bge-m3"),
    )
    backends: list[SemanticBackend] = []
    statuses: list[ShadowModelResult] = []
    for model_key, display_name, model_path in definitions:
        if not model_path.is_dir():
            statuses.append(
                ShadowModelResult(
                    model_key,
                    display_name,
                    ShadowModelStatus.UNAVAILABLE,
                    f"Local model directory not found: {model_path}",
                    [],
                )
            )
            continue
        try:
            backends.append(
                SentenceTransformerBackend(
                    model_key,
                    display_name,
                    model_path,
                )
            )
        except Exception as exc:
            statuses.append(
                ShadowModelResult(
                    model_key,
                    display_name,
                    ShadowModelStatus.ERROR,
                    str(exc),
                    [],
                )
            )
    return backends, statuses


class SentenceTransformerBackend:
    """Adapter for a model already downloaded to a local directory."""

    def __init__(
        self,
        model_key: str,
        display_name: str,
        model_path: str | Path,
    ) -> None:
        self.model_key = model_key
        self.display_name = display_name
        self._model = _load_cpu_sentence_transformer(
            str(Path(model_path).resolve())
        )

    def score(self, query: str, candidates: list[str]) -> list[float]:
        embeddings = self._model.encode(
            [query, *candidates],
            normalize_embeddings=True,
            show_progress_bar=False,
        )
        query_embedding = embeddings[0]
        return [
            float(query_embedding @ candidate_embedding)
            for candidate_embedding in embeddings[1:]
        ]


@lru_cache(maxsize=2)
def _load_cpu_sentence_transformer(model_path: str):
    """Load a local model once and force CPU-safe float32 weights."""
    try:
        import torch
        from sentence_transformers import SentenceTransformer
    except ImportError as exc:
        raise RuntimeError(
            "sentence-transformers is not installed; live matching is unchanged"
        ) from exc

    model = SentenceTransformer(
        model_path,
        device="cpu",
        local_files_only=True,
        trust_remote_code=True,
    )
    model.to(dtype=torch.float32)
    model.eval()
    return model


def _run_semantic_backend(
    backend: SemanticBackend,
    query: str,
    accepted: list[tuple[TkpCatalogEntry, CandidateRuleResult, float]],
    *,
    limit: int,
) -> ShadowModelResult:
    try:
        scores = backend.score(
            query,
            [entry.item_name for entry, _, _ in accepted],
        )
        if len(scores) != len(accepted):
            raise ValueError("semantic backend returned an unexpected score count")
        candidates = [
            _shadow_candidate(
                entry,
                rule,
                lexical_score,
                semantic_score=max(0.0, min(float(score), 1.0)) * 100.0,
            )
            for (entry, rule, lexical_score), score in zip(accepted, scores)
        ]
        candidates.sort(
            key=lambda row: (
                -(row.semantic_score or 0.0),
                -row.lexical_score,
                row.entry.item_id,
            )
        )
        return ShadowModelResult(
            backend.model_key,
            backend.display_name,
            ShadowModelStatus.READY,
            "",
            candidates[:limit],
        )
    except Exception as exc:
        return ShadowModelResult(
            backend.model_key,
            backend.display_name,
            ShadowModelStatus.ERROR,
            str(exc),
            [],
        )


def _shadow_candidate(
    entry: TkpCatalogEntry,
    rule: CandidateRuleResult,
    lexical_score: float,
    *,
    semantic_score: float | None = None,
) -> ShadowCandidate:
    if rule.normalized_unit_price is None or rule.unit_conversion is None:
        raise ValueError("accepted candidate is missing normalized price data")
    return ShadowCandidate(
        entry=entry,
        lexical_score=lexical_score,
        semantic_score=semantic_score,
        normalized_unit_price=rule.normalized_unit_price,
        unit_price_factor=rule.unit_conversion.price_factor,
    )
