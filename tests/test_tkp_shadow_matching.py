"""Tests for isolated TKP shadow matching and strict candidate checks."""

import sys
from types import SimpleNamespace
from dataclasses import dataclass

from app.services.tkp_shadow import (
    MODEL_BGE_M3,
    MODEL_QWEN3,
    SentenceTransformerBackend,
    ShadowModelStatus,
    _load_cpu_sentence_transformer,
    build_shadow_comparison,
)
from core.tkp_matching import build_tkp_catalog_index


@dataclass(frozen=True)
class _Item:
    id: int
    item_name: str
    unit: str
    winner_unit_price_no_vat: float
    winner_name: str = "Winner"
    source_file_name: str = "source.xlsx"
    task_no: str = "TASK"
    section_name: str = ""
    subsection_name: str = ""


class _FakeSemanticBackend:
    model_key = MODEL_QWEN3
    display_name = "Fake Qwen"

    def score(self, query: str, candidates: list[str]) -> list[float]:
        scores = []
        for candidate in candidates:
            scores.append(0.95 if "mineral wool" in candidate else 0.2)
        return scores


def test_local_semantic_backend_forces_cpu_float32_and_reuses_model(
    monkeypatch,
    tmp_path,
) -> None:
    calls: list[tuple[str, dict[str, object]]] = []

    class _FakeModel:
        def __init__(self) -> None:
            self.to_calls: list[dict[str, object]] = []
            self.eval_calls = 0

        def to(self, **kwargs):
            self.to_calls.append(kwargs)
            return self

        def eval(self):
            self.eval_calls += 1
            return self

    fake_model = _FakeModel()

    def fake_sentence_transformer(path: str, **kwargs):
        calls.append((path, kwargs))
        return fake_model

    monkeypatch.setitem(
        sys.modules,
        "torch",
        SimpleNamespace(float32="float32"),
    )
    monkeypatch.setitem(
        sys.modules,
        "sentence_transformers",
        SimpleNamespace(SentenceTransformer=fake_sentence_transformer),
    )
    _load_cpu_sentence_transformer.cache_clear()
    model_dir = tmp_path / "model"
    model_dir.mkdir()

    first = SentenceTransformerBackend("one", "One", model_dir)
    second = SentenceTransformerBackend("two", "Two", model_dir)

    assert first._model is second._model
    assert len(calls) == 1
    assert calls[0][1] == {
        "device": "cpu",
        "local_files_only": True,
        "trust_remote_code": True,
    }
    assert fake_model.to_calls == [{"dtype": "float32"}]
    assert fake_model.eval_calls == 1
    _load_cpu_sentence_transformer.cache_clear()


def test_shadow_matching_converts_prefixed_units_and_blocks_demolition() -> None:
    items = [
        _Item(
            1,
            "Metal sheet insulation covering",
            "100 m2",
            315_723.62,
            section_name="Restoration of insulation",
        ),
        _Item(
            2,
            "Metal sheet insulation covering",
            "100 m2",
            100_000.0,
            section_name="Demolition of insulation",
        ),
    ]
    index = build_tkp_catalog_index(items)

    result = build_shadow_comparison(
        "Metal sheet insulation covering",
        "m2",
        index,
        semantic_backends=[],
    )

    assert result.strict_candidates[0].entry.item_id == 1
    assert result.strict_candidates[0].normalized_unit_price == 3157.2362
    assert {row.entry.item_id for row in result.strict_candidates} == {1}
    assert any(
        row.entry.item_id == 2 and row.reason == "work_type_conflict"
        for row in result.rejected_candidates
    )


def test_shadow_semantic_backend_surfaces_low_lexical_candidate() -> None:
    items = [
        _Item(1, "mineral wool pipe insulation mats", "m3", 26_863.52),
        _Item(2, "silicone cable penetration sealing", "m", 2_000.0),
    ]
    index = build_tkp_catalog_index(items)

    result = build_shadow_comparison(
        "thermal insulation with stone fiber slabs",
        "m3",
        index,
        semantic_backends=[_FakeSemanticBackend()],
    )

    qwen = next(model for model in result.semantic_models if model.model_key == MODEL_QWEN3)
    assert qwen.status == ShadowModelStatus.READY
    assert qwen.candidates[0].entry.item_id == 1
    assert qwen.candidates[0].semantic_score == 95.0
    assert all(model.model_key != MODEL_BGE_M3 for model in result.semantic_models)


def test_shadow_matching_keeps_live_match_result_separate() -> None:
    index = build_tkp_catalog_index(
        [_Item(1, "assembly of small brackets", "t", 199_482.0)]
    )

    result = build_shadow_comparison(
        "installation of small structures",
        "t",
        index,
        semantic_backends=[],
    )

    assert result.live_candidates == []
    assert result.shadow_only is True


def test_strict_shadow_hides_zero_relevance_candidates() -> None:
    index = build_tkp_catalog_index(
        [_Item(1, "painting walls with acrylic paint", "m", 1_000.0)]
    )

    result = build_shadow_comparison(
        "sealing joints with silicone sealant",
        "m",
        index,
        semantic_backends=[],
    )

    assert result.strict_candidates == []
