from pathlib import Path
from types import SimpleNamespace

from app.services import postprocessing


class _FakeWorksheet:
    max_column = 3

    def cell(self, *args, **kwargs):
        raise AssertionError("read-only persistence must not use random cell access")

    def iter_rows(
        self,
        *,
        min_row: int,
        max_row: int,
        min_col: int,
        max_col: int,
        values_only: bool,
    ):
        assert min_col == 1
        assert max_col == 3
        assert values_only is True
        rows = {
            1: ("Source", "Average", "Other"),
            5: (10.0, "=AVERAGE(A5)", "x"),
        }
        for row_number in range(min_row, max_row + 1):
            yield rows.get(row_number, (None, None, None))


class _FakeWorkbook:
    def __init__(self) -> None:
        self.worksheet = _FakeWorksheet()
        self.closed = False

    def __getitem__(self, title: str):
        assert title == "Estimate"
        return self.worksheet

    def close(self) -> None:
        self.closed = True


def test_persist_processing_job_scans_read_only_sheet_sequentially(
    tmp_path: Path,
    monkeypatch,
) -> None:
    workbook = _FakeWorkbook()
    captured: dict[str, object] = {}

    monkeypatch.setattr(
        postprocessing,
        "load_workbook",
        lambda *args, **kwargs: workbook,
    )
    monkeypatch.setattr(
        postprocessing,
        "upsert_processing_job",
        lambda connection, **kwargs: captured.update(job=kwargs),
    )
    monkeypatch.setattr(
        postprocessing,
        "replace_processing_rows",
        lambda connection, job_id, rows: captured.update(rows=rows),
    )

    result_row = SimpleNamespace(
        recommended_price=123.4,
        has_analogs=True,
        risk_result=SimpleNamespace(is_flagged=False),
        has_tkp_analog=False,
        status="matched",
        match_result=SimpleNamespace(reason="matched"),
    )
    outcome = SimpleNamespace(
        output_path=tmp_path / "result.xlsx",
        sheet_title="Estimate",
        row_numbers=(5,),
        write_report=SimpleNamespace(
            header_row=1,
            average_column=2,
            analog_start_column=3,
            analog_column_count=1,
            tkp_start_column=None,
            analog_columns=(),
        ),
        result=SimpleNamespace(
            rows=(result_row,),
            matched_row_count=1,
            flagged_row_count=0,
            tkp_matched_row_count=0,
        ),
        regional_coefficient=1.0,
    )

    postprocessing.persist_processing_job(
        object(),
        job_id="job-1",
        owner_user_id=7,
        estimate_filename="estimate.xlsx",
        source_path=tmp_path / "source.xlsx",
        outcome=outcome,
        region="",
        use_tkp_analogs=False,
    )

    assert workbook.closed
    rows = captured["rows"]
    assert rows == [
        (
            1,
            5,
            [10.0, 123.4, "x"],
            {
                "has_analogs": True,
                "risk": False,
                "has_tkp": False,
                "status": "matched",
                "reason": "matched",
            },
        )
    ]
    job = captured["job"]
    assert job["column_schema"][0]["label"] == "Source"
    assert job["column_schema"][1]["kind"] == "average"
