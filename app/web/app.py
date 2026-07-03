"""FastAPI application over the matching pipeline.

Thin HTTP layer only: it accepts two uploads, delegates to
app.services.write_result.run_and_write, and renders the outcome. All
matching / pricing / reading logic stays in the tested core and service
modules. Uploaded files (and the produced `WA` workbook) are kept per session
token so the multi-sheet choice can re-run without a re-upload.

Flow:
  GET  /                    -> upload form
  POST /run                 -> save uploads, run, render result / sheet choice
  GET  /run?token=&sheet=   -> re-run stored uploads for a chosen sheet
  GET  /download?token=     -> download the produced WA workbook
"""

import tempfile
import uuid
from dataclasses import dataclass, field
from pathlib import Path

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, HTMLResponse

from app.services.catalog_source import CatalogNotAvailableError, database_has_catalog
from app.services.read_estimate import (
    EstimateReadError,
    KeyDataNotFoundError,
    MultipleSheetsError,
)
from app.services.write_result import run_and_write
from app.web.rendering import (
    XLSX_MIME,
    render_choice,
    render_error,
    render_index,
    render_result,
)

MAX_UPLOAD_BYTES = 64 * 1024 * 1024
_INVALID = object()


@dataclass
class UploadRecord:
    """Files and options stored for one upload token (survives sheet choice)."""

    directory: Path
    catalog_path: Path | None
    estimate_path: Path
    estimate_name: str
    coefficient: float | None
    use_database_catalog: bool = False
    output_path: Path | None = None
    output_name: str = ""


@dataclass
class AppState:
    """Shared server state (upload store + working directory)."""

    base_dir: Path
    store: dict[str, UploadRecord] = field(default_factory=dict)


def create_app(base_dir: str | Path | None = None) -> FastAPI:
    """Build the FastAPI app; `base_dir` holds per-session uploads/outputs."""
    resolved_base = (
        Path(base_dir)
        if base_dir is not None
        else Path(tempfile.gettempdir()) / "estimate_ai_web"
    )
    resolved_base.mkdir(parents=True, exist_ok=True)

    app = FastAPI(title="Estimate AI", docs_url=None, redoc_url=None)
    app.state.app_state = AppState(base_dir=resolved_base)

    @app.get("/", response_class=HTMLResponse)
    def index() -> HTMLResponse:
        return HTMLResponse(render_index())

    @app.post("/run", response_class=HTMLResponse)
    async def run(
        catalog: UploadFile | None = File(None),
        estimate: UploadFile | None = File(None),
        coefficient: str = Form(""),
    ) -> HTMLResponse:
        state: AppState = app.state.app_state

        coef = _parse_coefficient(coefficient)
        if coef is _INVALID:
            return HTMLResponse(
                render_index("\u041a\u043e\u044d\u0444\u0444\u0438\u0446\u0438\u0435\u043d\u0442 \u0434\u043e\u043b\u0436\u0435\u043d \u0431\u044b\u0442\u044c \u0447\u0438\u0441\u043b\u043e\u043c, \u043d\u0430\u043f\u0440\u0438\u043c\u0435\u0440 1.15."),
                status_code=400,
            )

        if estimate is None:
            return HTMLResponse(
                render_index("\u041d\u0443\u0436\u043d\u043e \u0432\u044b\u0431\u0440\u0430\u0442\u044c \u0444\u0430\u0439\u043b \u0441\u043c\u0435\u0442\u044b."),
                status_code=400,
            )

        estimate_bytes = await estimate.read()
        if not estimate_bytes:
            return HTMLResponse(
                render_index("\u041d\u0443\u0436\u043d\u043e \u0432\u044b\u0431\u0440\u0430\u0442\u044c \u0444\u0430\u0439\u043b \u0441\u043c\u0435\u0442\u044b."),
                status_code=400,
            )

        catalog_bytes = b""
        if catalog is not None:
            catalog_bytes = await catalog.read()

        use_database_catalog = False
        catalog_path: Path | None = None
        if catalog_bytes:
            pass
        elif database_has_catalog():
            use_database_catalog = True
        else:
            return HTMLResponse(
                render_index(
                    "\u0417\u0430\u0433\u0440\u0443\u0437\u0438\u0442\u0435 \u0431\u0430\u0437\u0443 \u0430\u043d\u0430\u043b\u043e\u0433\u043e\u0432 "
                    "(\u0438\u043b\u0438 \u0432\u044b\u043f\u043e\u043b\u043d\u0438\u0442\u0435 import-catalog)."
                ),
                status_code=400,
            )

        token = uuid.uuid4().hex
        directory = state.base_dir / token
        directory.mkdir(parents=True, exist_ok=True)
        if catalog_bytes:
            catalog_path = _save(directory, "catalog.xlsx", catalog_bytes)
        estimate_name = _safe_name(estimate.filename) or "estimate.xlsx"
        estimate_path = _save(directory, estimate_name, estimate_bytes)

        state.store[token] = UploadRecord(
            directory=directory,
            catalog_path=catalog_path,
            estimate_path=estimate_path,
            estimate_name=estimate_name,
            coefficient=None if coef is _INVALID else coef,  # type: ignore[arg-type]
            use_database_catalog=use_database_catalog,
        )
        return _process(state, token, selected_sheet=None)

    @app.get("/run", response_class=HTMLResponse)
    def run_sheet(token: str, sheet: str | None = None) -> HTMLResponse:
        state: AppState = app.state.app_state
        if token not in state.store:
            return HTMLResponse(
                render_error(
                    "\u0421\u0435\u0441\u0441\u0438\u044f \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u0430",
                    "\u0417\u0430\u0433\u0440\u0443\u0437\u0438\u0442\u0435 \u0444\u0430\u0439\u043b\u044b \u0437\u0430\u043d\u043e\u0432\u043e.",
                ),
                status_code=404,
            )
        return _process(state, token, selected_sheet=sheet)

    @app.get("/download")
    def download(token: str):
        state: AppState = app.state.app_state
        record = state.store.get(token)
        if record is None or record.output_path is None or not record.output_path.exists():
            return HTMLResponse(
                render_error(
                    "\u0424\u0430\u0439\u043b \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d",
                    "\u0420\u0435\u0437\u0443\u043b\u044c\u0442\u0430\u0442 \u043d\u0435\u0434\u043e\u0441\u0442\u0443\u043f\u0435\u043d, \u043f\u043e\u0432\u0442\u043e\u0440\u0438\u0442\u0435 \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0443.",
                ),
                status_code=404,
            )
        return FileResponse(
            record.output_path,
            media_type=XLSX_MIME,
            filename=record.output_name,
        )

    return app


def _process(state: AppState, token: str, selected_sheet: str | None) -> HTMLResponse:
    record = state.store[token]
    try:
        outcome = run_and_write(
            record.catalog_path,
            record.estimate_path,
            record.directory / _wa_name(record.estimate_name),
            selected_sheet_title=selected_sheet,
            regional_coefficient=record.coefficient,
        )
    except MultipleSheetsError as error:
        return HTMLResponse(render_choice(token, error.candidates))
    except KeyDataNotFoundError as error:
        return HTMLResponse(
            render_error(
                "\u041a\u043b\u044e\u0447\u0435\u0432\u044b\u0435 \u0434\u0430\u043d\u043d\u044b\u0435 \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u044b",
                f"\u041d\u0430 \u043b\u0438\u0441\u0442\u0435 \u00ab{error.sheet_title}\u00bb \u043d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043e\u043f\u0440\u0435\u0434\u0435\u043b\u0438\u0442\u044c \u043e\u0431\u044f\u0437\u0430\u0442\u0435\u043b\u044c\u043d\u044b\u0435 \u043f\u043e\u043b\u044f (\u043a\u043e\u0434, \u0435\u0434\u0438\u043d\u0438\u0446\u0430 \u0438\u0437\u043c\u0435\u0440\u0435\u043d\u0438\u044f, \u0431\u0430\u0437\u043e\u0432\u0430\u044f \u0446\u0435\u043d\u0430).",
                detail=error.report,
            ),
            status_code=422,
        )
    except EstimateReadError as error:
        return HTMLResponse(
            render_error("\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043f\u0440\u043e\u0447\u0438\u0442\u0430\u0442\u044c \u0441\u043c\u0435\u0442\u0443", str(error)),
            status_code=400,
        )
    except CatalogNotAvailableError as error:
        return HTMLResponse(
            render_error(
                "\u0411\u0430\u0437\u0430 \u0430\u043d\u0430\u043b\u043e\u0433\u043e\u0432 \u043d\u0435\u0434\u043e\u0441\u0442\u0443\u043f\u043d\u0430",
                str(error),
            ),
            status_code=400,
        )

    record.output_path = outcome.output_path
    record.output_name = _wa_name(record.estimate_name)
    return HTMLResponse(render_result(token, record.output_name, outcome))


def _parse_coefficient(text: str) -> float | None | object:
    value = (text or "").strip()
    if not value:
        return None
    try:
        return float(value.replace(",", "."))
    except ValueError:
        return _INVALID


def _save(directory: Path, name: str, content: bytes) -> Path:
    path = directory / name
    path.write_bytes(content)
    return path


def _safe_name(filename: str | None) -> str:
    if not filename:
        return ""
    return Path(filename.replace("\\", "/")).name


def _wa_name(estimate_name: str) -> str:
    source = Path(estimate_name)
    suffix = source.suffix or ".xlsx"
    return f"{source.stem} WA{suffix}"
