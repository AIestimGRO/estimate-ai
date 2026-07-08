"""FastAPI application over the matching pipeline.

Thin HTTP layer only: it accepts two uploads, delegates to
app.services.write_result.run_and_write, and renders the outcome. All
matching / pricing / reading logic stays in the tested core and service
modules. Uploaded files (and the produced `WA` workbook) are kept per session
token so the multi-sheet choice can re-run without a re-upload.

Flow:
  GET  /                    -> upload form
  GET  /admin               -> admin dashboard (placeholder sections)
  GET  /admin/{section}     -> admin section placeholder
  POST /run                 -> save uploads, run, render result / sheet choice
  GET  /run?token=&sheet=   -> re-run stored uploads for a chosen sheet
  GET  /download?token=     -> download the produced WA workbook
"""

import tempfile
import uuid
from urllib.parse import quote
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse

from app.services.catalog_source import CatalogNotAvailableError, database_has_catalog
from app.services.read_estimate import (
    EstimateReadError,
    KeyDataNotFoundError,
    MultipleSheetsError,
)
from app.services.write_result import run_and_write
from core.storage.catalog import (
    count_catalog_rows,
    import_legacy_file_log,
    list_catalog_sources,
    list_imported_files,
)
from core.storage.risk_log import (
    STATUS_OPEN,
    approve_risk,
    get_price_risk,
    list_gesn_exceptions,
    list_price_risks,
)
from core.storage.rules import (
    list_name_exclusion_rules,
    list_task_color_entries,
    set_name_exclusion_rule_enabled,
    set_task_color_enabled,
    upsert_name_exclusion_rule,
    upsert_task_color_entry,
)
from core.storage.connection import connect, default_database_path, init_database
from app.web.rendering import (
    ADMIN_SECTION_SLUGS,
    XLSX_MIME,
    render_admin_approvals,
    render_admin_gesn_exceptions,
    render_admin_imports,
    render_admin_name_exclusions,
    render_admin_index,
    render_admin_task_colors,
    render_admin_risks,
    render_admin_section,
    render_admin_settings,
    render_admin_sources,
    render_choice,
    render_error,
    render_index,
    render_result,
)

MAX_UPLOAD_BYTES = 64 * 1024 * 1024
VBA_DATE_BASE = date(1899, 12, 30)
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

    @app.get("/admin", response_class=HTMLResponse)
    def admin_index() -> HTMLResponse:
        return HTMLResponse(render_admin_index())

    @app.get("/admin/{section_slug}", response_class=HTMLResponse)
    def admin_section(section_slug: str, message: str = "") -> HTMLResponse:
        if section_slug not in ADMIN_SECTION_SLUGS:
            return HTMLResponse(
                render_error(
                    "\u0420\u0430\u0437\u0434\u0435\u043b \u0430\u0434\u043c\u0438\u043d\u043a\u0438 \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d",
                    f"\u041d\u0435\u0438\u0437\u0432\u0435\u0441\u0442\u043d\u044b\u0439 \u0440\u0430\u0437\u0434\u0435\u043b: {section_slug}",
                ),
                status_code=404,
            )

        connection = connect(default_database_path())
        try:
            init_database(connection)
            if section_slug == "sources":
                sources = list_catalog_sources(connection)
                return HTMLResponse(render_admin_sources(sources))
            if section_slug == "imports":
                imports = list_imported_files(connection)
                return HTMLResponse(render_admin_imports(imports, notice=message))
            if section_slug == "risks":
                risks = list_price_risks(connection)
                return HTMLResponse(render_admin_risks(risks))
            if section_slug == "gesn-exceptions":
                exceptions = list_gesn_exceptions(connection)
                return HTMLResponse(render_admin_gesn_exceptions(exceptions))
            if section_slug == "approvals":
                risks = list_price_risks(connection, status="open")
                return HTMLResponse(render_admin_approvals(risks))
            if section_slug == "task-colors":
                entries = list_task_color_entries(connection)
                return HTMLResponse(render_admin_task_colors(entries))
            if section_slug == "name-exclusions":
                rules = list_name_exclusion_rules(connection)
                return HTMLResponse(render_admin_name_exclusions(rules))
            if section_slug == "settings":
                settings_rows = _admin_settings_rows(connection)
                return HTMLResponse(render_admin_settings(settings_rows))
        finally:
            connection.close()

        return HTMLResponse(render_admin_section(section_slug))


    @app.post("/admin/imports/file-log", response_class=HTMLResponse)
    async def admin_import_file_log(file_log: UploadFile = File(...)):
        raw_name = file_log.filename or ""
        if _safe_name(raw_name) == "":
            connection = connect(default_database_path())
            try:
                init_database(connection)
                imports = list_imported_files(connection)
                return HTMLResponse(
                    render_admin_imports(
                        imports,
                        error="Файл File_Log.xlsx не выбран.",
                    ),
                    status_code=400,
                )
            finally:
                connection.close()

        upload_dir = app.state.app_state.base_dir / "admin_imports"
        upload_dir.mkdir(parents=True, exist_ok=True)
        saved_path = _save(
            upload_dir,
            f"file_log_{uuid.uuid4().hex}{Path(_safe_name(raw_name)).suffix or '.xlsx'}",
            await file_log.read(),
        )
        connection = connect(default_database_path())
        try:
            init_database(connection)
            try:
                result = import_legacy_file_log(connection, saved_path)
            except Exception as exc:
                imports = list_imported_files(connection)
                return HTMLResponse(
                    render_admin_imports(
                        imports,
                        error=f"Не удалось импортировать File_Log.xlsx: {exc}",
                    ),
                    status_code=400,
                )
        finally:
            connection.close()
        message = (
            f"File_Log.xlsx импортирован: {result.rows_imported} записей, "
            f"дубликатов имен: {result.duplicates}."
        )
        return RedirectResponse(f"/admin/imports?message={quote(message)}", status_code=303)

    @app.post("/admin/approvals/approve", response_class=HTMLResponse)
    def admin_approval_approve(exception_key: str = Form("")):
        normalized_key = exception_key.strip()
        connection = connect(default_database_path())
        try:
            init_database(connection)
            risk = get_price_risk(
                connection,
                exception_key=normalized_key,
                status=STATUS_OPEN,
            )
            if risk is None:
                open_risks = list_price_risks(connection, status=STATUS_OPEN)
                return HTMLResponse(
                    render_admin_approvals(
                        open_risks,
                        error="Открытый риск для одобрения не найден.",
                    ),
                    status_code=404,
                )
            if risk.min_price is None or risk.max_price is None:
                open_risks = list_price_risks(connection, status=STATUS_OPEN)
                return HTMLResponse(
                    render_admin_approvals(
                        open_risks,
                        error="У риска нет min/max цен для одобрения диапазона.",
                    ),
                    status_code=400,
                )
            approve_risk(
                connection,
                normalized_key,
                proposed_min=risk.min_price,
                proposed_max=risk.max_price,
                proposed_date_serial=_today_vba_date_serial(),
            )
        finally:
            connection.close()
        return RedirectResponse("/admin/approvals", status_code=303)


    @app.post("/admin/name-exclusions/add", response_class=HTMLResponse)
    def admin_name_exclusion_add(
        scope: str = Form(""),
        match_mode: str = Form(""),
        pattern: str = Form(""),
        rule_group: str = Form(""),
        comment: str = Form(""),
    ):
        connection = connect(default_database_path())
        try:
            init_database(connection)
            try:
                upsert_name_exclusion_rule(
                    connection,
                    scope=scope,
                    match_mode=match_mode,
                    pattern=pattern,
                    rule_group=rule_group,
                    comment=comment,
                    enabled=True,
                )
            except ValueError as error:
                rules = list_name_exclusion_rules(connection)
                return HTMLResponse(
                    render_admin_name_exclusions(
                        rules,
                        error=_name_exclusion_error_message(error),
                    ),
                    status_code=400,
                )
        finally:
            connection.close()
        return RedirectResponse("/admin/name-exclusions", status_code=303)

    @app.post("/admin/name-exclusions/toggle", response_class=HTMLResponse)
    def admin_name_exclusion_toggle(
        scope: str = Form(""),
        match_mode: str = Form(""),
        pattern: str = Form(""),
        enabled: str = Form("0"),
    ):
        connection = connect(default_database_path())
        try:
            init_database(connection)
            try:
                changed = set_name_exclusion_rule_enabled(
                    connection,
                    scope=scope,
                    match_mode=match_mode,
                    pattern=pattern,
                    enabled=enabled == "1",
                )
            except ValueError as error:
                rules = list_name_exclusion_rules(connection)
                return HTMLResponse(
                    render_admin_name_exclusions(
                        rules,
                        error=_name_exclusion_error_message(error),
                    ),
                    status_code=400,
                )
            if not changed:
                rules = list_name_exclusion_rules(connection)
                return HTMLResponse(
                    render_admin_name_exclusions(
                        rules,
                        error="Правило для изменения не найдено.",
                    ),
                    status_code=404,
                )
        finally:
            connection.close()
        return RedirectResponse("/admin/name-exclusions", status_code=303)

    @app.post("/admin/task-colors/add", response_class=HTMLResponse)
    def admin_task_color_add(
        task_number: str = Form(""),
        reason: str = Form(""),
        comment: str = Form(""),
    ):
        connection = connect(default_database_path())
        try:
            init_database(connection)
            try:
                upsert_task_color_entry(
                    connection,
                    task_number=task_number,
                    reason=reason,
                    comment=comment,
                    enabled=True,
                )
            except ValueError:
                entries = list_task_color_entries(connection)
                return HTMLResponse(
                    render_admin_task_colors(
                        entries,
                        error="Номер задачи обязателен.",
                    ),
                    status_code=400,
                )
        finally:
            connection.close()
        return RedirectResponse("/admin/task-colors", status_code=303)

    @app.post("/admin/task-colors/toggle", response_class=HTMLResponse)
    def admin_task_color_toggle(
        task_number: str = Form(""),
        enabled: str = Form("0"),
    ):
        connection = connect(default_database_path())
        try:
            init_database(connection)
            changed = set_task_color_enabled(
                connection,
                task_number=task_number,
                enabled=enabled == "1",
            )
            if not changed:
                entries = list_task_color_entries(connection)
                return HTMLResponse(
                    render_admin_task_colors(
                        entries,
                        error="Задача для изменения не найдена.",
                    ),
                    status_code=404,
                )
        finally:
            connection.close()
        return RedirectResponse("/admin/task-colors", status_code=303)

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
            database_path=default_database_path(),
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



def _admin_settings_rows(connection) -> list[tuple[str, str]]:
    database_path = default_database_path()
    sources = list_catalog_sources(connection)
    imports = list_imported_files(connection)
    risks = list_price_risks(connection)
    exceptions = list_gesn_exceptions(connection)
    task_colors = list_task_color_entries(connection)
    name_rules = list_name_exclusion_rules(connection)
    return [
        ("Database path", str(database_path)),
        ("Database exists", "yes" if database_path.is_file() else "no"),
        ("Catalog rows", str(count_catalog_rows(connection))),
        ("Sources", str(len(sources))),
        ("Imported files", str(len(imports))),
        ("Open risks", str(len([risk for risk in risks if risk.status == "open"]))),
        ("Approved risks", str(len([risk for risk in risks if risk.status == "approved"]))),
        ("GESN exceptions", str(len(exceptions))),
        ("Task color entries", str(len(task_colors))),
        ("Name exclusion rules", str(len(name_rules))),
    ]


def _name_exclusion_error_message(error: ValueError) -> str:
    message = str(error)
    if message == "pattern is required":
        return "Pattern обязателен."
    if message == "invalid scope":
        return "Scope должен быть SMETA, CATALOG или BOTH."
    if message == "invalid match_mode":
        return "Match mode должен быть CONTAINS или ALL_WORDS."
    return "Правило исключения заполнено некорректно."

def _today_vba_date_serial() -> float:
    return float((date.today() - VBA_DATE_BASE).days)


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
