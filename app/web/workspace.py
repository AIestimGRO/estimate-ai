"""Authentication and desktop post-processing workspace routes."""

from __future__ import annotations

import html
import json
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse

from app.services.postprocessing import build_postprocessed_workbook, resolve_preview_row_values
from core.storage.accounts import (
    ROLE_ADMIN,
    ROLE_SPECIALIST,
    authenticate_user,
    count_users,
    create_session,
    create_user,
    list_users,
    reset_user_password,
    revoke_session,
    set_user_active,
    user_from_session,
)
from core.storage.connection import connect, default_database_path, init_database
from core.storage.rules import upsert_task_color_entry
from core.storage.workspace import (
    REQUEST_ANALOG_CHANGE,
    REQUEST_APPROVED,
    REQUEST_BLUE_TASK,
    REQUEST_OTHER,
    REQUEST_PENDING,
    REQUEST_REJECTED,
    REQUEST_RULE_CHANGE,
    create_change_request,
    get_change_request,
    get_processing_job,
    get_user_preference,
    list_activity_events,
    list_cell_overrides,
    list_change_requests,
    list_processing_jobs,
    list_processing_rows,
    record_activity_event,
    review_change_request,
    save_cell_override,
    save_user_preference,
    touch_job,
)


SESSION_COOKIE = "estimate_ai_session"
SESSION_DAYS = 7
GRID_PREFERENCE_KEY = "workspace_grid_v1"

TEXT = {
    "login": "\u0412\u0445\u043e\u0434",
    "login_label": "\u041b\u043e\u0433\u0438\u043d",
    "password": "\u041f\u0430\u0440\u043e\u043b\u044c",
    "sign_in": "\u0412\u043e\u0439\u0442\u0438",
    "bad_login": "\u041d\u0435\u0432\u0435\u0440\u043d\u044b\u0439 \u043b\u043e\u0433\u0438\u043d \u0438\u043b\u0438 \u043f\u0430\u0440\u043e\u043b\u044c.",
    "setup": "\u041f\u0435\u0440\u0432\u0438\u0447\u043d\u0430\u044f \u043d\u0430\u0441\u0442\u0440\u043e\u0439\u043a\u0430",
    "full_name": "\u0424\u0418\u041e",
    "create_admin": "\u0421\u043e\u0437\u0434\u0430\u0442\u044c \u0430\u0434\u043c\u0438\u043d\u0438\u0441\u0442\u0440\u0430\u0442\u043e\u0440\u0430",
    "jobs": "\u041c\u043e\u0438 \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0438",
    "all_jobs": "\u0412\u0441\u0435 \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0438",
    "new_run": "\u041d\u043e\u0432\u044b\u0439 \u043f\u043e\u0434\u0431\u043e\u0440",
    "users": "\u041f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u0442\u0435\u043b\u0438",
    "requests": "\u0417\u0430\u043f\u0440\u043e\u0441\u044b \u0441\u043f\u0435\u0446\u0438\u0430\u043b\u0438\u0441\u0442\u043e\u0432",
    "my_requests": "\u041c\u043e\u0438 \u0437\u0430\u043f\u0440\u043e\u0441\u044b",
    "audit": "\u0416\u0443\u0440\u043d\u0430\u043b \u0434\u0435\u0439\u0441\u0442\u0432\u0438\u0439",
    "restore_original": "\u0412\u0435\u0440\u043d\u0443\u0442\u044c \u0438\u0441\u0445\u043e\u0434\u043d\u043e\u0435 \u0437\u043d\u0430\u0447\u0435\u043d\u0438\u0435",
    "logout": "\u0412\u044b\u0439\u0442\u0438",
    "download": "\u0421\u043a\u0430\u0447\u0430\u0442\u044c Excel",
    "search": "\u041f\u043e\u0438\u0441\u043a \u043f\u043e \u0442\u0430\u0431\u043b\u0438\u0446\u0435",
    "columns": "\u041a\u043e\u043b\u043e\u043d\u043a\u0438",
    "changed_only": "\u0422\u043e\u043b\u044c\u043a\u043e \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u043d\u044b\u0435",
    "flagged_only": "\u0422\u043e\u043b\u044c\u043a\u043e \u0441 \u0437\u0430\u043f\u0440\u043e\u0441\u0430\u043c\u0438",
    "risk_only": "\u0422\u043e\u043b\u044c\u043a\u043e \u0440\u0438\u0441\u043a\u0438",
    "with_analog_only": "\u0422\u043e\u043b\u044c\u043a\u043e \u0441 \u0430\u043d\u0430\u043b\u043e\u0433\u0430\u043c\u0438",
    "no_analog_only": "\u0411\u0435\u0437 \u0430\u043d\u0430\u043b\u043e\u0433\u0430",
    "saved": "\u0412\u0441\u0435 \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u044f \u0441\u043e\u0445\u0440\u0430\u043d\u0435\u043d\u044b",
    "saving": "\u0421\u043e\u0445\u0440\u0430\u043d\u0435\u043d\u0438\u0435...",
    "offline": "\u041d\u0435\u0442 \u0441\u043e\u0435\u0434\u0438\u043d\u0435\u043d\u0438\u044f. \u0418\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u044f \u043e\u0441\u0442\u0430\u043b\u0438\u0441\u044c \u0432 \u043e\u0447\u0435\u0440\u0435\u0434\u0438.",
    "request_change": "\u0417\u0430\u043f\u0440\u043e\u0441\u0438\u0442\u044c \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u0435",
    "request_blue": "\u0417\u0430\u043f\u0440\u043e\u0441\u0438\u0442\u044c \u0441\u0438\u043d\u044e\u044e \u0437\u0430\u0434\u0430\u0447\u0443",
    "comment": "\u041a\u043e\u043c\u043c\u0435\u043d\u0442\u0430\u0440\u0438\u0439",
    "send": "\u041e\u0442\u043f\u0440\u0430\u0432\u0438\u0442\u044c",
    "cancel": "\u041e\u0442\u043c\u0435\u043d\u0430",
    "approved": "\u041e\u0434\u043e\u0431\u0440\u0435\u043d\u043e",
    "rejected": "\u041e\u0442\u043a\u043b\u043e\u043d\u0435\u043d\u043e",
    "pending": "\u041d\u0430 \u0441\u043e\u0433\u043b\u0430\u0441\u043e\u0432\u0430\u043d\u0438\u0438",
    "admin": "\u0410\u0434\u043c\u0438\u043d\u043a\u0430",
    "file": "\u0424\u0430\u0439\u043b",
    "owner": "\u0421\u043f\u0435\u0446\u0438\u0430\u043b\u0438\u0441\u0442",
    "rows": "\u0421\u0442\u0440\u043e\u043a\u0438",
    "matched": "\u041f\u043e\u0434\u043e\u0431\u0440\u0430\u043d\u043e",
    "risks": "\u0420\u0438\u0441\u043a\u0438",
    "status": "\u0421\u0442\u0430\u0442\u0443\u0441",
    "open": "\u041e\u0442\u043a\u0440\u044b\u0442\u044c",
    "help_edit": "\u0414\u0432\u043e\u0439\u043d\u043e\u0439 \u043a\u043b\u0438\u043a \u2014 \u0438\u0437\u043c\u0435\u043d\u0438\u0442\u044c. \u041f\u0440\u0430\u0432\u0430\u044f \u043a\u043d\u043e\u043f\u043a\u0430 \u2014 \u0437\u0430\u043f\u0440\u043e\u0441 \u0438\u043b\u0438 \u0432\u043e\u0437\u0432\u0440\u0430\u0442.",
    "restore_confirm": "\u0412\u0435\u0440\u043d\u0443\u0442\u044c \u044d\u0442\u0443 \u044f\u0447\u0435\u0439\u043a\u0443 \u043a \u0438\u0441\u0445\u043e\u0434\u043d\u043e\u043c\u0443 \u0437\u043d\u0430\u0447\u0435\u043d\u0438\u044e?",
    "approve_confirm": "\u041e\u0434\u043e\u0431\u0440\u0438\u0442\u044c \u044d\u0442\u043e\u0442 \u0437\u0430\u043f\u0440\u043e\u0441 \u0438 \u043f\u0440\u0438\u043c\u0435\u043d\u0438\u0442\u044c \u0441\u0438\u0441\u0442\u0435\u043c\u043d\u043e\u0435 \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u0435?",
    "deactivate_confirm": "\u0417\u0430\u0431\u043b\u043e\u043a\u0438\u0440\u043e\u0432\u0430\u0442\u044c \u044d\u0442\u043e\u0433\u043e \u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u0442\u0435\u043b\u044f?",
}


def install_workspace_routes(app: FastAPI) -> None:
    if getattr(app.state, "workspace_routes_installed", False):
        return
    app.state.workspace_routes_installed = True

    @app.middleware("http")
    async def workspace_auth_guard(request: Request, call_next):
        path = request.url.path
        connection = connect(default_database_path())
        try:
            init_database(connection)
            has_users = count_users(connection) > 0
            request.state.user = None
            if not has_users:
                return await call_next(request)

            if path in {"/login", "/setup", "/favicon.ico"}:
                return await call_next(request)

            token = request.cookies.get(SESSION_COOKIE, "")
            user = user_from_session(connection, token)
            if user is None:
                if path.startswith("/api/"):
                    return JSONResponse({"ok": False, "error": "authentication_required"}, status_code=401)
                next_path = _request_path(request)
                return RedirectResponse(
                    f"/login?next={quote(next_path)}",
                    status_code=303,
                )

            request.state.user = user
            if path.startswith("/admin") and user.role != ROLE_ADMIN:
                return HTMLResponse(
                    _render_message(
                        "\u041d\u0435\u0442 \u0434\u043e\u0441\u0442\u0443\u043f\u0430",
                        "\u042d\u0442\u043e\u0442 \u0440\u0430\u0437\u0434\u0435\u043b \u0434\u043e\u0441\u0442\u0443\u043f\u0435\u043d \u0442\u043e\u043b\u044c\u043a\u043e \u0430\u0434\u043c\u0438\u043d\u0438\u0441\u0442\u0440\u0430\u0442\u043e\u0440\u0443.",
                        user,
                    ),
                    status_code=403,
                )
            return await call_next(request)
        finally:
            connection.close()

    @app.get("/setup", response_class=HTMLResponse)
    def setup_page() -> HTMLResponse:
        connection = connect(default_database_path())
        try:
            init_database(connection)
            if count_users(connection) > 0:
                return RedirectResponse("/login", status_code=303)
        finally:
            connection.close()
        return HTMLResponse(_render_setup())

    @app.post("/setup")
    def setup_submit(
        request: Request,
        full_name: str = Form(...),
        login: str = Form(...),
        password: str = Form(...),
    ):
        connection = connect(default_database_path())
        try:
            init_database(connection)
            if count_users(connection) > 0:
                return RedirectResponse("/login", status_code=303)
            try:
                user_id = create_user(
                    connection,
                    full_name=full_name,
                    login=login,
                    password=password,
                    role=ROLE_ADMIN,
                )
                token = create_session(connection, user_id, lifetime_days=SESSION_DAYS)
            except ValueError as error:
                return HTMLResponse(_render_setup(str(error)), status_code=400)
        finally:
            connection.close()
        response = RedirectResponse("/admin/users", status_code=303)
        _set_session_cookie(response, token, secure=_request_is_secure(request))
        return response

    @app.get("/login", response_class=HTMLResponse)
    def login_page(next: str = "") -> HTMLResponse:
        return HTMLResponse(_render_login(_safe_next(next)))

    @app.post("/login")
    def login_submit(
        request: Request,
        login: str = Form(...),
        password: str = Form(...),
        next: str = Form(""),
    ):
        connection = connect(default_database_path())
        try:
            init_database(connection)
            user = authenticate_user(connection, login, password)
            if user is None:
                return HTMLResponse(
                    _render_login(_safe_next(next), TEXT["bad_login"]),
                    status_code=401,
                )
            token = create_session(connection, user.id, lifetime_days=SESSION_DAYS)
        finally:
            connection.close()
        destination = _safe_next(next) or ("/admin" if user.role == ROLE_ADMIN else "/jobs")
        response = RedirectResponse(destination, status_code=303)
        _set_session_cookie(response, token, secure=_request_is_secure(request))
        return response

    @app.post("/logout")
    def logout(request: Request):
        token = request.cookies.get(SESSION_COOKIE, "")
        connection = connect(default_database_path())
        try:
            init_database(connection)
            revoke_session(connection, token)
        finally:
            connection.close()
        response = RedirectResponse("/login", status_code=303)
        response.delete_cookie(SESSION_COOKIE)
        return response

    @app.get("/jobs", response_class=HTMLResponse)
    def jobs_page(request: Request) -> HTMLResponse:
        user = request.state.user
        connection = connect(default_database_path())
        try:
            init_database(connection)
            jobs = list_processing_jobs(
                connection,
                owner_user_id=None if user.role == ROLE_ADMIN else user.id,
            )
        finally:
            connection.close()
        return HTMLResponse(_render_jobs(jobs, user))

    @app.get("/requests", response_class=HTMLResponse)
    def my_requests_page(request: Request) -> HTMLResponse:
        user = request.state.user
        connection = connect(default_database_path())
        try:
            init_database(connection)
            requests = list_change_requests(
                connection,
                submitted_by=user.id,
                limit=5000,
            )
        finally:
            connection.close()
        return HTMLResponse(_render_my_requests(requests, user))

    @app.get("/jobs/{job_id}", response_class=HTMLResponse)
    def workspace_page(request: Request, job_id: str) -> HTMLResponse:
        user = request.state.user
        connection = connect(default_database_path())
        try:
            init_database(connection)
            job = _accessible_job(connection, job_id, user)
            if job is None:
                return HTMLResponse(
                    _render_message(
                        "\u041e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0430 \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d\u0430",
                        "\u0424\u0430\u0439\u043b \u043d\u0435\u0434\u043e\u0441\u0442\u0443\u043f\u0435\u043d \u0438\u043b\u0438 \u043d\u0435 \u043f\u0440\u0438\u043d\u0430\u0434\u043b\u0435\u0436\u0438\u0442 \u0442\u0435\u043a\u0443\u0449\u0435\u043c\u0443 \u0441\u043f\u0435\u0446\u0438\u0430\u043b\u0438\u0441\u0442\u0443.",
                        user,
                    ),
                    status_code=404,
                )
            touch_job(connection, job.id)
        finally:
            connection.close()
        return HTMLResponse(_render_workspace(job, user))

    @app.get("/api/jobs/{job_id}/grid")
    def workspace_grid(request: Request, job_id: str):
        user = request.state.user
        connection = connect(default_database_path())
        try:
            init_database(connection)
            job = _accessible_job(connection, job_id, user)
            if job is None:
                return JSONResponse({"ok": False, "error": "not_found"}, status_code=404)
            rows = list_processing_rows(connection, job.id)
            preview_values = resolve_preview_row_values(job, rows)
            overrides = list_cell_overrides(connection, job.id)
            all_requests = list_change_requests(
                connection,
                submitted_by=None if user.role == ROLE_ADMIN else user.id,
                limit=5000,
            )
            requests = [item for item in all_requests if item.job_id == job.id]
            preference = get_user_preference(
                connection,
                user.id,
                key=GRID_PREFERENCE_KEY,
                default={},
            )
        finally:
            connection.close()

        return JSONResponse(
            {
                "ok": True,
                "job": {
                    "id": job.id,
                    "filename": job.estimate_filename,
                    "status": job.status,
                    "owner_name": job.owner_name,
                    "total_rows": job.total_rows,
                    "matched_rows": job.matched_rows,
                    "flagged_rows": job.flagged_rows,
                    "tkp_matched_rows": job.tkp_matched_rows,
                },
                "columns": job.column_schema,
                "rows": [
                    {
                        "id": row.id,
                        "row_index": row.row_index,
                        "excel_row_number": row.excel_row_number,
                        "values": preview_values.get(row.id, row.values),
                        "metadata": row.metadata,
                    }
                    for row in rows
                ],
                "overrides": [
                    {
                        "id": item.id,
                        "row_id": item.row_id,
                        "column_index": item.column_index,
                        "original_value": item.original_value,
                        "current_value": item.current_value,
                        "editor_name": item.editor_name,
                        "revision": item.revision,
                        "updated_at": item.updated_at,
                    }
                    for item in overrides
                ],
                "requests": [
                    {
                        "id": item.id,
                        "row_id": item.row_id,
                        "column_index": item.column_index,
                        "request_type": item.request_type,
                        "task_number": item.task_number,
                        "comment": item.comment,
                        "status": item.status,
                        "submitted_at": item.submitted_at,
                        "review_comment": item.review_comment,
                    }
                    for item in requests
                ],
                "preference": preference,
            }
        )

    @app.post("/api/jobs/{job_id}/edit")
    async def workspace_edit(request: Request, job_id: str):
        user = request.state.user
        payload = await request.json()
        connection = connect(default_database_path())
        try:
            init_database(connection)
            job = _accessible_job(connection, job_id, user)
            if job is None:
                return JSONResponse({"ok": False, "error": "not_found"}, status_code=404)
            try:
                row_id = int(payload.get("row_id"))
                column_index = int(payload.get("column_index"))
                client_change_id = str(payload.get("client_change_id") or "")
                new_value = _scalar_value(payload.get("new_value"))
                saved = save_cell_override(
                    connection,
                    job_id=job.id,
                    row_id=row_id,
                    column_index=column_index,
                    new_value=new_value,
                    editor_user_id=user.id,
                    client_change_id=client_change_id,
                )
            except (TypeError, ValueError) as error:
                return JSONResponse(
                    {"ok": False, "error": str(error)},
                    status_code=400,
                )
        finally:
            connection.close()

        return JSONResponse(
            {
                "ok": True,
                "override": {
                    "id": saved.id,
                    "row_id": saved.row_id,
                    "column_index": saved.column_index,
                    "original_value": saved.original_value,
                    "current_value": saved.current_value,
                    "editor_name": saved.editor_name,
                    "revision": saved.revision,
                    "updated_at": saved.updated_at,
                },
            }
        )

    @app.post("/api/jobs/{job_id}/requests")
    async def workspace_request_change(request: Request, job_id: str):
        user = request.state.user
        payload = await request.json()
        connection = connect(default_database_path())
        try:
            init_database(connection)
            job = _accessible_job(connection, job_id, user)
            if job is None:
                return JSONResponse({"ok": False, "error": "not_found"}, status_code=404)
            try:
                request_id = create_change_request(
                    connection,
                    job_id=job.id,
                    row_id=_optional_int(payload.get("row_id")),
                    column_index=_optional_int(payload.get("column_index")),
                    request_type=str(payload.get("request_type") or REQUEST_OTHER),
                    task_number=str(payload.get("task_number") or ""),
                    comment=str(payload.get("comment") or ""),
                    submitted_by=user.id,
                )
            except (TypeError, ValueError) as error:
                return JSONResponse(
                    {"ok": False, "error": str(error)},
                    status_code=400,
                )
        finally:
            connection.close()
        return JSONResponse({"ok": True, "request_id": request_id})

    @app.post("/api/preferences/grid")
    async def workspace_save_preference(request: Request):
        user = request.state.user
        payload = await request.json()
        connection = connect(default_database_path())
        try:
            init_database(connection)
            save_user_preference(
                connection,
                user.id,
                key=GRID_PREFERENCE_KEY,
                value=payload,
            )
        finally:
            connection.close()
        return JSONResponse({"ok": True})

    @app.get("/jobs/{job_id}/download")
    def workspace_download(request: Request, job_id: str):
        user = request.state.user
        connection = connect(default_database_path())
        try:
            init_database(connection)
            job = _accessible_job(connection, job_id, user)
            if job is None:
                return HTMLResponse(
                    _render_message(
                        "\u0424\u0430\u0439\u043b \u043d\u0435 \u043d\u0430\u0439\u0434\u0435\u043d",
                        "\u041e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0430 \u043d\u0435\u0434\u043e\u0441\u0442\u0443\u043f\u043d\u0430.",
                        user,
                    ),
                    status_code=404,
                )
            try:
                final_path = build_postprocessed_workbook(connection, job.id)
                record_activity_event(
                    connection,
                    actor_user_id=user.id,
                    event_type="excel_downloaded",
                    entity_type="processing_job",
                    entity_id=job.id,
                    job_id=job.id,
                    details=Path(final_path).name,
                )
            except (FileNotFoundError, ValueError) as error:
                return HTMLResponse(
                    _render_message(
                        "\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u0441\u043e\u0431\u0440\u0430\u0442\u044c Excel",
                        str(error),
                        user,
                    ),
                    status_code=500,
                )
        finally:
            connection.close()
        media_type = (
            "application/vnd.ms-excel.sheet.macroEnabled.12"
            if Path(final_path).suffix.lower() == ".xlsm"
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        return FileResponse(
            final_path,
            filename=Path(final_path).name,
            media_type=media_type,
        )

    @app.get("/admin/audit", response_class=HTMLResponse)
    def admin_audit_page(request: Request):
        connection = connect(default_database_path())
        try:
            init_database(connection)
            events = list_activity_events(connection, limit=5000)
        finally:
            connection.close()
        return HTMLResponse(_render_audit(events, request.state.user))

    @app.get("/admin/users", response_class=HTMLResponse)
    def admin_users_page(request: Request, message: str = "", error: str = ""):
        connection = connect(default_database_path())
        try:
            init_database(connection)
            users = list_users(connection)
        finally:
            connection.close()
        return HTMLResponse(_render_users(users, request.state.user, message, error))

    @app.post("/admin/users/create")
    def admin_users_create(
        request: Request,
        full_name: str = Form(...),
        login: str = Form(...),
        password: str = Form(...),
        role: str = Form(ROLE_SPECIALIST),
    ):
        connection = connect(default_database_path())
        try:
            init_database(connection)
            try:
                create_user(
                    connection,
                    full_name=full_name,
                    login=login,
                    password=password,
                    role=role,
                )
            except ValueError as error:
                return RedirectResponse(
                    f"/admin/users?error={quote(str(error))}",
                    status_code=303,
                )
        finally:
            connection.close()
        return RedirectResponse(
            f"/admin/users?message={quote('User created')}",
            status_code=303,
        )

    @app.post("/admin/users/toggle")
    def admin_users_toggle(
        request: Request,
        user_id: int = Form(...),
        is_active: str = Form("0"),
    ):
        current = request.state.user
        if int(user_id) == int(current.id) and is_active != "1":
            return RedirectResponse(
                f"/admin/users?error={quote('Cannot deactivate current admin')}",
                status_code=303,
            )
        connection = connect(default_database_path())
        try:
            init_database(connection)
            set_user_active(connection, user_id, is_active=is_active == "1")
        finally:
            connection.close()
        return RedirectResponse("/admin/users", status_code=303)

    @app.post("/admin/users/password")
    def admin_users_password(
        user_id: int = Form(...),
        password: str = Form(...),
    ):
        connection = connect(default_database_path())
        try:
            init_database(connection)
            try:
                changed = reset_user_password(
                    connection,
                    user_id,
                    password=password,
                )
            except ValueError as error:
                return RedirectResponse(
                    f"/admin/users?error={quote(str(error))}",
                    status_code=303,
                )
        finally:
            connection.close()
        if not changed:
            return RedirectResponse(
                f"/admin/users?error={quote('User not found')}",
                status_code=303,
            )
        return RedirectResponse(
            f"/admin/users?message={quote('Password changed')}",
            status_code=303,
        )

    @app.get("/admin/change-requests", response_class=HTMLResponse)
    def admin_requests_page(
        request: Request,
        status: str = REQUEST_PENDING,
        message: str = "",
        error: str = "",
    ):
        connection = connect(default_database_path())
        try:
            init_database(connection)
            requests = list_change_requests(
                connection,
                status="" if status == "all" else status,
                limit=5000,
            )
        finally:
            connection.close()
        return HTMLResponse(
            _render_requests(
                requests,
                request.state.user,
                status,
                message,
                error,
            )
        )

    @app.post("/admin/change-requests/approve")
    def admin_request_approve(
        request: Request,
        request_id: int = Form(...),
        review_comment: str = Form(""),
    ):
        user = request.state.user
        connection = connect(default_database_path())
        try:
            init_database(connection)
            item = get_change_request(connection, request_id)
            if item is None or item.status != REQUEST_PENDING:
                return RedirectResponse(
                    f"/admin/change-requests?error={quote('Pending request not found')}",
                    status_code=303,
                )
            try:
                if item.request_type == REQUEST_BLUE_TASK:
                    if not item.task_number.strip():
                        raise ValueError("Task number is required for blue task")
                    upsert_task_color_entry(
                        connection,
                        task_number=item.task_number,
                        reason="TKP_PLUS1",
                        comment=item.comment,
                        enabled=True,
                    )
                changed = review_change_request(
                    connection,
                    request_id,
                    status=REQUEST_APPROVED,
                    reviewed_by=user.id,
                    review_comment=review_comment,
                )
                if not changed:
                    raise ValueError("Request was already reviewed")
            except ValueError as error:
                return RedirectResponse(
                    f"/admin/change-requests?error={quote(str(error))}",
                    status_code=303,
                )
        finally:
            connection.close()
        return RedirectResponse(
            f"/admin/change-requests?message={quote('Request approved')}",
            status_code=303,
        )

    @app.post("/admin/change-requests/reject")
    def admin_request_reject(
        request: Request,
        request_id: int = Form(...),
        review_comment: str = Form(...),
    ):
        comment = review_comment.strip()
        if not comment:
            return RedirectResponse(
                f"/admin/change-requests?error={quote('Review comment is required')}",
                status_code=303,
            )
        connection = connect(default_database_path())
        try:
            init_database(connection)
            changed = review_change_request(
                connection,
                request_id,
                status=REQUEST_REJECTED,
                reviewed_by=request.state.user.id,
                review_comment=comment,
            )
        finally:
            connection.close()
        if not changed:
            return RedirectResponse(
                f"/admin/change-requests?error={quote('Pending request not found')}",
                status_code=303,
            )
        return RedirectResponse(
            f"/admin/change-requests?message={quote('Request rejected')}",
            status_code=303,
        )


def _accessible_job(connection, job_id: str, user):
    job = get_processing_job(connection, str(job_id))
    if job is None:
        return None
    if user.role == ROLE_ADMIN or int(job.owner_user_id) == int(user.id):
        return job
    return None


def _request_is_secure(request: Request) -> bool:
    forwarded = str(request.headers.get("x-forwarded-proto") or "").split(",", 1)[0].strip()
    return request.url.scheme == "https" or forwarded == "https"


def _set_session_cookie(response, token: str, *, secure: bool = False) -> None:
    response.set_cookie(
        SESSION_COOKIE,
        token,
        max_age=SESSION_DAYS * 24 * 60 * 60,
        httponly=True,
        secure=secure,
        samesite="strict",
        path="/",
    )


def _request_path(request: Request) -> str:
    query = str(request.url.query)
    return str(request.url.path) + (f"?{query}" if query else "")


def _safe_next(value: str) -> str:
    text = str(value or "").strip()
    if not text.startswith("/") or text.startswith("//"):
        return ""
    return text


def _optional_int(value: object) -> int | None:
    if value in (None, ""):
        return None
    return int(value)


def _scalar_value(value: object) -> object:
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    raise ValueError("Cell value must be a scalar")


def _escape(value: object) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def _nav(user) -> str:
    if user is None:
        return ""
    admin_links = ""
    if user.role == ROLE_ADMIN:
        admin_links = (
            f'<a href="/admin">{_escape(TEXT["admin"])}</a>'
            f'<a href="/admin/users">{_escape(TEXT["users"])}</a>'
            f'<a href="/admin/change-requests">{_escape(TEXT["requests"])}</a>'
            f'<a href="/admin/audit">{_escape(TEXT["audit"])}</a>'
        )
    else:
        admin_links = f'<a href="/requests">{_escape(TEXT["my_requests"])}</a>'
    return (
        '<nav class="top-nav">'
        '<a class="brand" href="/">Estimate AI</a>'
        f'<a href="/jobs">{_escape(TEXT["all_jobs"] if user.role == ROLE_ADMIN else TEXT["jobs"])}</a>'
        f'<a href="/">{_escape(TEXT["new_run"])}</a>'
        f'{admin_links}'
        '<span class="nav-spacer"></span>'
        f'<span class="user-chip">{_escape(user.full_name)}</span>'
        '<form action="/logout" method="post">'
        f'<button class="ghost" type="submit">{_escape(TEXT["logout"])}</button>'
        '</form>'
        '</nav>'
    )


def _page(title: str, body: str, user=None, *, script: str = "") -> str:
    return f"""<!doctype html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{_escape(title)} - Estimate AI</title>
<style>
:root{{--bg:#f5f7fb;--panel:#fff;--line:#dfe4ec;--text:#1e293b;--muted:#64748b;--accent:#2563eb;--danger:#b42318;--ok:#157f3c;--warn:#9a6700;--shadow:0 10px 30px rgba(15,23,42,.08)}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--bg);color:var(--text);font:14px/1.45 Inter,Segoe UI,Arial,sans-serif}}
a{{color:var(--accent);text-decoration:none}}
button,input,select,textarea{{font:inherit}}
.top-nav{{height:58px;display:flex;align-items:center;gap:18px;padding:0 24px;background:#fff;border-bottom:1px solid var(--line);position:sticky;top:0;z-index:50}}
.top-nav .brand{{font-weight:800;color:#111827;margin-right:8px}}
.nav-spacer{{flex:1}}
.user-chip{{font-weight:600;color:#334155}}
button,.button{{border:0;border-radius:9px;background:var(--accent);color:#fff;padding:9px 14px;font-weight:650;cursor:pointer}}
button.ghost,.button.ghost{{background:#eef2f7;color:#334155}}
button.danger{{background:#fee4e2;color:#b42318}}
button.success{{background:#e7f6ec;color:#157f3c}}
button:disabled{{opacity:.5;cursor:not-allowed}}
main{{padding:24px;max-width:1600px;margin:0 auto}}
.card{{background:var(--panel);border:1px solid var(--line);border-radius:14px;box-shadow:var(--shadow);padding:18px}}
.auth-wrap{{min-height:100vh;display:grid;place-items:center;padding:24px}}
.auth-card{{width:min(440px,100%);background:#fff;border:1px solid var(--line);border-radius:18px;padding:28px;box-shadow:var(--shadow)}}
.auth-card h1{{margin:0 0 8px}}
.field{{display:grid;gap:6px;margin:14px 0}}
.field label{{font-weight:650}}
.field input,.field select,.field textarea{{width:100%;border:1px solid #cbd5e1;border-radius:9px;padding:10px 12px;background:#fff}}
.alert{{padding:11px 13px;border-radius:9px;margin:12px 0}}
.alert.error{{background:#fee4e2;color:#8a1c14}}
.alert.ok{{background:#e8f7ed;color:#176b36}}
.page-head{{display:flex;align-items:flex-start;gap:16px;margin-bottom:18px}}
.page-head h1{{margin:0;font-size:24px}}
.page-head p{{margin:5px 0 0;color:var(--muted)}}
.page-head .actions{{margin-left:auto;display:flex;gap:8px}}
.metrics{{display:grid;grid-template-columns:repeat(4,minmax(140px,1fr));gap:10px;margin:0 0 14px}}
.metric{{background:#fff;border:1px solid var(--line);border-radius:12px;padding:12px 14px}}
.metric strong{{display:block;font-size:20px}}
.metric span{{color:var(--muted);font-size:12px}}
.simple-table{{width:100%;border-collapse:collapse;background:#fff}}
.simple-table th,.simple-table td{{padding:10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}}
.simple-table th{{background:#f8fafc;font-size:12px;color:#475569}}
.inline{{display:flex;gap:8px;align-items:center;flex-wrap:wrap}}
.inline input,.inline select{{border:1px solid #cbd5e1;border-radius:8px;padding:8px}}
.toolbar{{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px}}
.toolbar input[type=search]{{min-width:280px;border:1px solid #cbd5e1;border-radius:9px;padding:9px 11px}}
.save-state{{margin-left:auto;font-weight:650;color:var(--ok)}}
.save-state.offline{{color:var(--warn)}}
.grid-shell{{position:relative;background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden}}
.grid-hscroll{{height:18px;overflow-x:auto;overflow-y:hidden;background:#f8fafc;border-bottom:1px solid var(--line)}}
.grid-hscroll-inner{{height:1px}}
.grid-scroll{{height:calc(100vh - 308px);min-height:420px;overflow-y:auto;overflow-x:hidden;position:relative}}
.review-grid{{border-collapse:separate;border-spacing:0;table-layout:fixed;min-width:100%;width:max-content}}
.review-grid th,.review-grid td{{border-right:1px solid #e5e7eb;border-bottom:1px solid #e5e7eb;background:#fff}}
.review-grid td{{padding:6px 8px;vertical-align:top;white-space:normal;overflow:hidden;text-overflow:clip}}
.review-grid td .cell-text{{white-space:normal;overflow-wrap:anywhere;word-break:break-word;line-height:18px;overflow:hidden}}
.review-grid th{{position:sticky;top:0;height:104px;z-index:12;background:#f8fafc;vertical-align:top;padding:6px}}
.review-grid th .head-stack{{height:92px;display:grid;grid-template-rows:46px 14px 26px;gap:3px;min-width:0;align-content:start}}
.review-grid th .head-label{{font-size:12px;line-height:14px;font-weight:700;cursor:pointer;display:flex;align-items:flex-start;justify-content:center;text-align:center;white-space:normal;overflow:hidden;overflow-wrap:anywhere;word-break:break-word;min-width:0}}
.review-grid th .head-sub{{font-size:10px;line-height:14px;color:#64748b;display:block;text-align:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:0}}
.review-grid th input{{width:100%;height:26px;min-width:0;margin:0;border:1px solid #d6dce5;border-radius:5px;padding:3px 5px;font-size:11px;align-self:end}}
.review-grid td.changed{{background:#fff7d6}}
.review-grid td.pending{{outline:2px solid #f59e0b;outline-offset:-2px}}
.review-grid td.requested{{box-shadow:inset 0 -3px #7c3aed}}
.review-grid td.selected{{outline:2px solid var(--accent);outline-offset:-2px}}
.review-grid td input{{width:100%;height:30px;border:1px solid var(--accent);border-radius:5px;padding:3px 5px}}
.sticky-cell{{position:sticky!important;z-index:8!important;background:#fff!important}}
.review-grid th.sticky-cell{{z-index:20!important;background:#f8fafc!important}}
.column-panel{{position:absolute;right:12px;top:48px;width:320px;max-height:430px;overflow:auto;background:#fff;border:1px solid var(--line);border-radius:10px;padding:10px;box-shadow:var(--shadow);z-index:40}}
.column-panel .column-row{{display:flex;gap:6px;align-items:center;padding:4px}}
.column-panel .column-row label{{display:flex;gap:8px;align-items:center;flex:1;padding:2px}}
.column-panel .column-row button{{padding:3px 7px;border-radius:6px;background:#eef2f7;color:#334155}}
.context-menu{{position:fixed;display:none;background:#fff;border:1px solid var(--line);border-radius:9px;box-shadow:var(--shadow);padding:6px;z-index:100}}
.context-menu button{{display:block;width:100%;text-align:left;background:#fff;color:#334155;padding:8px 10px}}
.context-menu button:hover{{background:#f1f5f9}}
.modal-backdrop{{position:fixed;inset:0;background:rgba(15,23,42,.35);display:none;align-items:center;justify-content:center;z-index:120}}
.modal{{width:min(560px,calc(100vw - 40px));background:#fff;border-radius:14px;padding:20px;box-shadow:var(--shadow)}}
.modal textarea{{width:100%;min-height:120px;border:1px solid #cbd5e1;border-radius:9px;padding:10px}}
.muted{{color:var(--muted)}}
.status{{font-weight:650}}
.status.pending{{color:#9a6700}} .status.approved{{color:#157f3c}} .status.rejected{{color:#b42318}}
@media(max-width:900px){{main{{padding:14px}}.metrics{{grid-template-columns:repeat(2,1fr)}}.top-nav{{padding:0 12px;gap:10px}}}}
</style>
</head>
<body>
{_nav(user)}
{body}
{script}
</body>
</html>"""


def _render_login(next_path: str, error: str = "") -> str:
    error_html = f'<div class="alert error">{_escape(error)}</div>' if error else ""
    body = f"""<div class="auth-wrap"><div class="auth-card">
<h1>{_escape(TEXT["login"])}</h1>
<p class="muted">Estimate AI</p>
{error_html}
<form method="post" action="/login">
<input type="hidden" name="next" value="{_escape(next_path)}">
<div class="field"><label>{_escape(TEXT["login_label"])}</label><input name="login" autocomplete="username" required autofocus></div>
<div class="field"><label>{_escape(TEXT["password"])}</label><input type="password" name="password" autocomplete="current-password" required></div>
<button type="submit">{_escape(TEXT["sign_in"])}</button>
</form></div></div>"""
    return _page(TEXT["login"], body)


def _render_setup(error: str = "") -> str:
    error_html = f'<div class="alert error">{_escape(error)}</div>' if error else ""
    body = f"""<div class="auth-wrap"><div class="auth-card">
<h1>{_escape(TEXT["setup"])}</h1>
<p class="muted">Create the first administrator account. This screen closes after setup.</p>
{error_html}
<form method="post" action="/setup">
<div class="field"><label>{_escape(TEXT["full_name"])}</label><input name="full_name" required autofocus></div>
<div class="field"><label>{_escape(TEXT["login_label"])}</label><input name="login" required></div>
<div class="field"><label>{_escape(TEXT["password"])}</label><input type="password" name="password" minlength="8" required></div>
<button type="submit">{_escape(TEXT["create_admin"])}</button>
</form></div></div>"""
    return _page(TEXT["setup"], body)


def _render_jobs(jobs, user) -> str:
    rows = "".join(
        f"""<tr>
<td><a href="/jobs/{_escape(job.id)}"><strong>{_escape(job.estimate_filename)}</strong></a><div class="muted">{_escape(job.updated_at)}</div></td>
<td>{_escape(job.owner_name)}</td>
<td>{job.total_rows}</td><td>{job.matched_rows}</td><td>{job.flagged_rows}</td>
<td><span class="status">{_escape(job.status)}</span></td>
<td><a class="button ghost" href="/jobs/{_escape(job.id)}">{_escape(TEXT["open"])}</a></td>
</tr>"""
        for job in jobs
    )
    title = TEXT["all_jobs"] if user.role == ROLE_ADMIN else TEXT["jobs"]
    body = f"""<main>
<div class="page-head"><div><h1>{_escape(title)}</h1><p>Saved processing workspaces and manual review history.</p></div>
<div class="actions"><a class="button" href="/">{_escape(TEXT["new_run"])}</a></div></div>
<div class="card"><table class="simple-table"><thead><tr>
<th>{_escape(TEXT["file"])}</th><th>{_escape(TEXT["owner"])}</th><th>{_escape(TEXT["rows"])}</th><th>{_escape(TEXT["matched"])}</th><th>{_escape(TEXT["risks"])}</th><th>{_escape(TEXT["status"])}</th><th></th>
</tr></thead><tbody>{rows or '<tr><td colspan="7">No processing jobs yet.</td></tr>'}</tbody></table></div>
</main>"""
    return _page(title, body, user)


def _render_workspace(job, user) -> str:
    labels = json.dumps(TEXT, ensure_ascii=True)
    job_id = json.dumps(job.id)
    body = f"""<main style="max-width:none">
<div class="page-head">
<div><h1>{_escape(job.estimate_filename)}</h1><p>{_escape(job.owner_name)} &middot; {_escape(job.sheet_title)} &middot; {_escape(job.region)}</p></div>
<div class="actions"><a class="button ghost" href="/jobs">{_escape(TEXT["jobs"])}</a><button class="button" id="downloadButton" type="button">{_escape(TEXT["download"])}</button></div>
</div>
<div class="metrics">
<div class="metric"><strong>{job.total_rows}</strong><span>{_escape(TEXT["rows"])}</span></div>
<div class="metric"><strong>{job.matched_rows}</strong><span>{_escape(TEXT["matched"])}</span></div>
<div class="metric"><strong>{job.flagged_rows}</strong><span>{_escape(TEXT["risks"])}</span></div>
<div class="metric"><strong>{job.tkp_matched_rows}</strong><span>TKP</span></div>
</div>
<div class="toolbar">
<input id="globalSearch" type="search" placeholder="{_escape(TEXT["search"])}">
<button class="ghost" id="columnsButton" type="button">{_escape(TEXT["columns"])}</button>
<label><input type="checkbox" id="changedOnly"> {_escape(TEXT["changed_only"])}</label>
<label><input type="checkbox" id="requestedOnly"> {_escape(TEXT["flagged_only"])}</label>
<label><input type="checkbox" id="riskOnly"> {_escape(TEXT["risk_only"])}</label>
<label><input type="checkbox" id="withAnalogOnly"> {_escape(TEXT["with_analog_only"])}</label>
<label><input type="checkbox" id="noAnalogOnly"> {_escape(TEXT["no_analog_only"])}</label>
<span class="muted">{_escape(TEXT["help_edit"])}</span>
<span id="saveState" class="save-state">{_escape(TEXT["saved"])}</span>
</div>
<div class="grid-shell">
<div id="columnPanel" class="column-panel" hidden></div>
<div id="gridHScroll" class="grid-hscroll"><div id="gridHScrollInner" class="grid-hscroll-inner"></div></div>
<div id="gridScroll" class="grid-scroll"><table id="grid" class="review-grid"></table></div>
</div>
<div id="contextMenu" class="context-menu">
<button type="button" data-action="change"></button>
<button type="button" data-action="blue"></button>
<button type="button" data-action="restore"></button>
</div>
<div id="requestModal" class="modal-backdrop"><div class="modal">
<h2 id="requestTitle"></h2>
<p id="requestMeta" class="muted"></p>
<textarea id="requestComment" placeholder="{_escape(TEXT["comment"])}"></textarea>
<div class="inline" style="justify-content:flex-end;margin-top:12px">
<button class="ghost" type="button" id="requestCancel">{_escape(TEXT["cancel"])}</button>
<button type="button" id="requestSend">{_escape(TEXT["send"])}</button>
</div></div></div>
</main>"""
    script = f"""<script>
(() => {{
const JOB_ID = {job_id};
const T = {labels};
const BASE_ROW_HEIGHT = 36;
const BUFFER = 20;
const queueKey = 'estimate-ai-edit-queue:' + JOB_ID;
const localPrefKey = 'estimate-ai-grid-pref-v1:{int(user.id)}';
const state = {{
  columns: [], rows: [], baseRows: [], visibleColumns: [], overrides: new Map(),
  changed: new Set(), requested: new Set(), queue: [], sortColumn: null,
  sortDirection: 1, columnFilters: {{}}, hidden: new Set(), widths: {{}},
  order: [], selected: null, requestType: null
}};
const grid = document.getElementById('grid');
const scroll = document.getElementById('gridScroll');
const saveState = document.getElementById('saveState');
const globalSearch = document.getElementById('globalSearch');
const changedOnly = document.getElementById('changedOnly');
const requestedOnly = document.getElementById('requestedOnly');
const riskOnly = document.getElementById('riskOnly');
const withAnalogOnly = document.getElementById('withAnalogOnly');
const noAnalogOnly = document.getElementById('noAnalogOnly');
const topScroll = document.getElementById('gridHScroll');
const topScrollInner = document.getElementById('gridHScrollInner');
const panel = document.getElementById('columnPanel');
const contextMenu = document.getElementById('contextMenu');
const modal = document.getElementById('requestModal');
const requestTitle = document.getElementById('requestTitle');
const requestMeta = document.getElementById('requestMeta');
const requestComment = document.getElementById('requestComment');

function key(rowId, colIndex) {{ return rowId + ':' + colIndex; }}
function asText(value) {{
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') return new Intl.NumberFormat('ru-RU', {{maximumFractionDigits: 4}}).format(value);
  return String(value);
}}
function rawText(value) {{ return value === null || value === undefined ? '' : String(value); }}
function cellValue(row, column) {{ return row.values[column.index - 1]; }}
function setCellValue(rowId, columnIndex, value) {{
  const row = state.rows.find(item => item.id === rowId);
  if (row) row.values[columnIndex - 1] = value;
}}
function parseInput(raw, previous) {{
  const text = raw.trim();
  if (text === '') return null;
  if (typeof previous === 'number') {{
    const normalized = text.replace(/\s/g, '').replace(',', '.');
    const value = Number(normalized);
    return Number.isFinite(value) ? value : text;
  }}
  return text;
}}
function loadQueue() {{
  try {{ state.queue = JSON.parse(localStorage.getItem(queueKey) || '[]'); }}
  catch (_) {{ state.queue = []; }}
}}
function saveQueue() {{
  localStorage.setItem(queueKey, JSON.stringify(state.queue));
  updateSaveState();
}}
function uuid() {{
  if (window.crypto && window.crypto.randomUUID) return window.crypto.randomUUID();
  return Date.now().toString(36) + '-' + Math.random().toString(36).slice(2);
}}
function updateSaveState() {{
  if (!navigator.onLine) {{
    saveState.textContent = T.offline + (state.queue.length ? ' (' + state.queue.length + ')' : '');
    saveState.classList.add('offline');
    return;
  }}
  saveState.classList.remove('offline');
  saveState.textContent = state.queue.length ? T.saving + ' (' + state.queue.length + ')' : T.saved;
}}
async function loadData() {{
  loadQueue();
  const response = await fetch('/api/jobs/' + JOB_ID + '/grid', {{cache:'no-store'}});
  if (!response.ok) throw new Error('grid_load_failed');
  const data = await response.json();
  state.columns = data.columns || [];
  state.rows = data.rows || [];
  state.baseRows = state.rows;
  (data.overrides || []).forEach(item => {{
    const editKey=key(item.row_id, item.column_index);
    state.overrides.set(editKey, item);
    if (JSON.stringify(item.current_value) !== JSON.stringify(item.original_value)) state.changed.add(editKey);
    setCellValue(item.row_id, item.column_index, item.current_value);
  }});
  (data.requests || []).forEach(item => {{
    if (item.row_id && item.column_index) state.requested.add(key(item.row_id, item.column_index));
  }});
  const pref = (data.preference && typeof data.preference === 'object') ? data.preference : {{}};
  applyPreference(pref);
  state.queue.forEach(item => {{
    setCellValue(item.row_id, item.column_index, item.new_value);
    state.changed.add(key(item.row_id, item.column_index));
  }});
  rebuildColumnPanel();
  applyFilters();
  await flushQueue();
}}
function applyPreference(pref) {{
  const local = (() => {{ try {{ return JSON.parse(localStorage.getItem(localPrefKey) || '{{}}'); }} catch (_) {{ return {{}}; }} }})();
  const merged = Object.assign({{}}, pref || {{}}, local || {{}});
  state.hidden = new Set(Array.isArray(merged.hidden) ? merged.hidden.map(Number) : []);
  state.widths = merged.widths && typeof merged.widths === 'object' ? merged.widths : {{}};
  const known = state.columns.map(column => Number(column.index));
  const savedOrder = Array.isArray(merged.order) ? merged.order.map(Number) : [];
  state.order = savedOrder.filter(index => known.includes(index));
  known.forEach(index => {{ if (!state.order.includes(index)) state.order.push(index); }});
}}
let prefTimer = null;
function persistPreference() {{
  const payload = {{hidden:[...state.hidden], widths:state.widths, order:state.order}};
  localStorage.setItem(localPrefKey, JSON.stringify(payload));
  clearTimeout(prefTimer);
  prefTimer = setTimeout(() => {{
    fetch('/api/preferences/grid', {{
      method:'POST', headers:{{'Content-Type':'application/json'}}, body:JSON.stringify(payload)
    }}).catch(() => {{}});
  }}, 400);
}}
function orderedColumns() {{
  const byIndex = new Map(state.columns.map(column => [Number(column.index), column]));
  const ordered = state.order.map(index => byIndex.get(Number(index))).filter(Boolean);
  state.columns.forEach(column => {{
    if (!ordered.some(item => Number(item.index) === Number(column.index))) ordered.push(column);
  }});
  return ordered;
}}
function moveColumn(columnIndex, delta) {{
  const index = state.order.indexOf(Number(columnIndex));
  const target = index + delta;
  if (index < 0 || target < 0 || target >= state.order.length) return;
  const copy = state.order.slice();
  [copy[index], copy[target]] = [copy[target], copy[index]];
  state.order = copy;
  rebuildColumnPanel(); renderGrid(); persistPreference();
}}
function rebuildColumnPanel() {{
  panel.innerHTML = '';
  orderedColumns().forEach(column => {{
    const row = document.createElement('div'); row.className='column-row';
    const label = document.createElement('label');
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.checked = !state.hidden.has(column.index);
    input.addEventListener('change', () => {{
      if (input.checked) state.hidden.delete(column.index); else state.hidden.add(column.index);
      persistPreference(); renderGrid();
    }});
    const span = document.createElement('span');
    span.textContent = column.label + (column.sublabel ? ' - ' + column.sublabel : '');
    const up=document.createElement('button'); up.type='button'; up.textContent='up';
    const down=document.createElement('button'); down.type='button'; down.textContent='down';
    up.addEventListener('click',()=>moveColumn(column.index,-1));
    down.addEventListener('click',()=>moveColumn(column.index,1));
    label.append(input, span); row.append(label,up,down); panel.appendChild(row);
  }});
}}
function applyFilters(focusColumn = null) {{
  const query = globalSearch.value.trim().toLocaleLowerCase('ru-RU');
  const filters = state.columnFilters;
  let rows = state.baseRows.filter(row => {{
    if (query) {{
      const haystack = row.values.map(asText).join(' ').toLocaleLowerCase('ru-RU');
      if (!haystack.includes(query)) return false;
    }}
    for (const [rawIndex, rawFilter] of Object.entries(filters)) {{
      const filter = rawFilter.trim().toLocaleLowerCase('ru-RU');
      if (!filter) continue;
      const value = asText(row.values[Number(rawIndex) - 1]).toLocaleLowerCase('ru-RU');
      if (!value.includes(filter)) return false;
    }}
    if (changedOnly.checked) {{
      const hasChanged = state.columns.some(col => state.changed.has(key(row.id, col.index)));
      if (!hasChanged) return false;
    }}
    if (requestedOnly.checked) {{
      const hasRequested = state.columns.some(col => state.requested.has(key(row.id, col.index)));
      if (!hasRequested) return false;
    }}
    if (riskOnly.checked && !(row.metadata && row.metadata.risk)) return false;
    const hasAnyAnalog = Boolean(row.metadata && (row.metadata.has_analogs || row.metadata.has_tkp));
    if (withAnalogOnly.checked && !hasAnyAnalog) return false;
    if (noAnalogOnly.checked && hasAnyAnalog) return false;
    return true;
  }});
  if (state.sortColumn) {{
    const index = state.sortColumn.index - 1;
    const dir = state.sortDirection;
    rows = rows.slice().sort((a,b) => {{
      const av = a.values[index], bv = b.values[index];
      if (av === bv) return 0;
      if (av === null || av === undefined) return -1 * dir;
      if (bv === null || bv === undefined) return 1 * dir;
      if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * dir;
      return String(av).localeCompare(String(bv), 'ru', {{numeric:true}}) * dir;
    }});
  }}
  state.rows = rows;
  scroll.scrollTop = 0;
  renderGrid();
  if (focusColumn !== null) {{
    const current = grid.querySelector('input[data-filter-column="' + focusColumn + '"]');
    if (current) {{
      current.focus();
      const end = current.value.length;
      current.setSelectionRange(end, end);
    }}
  }}
}}
function visibleColumns() {{
  return orderedColumns().filter(column => !state.hidden.has(column.index));
}}
function makeColGroup(columns) {{
  const group = document.createElement('colgroup');
  columns.forEach(column => {{
    const col = document.createElement('col');
    col.dataset.column = column.index;
    col.style.width = (Number(state.widths[column.index]) || (column.kind === 'source' ? 170 : 140)) + 'px';
    group.appendChild(col);
  }});
  return group;
}}
function makeHeader(columns) {{
  const thead = document.createElement('thead');
  const tr = document.createElement('tr');
  columns.forEach((column, visibleIndex) => {{
    const th = document.createElement('th');
    const stack = document.createElement('div'); stack.className = 'head-stack';
    const head = document.createElement('span'); head.className = 'head-label'; head.textContent = column.label;
    const sub = document.createElement('span'); sub.className = 'head-sub'; sub.textContent = column.sublabel || column.letter;
    const filter = document.createElement('input'); filter.placeholder = 'filter'; filter.value = state.columnFilters[column.index] || '';
    filter.dataset.filterColumn = column.index;
    filter.addEventListener('click', event => event.stopPropagation());
    filter.addEventListener('input', () => {{ state.columnFilters[column.index] = filter.value; applyFilters(column.index); }});
    head.addEventListener('click', () => {{
      if (state.sortColumn && state.sortColumn.index === column.index) state.sortDirection *= -1;
      else {{ state.sortColumn = column; state.sortDirection = 1; }}
      applyFilters();
    }});
    const resizer = document.createElement('span');
    resizer.style.cssText = 'position:absolute;right:0;top:0;width:6px;height:100%;cursor:col-resize';
    resizer.addEventListener('mousedown', event => startResize(event, column));
    th.style.position = 'sticky';
    stack.append(head, sub, filter);
    th.append(stack, resizer);
    if (visibleIndex < 2) th.classList.add('sticky-cell');
    tr.appendChild(th);
  }});
  thead.appendChild(tr);
  return thead;
}}
function startResize(event, column) {{
  event.preventDefault(); event.stopPropagation();
  const startX = event.clientX;
  const startWidth = Number(state.widths[column.index]) || (column.kind === 'source' ? 170 : 140);
  const move = e => {{
    state.widths[column.index] = Math.max(48, Math.min(520, startWidth + e.clientX - startX));
    renderGrid();
  }};
  const up = () => {{
    document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up);
    persistPreference();
  }};
  document.addEventListener('mousemove', move); document.addEventListener('mouseup', up);
}}
function currentRowHeight(columns) {{
  if (!columns.length) return BASE_ROW_HEIGHT;
  const narrowest = Math.min(...columns.map(column =>
    Number(state.widths[column.index]) || (column.kind === 'source' ? 170 : 140)
  ));
  if (narrowest < 64) return 108;
  if (narrowest < 86) return 90;
  if (narrowest < 116) return 72;
  if (narrowest < 150) return 54;
  return BASE_ROW_HEIGHT;
}}
function renderGrid() {{
  const columns = visibleColumns();
  state.visibleColumns = columns;
  const rowHeight = currentRowHeight(columns);
  const top = scroll.scrollTop;
  const viewportHeight = scroll.clientHeight || 600;
  const start = Math.max(0, Math.floor(top / rowHeight) - BUFFER);
  const count = Math.ceil(viewportHeight / rowHeight) + BUFFER * 2;
  const end = Math.min(state.rows.length, start + count);
  grid.innerHTML = '';
  grid.appendChild(makeColGroup(columns));
  grid.appendChild(makeHeader(columns));
  const tbody = document.createElement('tbody');
  if (start > 0) tbody.appendChild(spacerRow(start * rowHeight, columns.length));
  for (let i=start; i<end; i++) tbody.appendChild(makeRow(state.rows[i], columns, rowHeight));
  if (end < state.rows.length) tbody.appendChild(spacerRow((state.rows.length-end)*rowHeight, columns.length));
  grid.appendChild(tbody);
  applyStickyOffsets(columns);
  syncHorizontalScrollbar();
}}
function syncHorizontalScrollbar() {{
  const contentWidth = Math.max(scroll.clientWidth, grid.scrollWidth);
  topScrollInner.style.width = contentWidth + 'px';
  if (topScroll.scrollLeft !== scroll.scrollLeft) topScroll.scrollLeft = scroll.scrollLeft;
}}
function spacerRow(height, colspan) {{
  const tr=document.createElement('tr'); const td=document.createElement('td');
  td.colSpan=Math.max(1,colspan); td.style.height=height+'px'; td.style.padding='0'; td.style.border='0';
  tr.appendChild(td); return tr;
}}
function makeRow(row, columns, rowHeight) {{
  const tr = document.createElement('tr');
  tr.dataset.rowId = row.id;
  tr.style.height = rowHeight + 'px';
  columns.forEach((column, visibleIndex) => {{
    const td = document.createElement('td');
    const k = key(row.id, column.index);
    const cellText = document.createElement('div');
    cellText.className = 'cell-text';
    cellText.style.maxHeight = Math.max(18, rowHeight - 12) + 'px';
    cellText.textContent = asText(cellValue(row,column));
    td.appendChild(cellText);
    const savedOverride=state.overrides.get(k);
    td.title = savedOverride
      ? ('Original: '+asText(savedOverride.original_value)+' | '+savedOverride.editor_name+' | '+savedOverride.updated_at)
      : rawText(cellValue(row,column));
    td.dataset.rowId = row.id; td.dataset.columnIndex = column.index;
    if (state.changed.has(k)) td.classList.add('changed');
    if (state.queue.some(item => item.row_id===row.id && item.column_index===column.index)) td.classList.add('pending');
    if (state.requested.has(k)) td.classList.add('requested');
    if (visibleIndex < 2) td.classList.add('sticky-cell');
    td.addEventListener('dblclick', () => startEdit(td,row,column));
    td.addEventListener('click', () => selectCell(td,row,column));
    td.addEventListener('contextmenu', event => openContext(event,row,column));
    tr.appendChild(td);
  }});
  return tr;
}}
function applyStickyOffsets(columns) {{
  if (!columns.length) return;
  const firstWidth = Number(state.widths[columns[0].index]) || (columns[0].kind==='source'?170:140);
  grid.querySelectorAll('tr').forEach(tr => {{
    const cells = tr.children;
    if (cells[0] && cells[0].classList.contains('sticky-cell')) cells[0].style.left='0px';
    if (cells[1] && cells[1].classList.contains('sticky-cell')) cells[1].style.left=firstWidth+'px';
  }});
}}
function selectCell(td,row,column) {{
  grid.querySelectorAll('td.selected').forEach(cell=>cell.classList.remove('selected'));
  td.classList.add('selected'); state.selected={{row,column}};
}}
function startEdit(td,row,column) {{
  if (!column.editable || td.querySelector('input')) return;
  const previous = cellValue(row,column);
  const input = document.createElement('input');
  input.value = rawText(previous);
  td.textContent=''; td.appendChild(input); input.focus(); input.select();
  let closed=false;
  const close = save => {{
    if (closed) return; closed=true;
    if (save) {{
      const next = parseInput(input.value, previous);
      if (JSON.stringify(next) !== JSON.stringify(previous)) queueEdit(row,column,next);
    }}
    renderGrid();
  }};
  input.addEventListener('keydown', event => {{
    if (event.key==='Enter') {{ event.preventDefault(); close(true); }}
    if (event.key==='Escape') {{ event.preventDefault(); close(false); }}
  }});
  input.addEventListener('blur', () => close(true));
}}
function queueEdit(row,column,newValue) {{
  const editKey=key(row.id,column.index);
  const pending=state.queue.find(item=>item.row_id===row.id && item.column_index===column.index);
  const saved=state.overrides.get(editKey);
  const originalValue=pending ? pending.original_value : (saved ? saved.original_value : cellValue(row,column));
  setCellValue(row.id,column.index,newValue);
  if (JSON.stringify(newValue) === JSON.stringify(originalValue)) state.changed.delete(editKey);
  else state.changed.add(editKey);
  const item={{client_change_id:uuid(),row_id:row.id,column_index:column.index,new_value:newValue,original_value:originalValue}};
  state.queue.push(item); saveQueue(); renderGrid(); flushQueue();
}}
let flushing=false;
async function flushQueue() {{
  if (flushing || !navigator.onLine || !state.queue.length) {{ updateSaveState(); return; }}
  flushing=true; updateSaveState();
  try {{
    while (navigator.onLine && state.queue.length) {{
      const item=state.queue[0];
      const response=await fetch('/api/jobs/'+JOB_ID+'/edit', {{
        method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify(item)
      }});
      if (response.status===401) {{ location.href='/login'; return; }}
      if (!response.ok) throw new Error('save_failed');
      const data=await response.json();
      const editKey=key(item.row_id,item.column_index);
      state.overrides.set(editKey,data.override);
      if (JSON.stringify(data.override.current_value) === JSON.stringify(data.override.original_value)) state.changed.delete(editKey);
      else state.changed.add(editKey);
      state.queue.shift(); saveQueue();
    }}
  }} catch (_) {{
    updateSaveState();
  }} finally {{
    flushing=false; updateSaveState(); renderGrid();
  }}
}}
function openContext(event,row,column) {{
  event.preventDefault(); selectCell(event.currentTarget,row,column);
  state.selected={{row,column}};
  const change=contextMenu.querySelector('[data-action=change]');
  const blue=contextMenu.querySelector('[data-action=blue]');
  const restore=contextMenu.querySelector('[data-action=restore]');
  change.textContent=T.request_change;
  blue.textContent=T.request_blue;
  restore.textContent=T.restore_original;
  blue.style.display=column.task_number?'block':'none';
  restore.style.display=state.changed.has(key(row.id,column.index))?'block':'none';
  contextMenu.style.left=Math.min(event.clientX,window.innerWidth-330)+'px';
  contextMenu.style.top=Math.min(event.clientY,window.innerHeight-120)+'px';
  contextMenu.style.display='block';
}}
function restoreSelected() {{
  contextMenu.style.display='none';
  if (!state.selected) return;
  if (!window.confirm(T.restore_confirm)) return;
  const {{row,column}}=state.selected;
  const editKey=key(row.id,column.index);
  const saved=state.overrides.get(editKey);
  const pending=state.queue.filter(item=>item.row_id===row.id && item.column_index===column.index);
  const original=saved ? saved.original_value : (pending.length ? pending[0].original_value : undefined);
  if (original === undefined) return;
  state.queue=state.queue.filter(item=>!(item.row_id===row.id && item.column_index===column.index));
  setCellValue(row.id,column.index,original);
  state.changed.delete(editKey);
  saveQueue();
  if (saved && JSON.stringify(saved.current_value)!==JSON.stringify(original)) {{
    queueEdit(row,column,original);
  }} else {{
    renderGrid();
  }}
}}
async function downloadReviewed() {{
  await flushQueue();
  if (state.queue.length) {{
    alert(T.offline);
    return;
  }}
  window.location.href='/jobs/'+JOB_ID+'/download';
}}
function openRequest(type) {{
  contextMenu.style.display='none';
  if (!state.selected) return;
  state.requestType=type; const {{row,column}}=state.selected;
  requestTitle.textContent=type==='blue_task'?T.request_blue:T.request_change;
  requestMeta.textContent='Excel row '+row.excel_row_number+' - '+column.label+(column.task_number?' - '+column.task_number:'');
  requestComment.value=''; modal.style.display='flex'; requestComment.focus();
}}
async function sendRequest() {{
  if (!state.selected) return;
  const comment=requestComment.value.trim();
  if (!comment) {{ requestComment.focus(); return; }}
  const {{row,column}}=state.selected;
  const response=await fetch('/api/jobs/'+JOB_ID+'/requests', {{
    method:'POST',headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{
      row_id:row.id,column_index:column.index,request_type:state.requestType||'other',
      task_number:column.task_number||'',comment
    }})
  }});
  if (!response.ok) return;
  state.requested.add(key(row.id,column.index)); modal.style.display='none'; renderGrid();
}}
contextMenu.querySelector('[data-action=change]').addEventListener('click',()=>openRequest('analog_change'));
contextMenu.querySelector('[data-action=blue]').addEventListener('click',()=>openRequest('blue_task'));
contextMenu.querySelector('[data-action=restore]').addEventListener('click',restoreSelected);
document.getElementById('downloadButton').addEventListener('click',downloadReviewed);
document.getElementById('requestCancel').addEventListener('click',()=>modal.style.display='none');
document.getElementById('requestSend').addEventListener('click',sendRequest);
document.addEventListener('click',event=>{{if(!contextMenu.contains(event.target))contextMenu.style.display='none'}});
document.getElementById('columnsButton').addEventListener('click',()=>{{panel.hidden=!panel.hidden}});
globalSearch.addEventListener('input',()=>applyFilters());
changedOnly.addEventListener('change',()=>applyFilters());
requestedOnly.addEventListener('change',()=>applyFilters());
riskOnly.addEventListener('change',()=>applyFilters());
withAnalogOnly.addEventListener('change',()=>{{
  if (withAnalogOnly.checked) noAnalogOnly.checked = false;
  applyFilters();
}});
noAnalogOnly.addEventListener('change',()=>{{
  if (noAnalogOnly.checked) withAnalogOnly.checked = false;
  applyFilters();
}});
let syncingHorizontal = false;
let lastScrollTop = scroll.scrollTop;
topScroll.addEventListener('scroll',()=>{{
  if (syncingHorizontal) return;
  syncingHorizontal = true;
  scroll.scrollLeft = topScroll.scrollLeft;
  syncingHorizontal = false;
}});
scroll.addEventListener('scroll',()=>{{
  if (!syncingHorizontal && topScroll.scrollLeft !== scroll.scrollLeft) {{
    syncingHorizontal = true;
    topScroll.scrollLeft = scroll.scrollLeft;
    syncingHorizontal = false;
  }}
  if (scroll.scrollTop !== lastScrollTop) {{
    lastScrollTop = scroll.scrollTop;
    requestAnimationFrame(renderGrid);
  }}
}});
window.addEventListener('resize',syncHorizontalScrollbar);
window.addEventListener('online',flushQueue);
window.addEventListener('offline',updateSaveState);
window.addEventListener('beforeunload',event=>{{if(state.queue.length){{event.preventDefault();event.returnValue='';}}}});
setInterval(flushQueue,5000);
loadData().catch(()=>{{saveState.textContent='Grid load failed';saveState.classList.add('offline')}});
}})();
</script>"""
    return _page(job.estimate_filename, body, user, script=script)


def _render_audit(events, user) -> str:
    labels = {
        "processing_created": "\u0421\u043e\u0437\u0434\u0430\u043b \u043e\u0431\u0440\u0430\u0431\u043e\u0442\u043a\u0443",
        "cell_edited": "\u0418\u0437\u043c\u0435\u043d\u0438\u043b \u044f\u0447\u0435\u0439\u043a\u0443",
        "change_request_submitted": "\u041e\u0442\u043f\u0440\u0430\u0432\u0438\u043b \u0437\u0430\u043f\u0440\u043e\u0441",
        "change_request_approved": "\u041e\u0434\u043e\u0431\u0440\u0438\u043b \u0437\u0430\u043f\u0440\u043e\u0441",
        "change_request_rejected": "\u041e\u0442\u043a\u043b\u043e\u043d\u0438\u043b \u0437\u0430\u043f\u0440\u043e\u0441",
        "excel_downloaded": "\u0421\u043a\u0430\u0447\u0430\u043b Excel",
    }
    rows = []
    for event in events:
        details = _activity_details(event.event_type, event.details)
        rows.append(
            f"""<tr>
<td>{_escape(event.created_at)}</td>
<td><strong>{_escape(event.actor_name or "-")}</strong></td>
<td>{_escape(labels.get(event.event_type, event.event_type))}</td>
<td>{_escape(event.estimate_filename or "-")}</td>
<td>{_escape(details)}</td>
</tr>"""
        )
    body = f"""<main>
<div class="page-head"><div><h1>{_escape(TEXT["audit"])}</h1><p>Processing, edits, requests, approvals, and Excel downloads.</p></div></div>
<div class="card"><table class="simple-table"><thead><tr><th>Time</th><th>User</th><th>Action</th><th>File</th><th>Details</th></tr></thead>
<tbody>{''.join(rows) or '<tr><td colspan="5">No activity yet.</td></tr>'}</tbody></table></div>
</main>"""
    return _page(TEXT["audit"], body, user)


def _activity_details(event_type: str, details: str) -> str:
    if event_type != "cell_edited":
        return str(details or "")
    try:
        payload = json.loads(str(details or "{}"))
    except (TypeError, ValueError):
        return str(details or "")
    row = payload.get("excel_row")
    column = payload.get("column")
    old = payload.get("old")
    new = payload.get("new")
    return f"row {row}, column {column}: {old!s} -> {new!s}"


def _render_my_requests(requests, user) -> str:
    rows = []
    for item in requests:
        rows.append(
            f"""<tr>
<td>#{item.id}<div class="muted">{_escape(item.submitted_at)}</div></td>
<td><strong>{_escape(item.estimate_filename)}</strong><div class="muted">row {item.excel_row_number or '-'} / col {item.column_index or '-'}</div></td>
<td>{_escape(item.request_type)}<div class="muted">{_escape(item.task_number)}</div></td>
<td>{_escape(item.comment)}</td>
<td><span class="status {_escape(item.status)}">{_escape(item.status)}</span><div class="muted">{_escape(item.review_comment)}</div></td>
</tr>"""
        )
    body = f"""<main>
<div class="page-head"><div><h1>{_escape(TEXT["my_requests"])}</h1><p>Requests sent from reviewed estimate files.</p></div>
<div class="actions"><a class="button ghost" href="/jobs">{_escape(TEXT["jobs"])}</a></div></div>
<div class="card"><table class="simple-table"><thead><tr><th>ID</th><th>File</th><th>Type</th><th>Comment</th><th>Status</th></tr></thead>
<tbody>{''.join(rows) or '<tr><td colspan="5">No requests yet.</td></tr>'}</tbody></table></div>
</main>"""
    return _page(TEXT["my_requests"], body, user)


def _render_users(users, current_user, message: str, error: str) -> str:
    notice = f'<div class="alert ok">{_escape(message)}</div>' if message else ""
    failure = f'<div class="alert error">{_escape(error)}</div>' if error else ""
    rows = []
    for user in users:
        next_active = "0" if user.is_active else "1"
        toggle_label = "Deactivate" if user.is_active else "Activate"
        toggle_confirm = ""
        if user.is_active:
            toggle_confirm = (
                ' onsubmit="return confirm(\''
                + _escape(TEXT["deactivate_confirm"])
                + '\')"'
            )
        rows.append(
            f"""<tr><td><strong>{_escape(user.full_name)}</strong></td><td>{_escape(user.login)}</td>
<td>{_escape(user.role)}</td><td>{'active' if user.is_active else 'blocked'}</td><td>
<div class="inline">
<form method="post" action="/admin/users/toggle"{toggle_confirm}><input type="hidden" name="user_id" value="{user.id}"><input type="hidden" name="is_active" value="{next_active}"><button class="ghost" type="submit">{toggle_label}</button></form>
<form method="post" action="/admin/users/password"><input type="hidden" name="user_id" value="{user.id}"><input name="password" type="password" minlength="8" placeholder="New password" required><button class="ghost" type="submit">Reset</button></form>
</div></td></tr>"""
        )
    body = f"""<main>
<div class="page-head"><div><h1>{_escape(TEXT["users"])}</h1><p>Create specialist accounts and control access.</p></div></div>
{notice}{failure}
<div class="card" style="margin-bottom:14px"><form class="inline" method="post" action="/admin/users/create">
<input name="full_name" placeholder="{_escape(TEXT["full_name"])}" required>
<input name="login" placeholder="{_escape(TEXT["login_label"])}" required>
<input name="password" type="password" minlength="8" placeholder="{_escape(TEXT["password"])}" required>
<select name="role"><option value="specialist">specialist</option><option value="admin">admin</option></select>
<button type="submit">Create</button></form></div>
<div class="card"><table class="simple-table"><thead><tr><th>Name</th><th>Login</th><th>Role</th><th>Status</th><th>Actions</th></tr></thead><tbody>{''.join(rows)}</tbody></table></div>
</main>"""
    return _page(TEXT["users"], body, current_user)


def _render_requests(requests, user, selected_status: str, message: str, error: str) -> str:
    notice = f'<div class="alert ok">{_escape(message)}</div>' if message else ""
    failure = f'<div class="alert error">{_escape(error)}</div>' if error else ""
    rows = []
    for item in requests:
        actions = ""
        if item.status == REQUEST_PENDING:
            actions = f"""<div class="inline">
<form method="post" action="/admin/change-requests/approve" onsubmit="return confirm('{_escape(TEXT["approve_confirm"])}')"><input type="hidden" name="request_id" value="{item.id}"><input name="review_comment" placeholder="Comment"><button class="success" type="submit">Approve</button></form>
<form method="post" action="/admin/change-requests/reject"><input type="hidden" name="request_id" value="{item.id}"><input name="review_comment" placeholder="Reason" required><button class="danger" type="submit">Reject</button></form>
</div>"""
        rows.append(
            f"""<tr><td>#{item.id}<div class="muted">{_escape(item.submitted_at)}</div></td>
<td><strong>{_escape(item.estimate_filename)}</strong><div class="muted">row {item.excel_row_number or '-'} &middot; col {item.column_index or '-'}</div></td>
<td>{_escape(item.submitted_by_name)}</td><td>{_escape(item.request_type)}<div class="muted">{_escape(item.task_number)}</div></td>
<td>{_escape(item.comment)}</td><td><span class="status {_escape(item.status)}">{_escape(item.status)}</span><div class="muted">{_escape(item.review_comment)}</div></td><td>{actions}</td></tr>"""
        )
    body = f"""<main>
<div class="page-head"><div><h1>{_escape(TEXT["requests"])}</h1><p>Approval queue with specialist, file, row, and comment context.</p></div></div>
{notice}{failure}
<div class="toolbar">
<a class="button {'ghost' if selected_status!='pending' else ''}" href="/admin/change-requests?status=pending">Pending</a>
<a class="button {'ghost' if selected_status!='approved' else ''}" href="/admin/change-requests?status=approved">Approved</a>
<a class="button {'ghost' if selected_status!='rejected' else ''}" href="/admin/change-requests?status=rejected">Rejected</a>
<a class="button {'ghost' if selected_status!='all' else ''}" href="/admin/change-requests?status=all">All</a>
</div>
<div class="card"><table class="simple-table"><thead><tr><th>ID</th><th>File</th><th>Specialist</th><th>Type</th><th>Comment</th><th>Status</th><th>Actions</th></tr></thead><tbody>{''.join(rows) or '<tr><td colspan="7">No requests.</td></tr>'}</tbody></table></div>
</main>"""
    return _page(TEXT["requests"], body, user)


def _render_message(title: str, message: str, user=None) -> str:
    body = f"""<main><div class="card"><h1>{_escape(title)}</h1><p>{_escape(message)}</p><p><a href="/">Estimate AI</a></p></div></main>"""
    return _page(title, body, user)
