"""Isolated GrandSmeta UI routes over the existing OS import pipeline.

GSFX parsing and catalog normalization are read-only. Existing matching,
upload ownership rules, and RNMC/TKP imports remain unchanged.
"""
from __future__ import annotations

from decimal import Decimal, InvalidOperation
from html import escape
from pathlib import Path
import sqlite3
import uuid

from fastapi import FastAPI, File, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from openpyxl import Workbook

from app.services.grand_rim_report import RimReportError, read_filled_rim_report, attach_verified_rim_prices
from app.services.grand_os_export import export_grand_os
from app.services.grand_normalization import (
    load_catalog_index, normalize_rows, NormalizationSummary, CatalogIndex,
)
from core.grand_smeta import read_gsfx, GrandSmetaFormatError
from app.web.rendering import render_error


def install_grand_routes(app: FastAPI, *, record_type, preview, db_path,
                         db_available, access_allowed, max_upload_bytes: int,
                         xlsx_mime: str) -> None:
    """Register GSFX import without replacing any existing routes."""
    drafts: dict[str, dict] = {}
    app.state.grand_drafts = drafts
    app.state.app_state.grand_drafts = drafts

    @app.post("/grand/import", response_class=HTMLResponse)
    async def grand_import(request: Request, estimate: UploadFile = File(...), rim_export: UploadFile | None = File(None)) -> HTMLResponse:
        if not (estimate.filename or "").lower().endswith(".gsfx"):
            return HTMLResponse(render_error("Неверный формат", "Выберите файл .gsfx"), status_code=400)
        data = await estimate.read(max_upload_bytes + 1)
        if not data or len(data) > max_upload_bytes:
            return HTMLResponse(render_error("Неверный размер", "Файл пустой или слишком большой"), status_code=400)
        state = app.state.app_state
        token = uuid.uuid4().hex
        directory = state.base_dir / token
        directory.mkdir(parents=True, exist_ok=True)
        source = directory / "original.gsfx"
        source.write_bytes(data)
        try:
            parsed = read_gsfx(source)
        except GrandSmetaFormatError as exc:
            source.unlink(missing_ok=True)
            return HTMLResponse(render_error("Не удалось прочитать GSFX", str(exc)), status_code=422)
        if len(parsed.positions) > 10000:
            return HTMLResponse(render_error("Слишком много позиций", "Лимит 10000 строк"), status_code=422)
        rows = [{"number": p.number, "chapter": p.chapter,
                 "code": p.code, "name": p.name, "unit": p.unit,
                 "quantity": str(p.quantity) if p.quantity is not None else "",
                 "price": "", "source_id": p.source_id, "include": True}
                for p in parsed.positions]
        report_result = None
        if rim_export is not None and rim_export.filename:
            if not rim_export.filename.lower().endswith(".xlsx"):
                return HTMLResponse(render_error("Неверный формат", "Нужен заполненный Excel-отчёт РИМ в формате .xlsx"), status_code=400)
            report_bytes = await rim_export.read(max_upload_bytes + 1)
            if not report_bytes or len(report_bytes) > max_upload_bytes:
                return HTMLResponse(render_error("Неверный размер", "Файл РИМ пустой или слишком большой"), status_code=400)
            report_path = directory / "report_rim.xlsx"
            report_path.write_bytes(report_bytes)
            try:
                report_result = attach_verified_rim_prices(rows, read_filled_rim_report(report_path))
            except RimReportError as exc:
                return HTMLResponse(render_error("Сверка GSFX и РИМ", str(exc)), status_code=422)
        try:
            catalog_index = load_catalog_index(db_path())
            summary = normalize_rows(rows, catalog_index)
        except (sqlite3.DatabaseError, OSError):
            summary = normalize_rows(rows, CatalogIndex({}, 0))
            summary = NormalizationSummary(False, 0, summary.changed, summary.review,
                                           summary.untouched, "База аналогов недоступна")
        current_user = getattr(request.state, "user", None)
        record = record_type(directory=directory, catalog_path=None,
                             estimate_path=source,
                             estimate_name=estimate.filename or "original.gsfx",
                             use_database_catalog=True,
                             owner_user_id=(int(current_user.id) if current_user else None))
        metadata = {"local_number": parsed.local_number, "object_name": parsed.object_name,
                    "description": parsed.description, "region": parsed.region,
                    "task_number": "", "coefficient": ""}
        drafts[token] = {"record": record, "rows": rows, "directory": directory,
                         "name": estimate.filename or "original.gsfx",
                         "region": parsed.region, "normalization": summary,
                         "rim_result": report_result, "metadata": metadata}
        return HTMLResponse(_grand_preview_html(token, rows, parsed.region,
             summary=summary, rim_result=report_result, metadata=metadata))

    @app.get("/grand/preview", response_class=HTMLResponse)
    def grand_preview(request: Request, token: str) -> HTMLResponse:
        draft = drafts.get(token)
        if draft is None:
            return HTMLResponse(render_error("Сессия не найдена", "Загрузите GSFX заново"), status_code=404)
        if not access_allowed(draft["record"], getattr(request.state, "user", None)):
            return HTMLResponse(render_error("Доступ запрещён", "Чужая сессия обработки"), status_code=403)
        return HTMLResponse(_grand_preview_html(token, draft["rows"], draft["region"],
                                                 summary=draft.get("normalization"), rim_result=draft.get("rim_result"), metadata=draft.get("metadata")))

    @app.post("/grand/preview", response_class=HTMLResponse)
    async def grand_edit(request: Request) -> HTMLResponse:
        form = await request.form()
        token = str(form.get("token", ""))
        draft = drafts.get(token)
        if draft is None:
            return HTMLResponse(render_error("Сессия не найдена", "Загрузите GSFX заново"), status_code=404)
        if not access_allowed(draft["record"], getattr(request.state, "user", None)):
            return HTMLResponse(render_error("Доступ запрещён", "Чужая сессия обработки"), status_code=403)
        metadata = draft.setdefault("metadata", {})
        for field in ("task_number", "object_name", "description", "region", "coefficient"):
            if field in form:
                metadata[field] = str(form[field]).strip()[:2000]
        draft["region"] = metadata.get("region", draft["region"])
        rows = _grand_apply_form(draft["rows"], form)
        draft["rows"] = rows
        action = str(form.get("action", "save"))
        if action == "add":
            rows.append({"number": "", "chapter": "", "code": "", "name": "",
                         "unit": "", "quantity": "", "price": "", "source_id": "", "include": True,
                         "normalization_status": "manual", "normalization_note": "Добавлено вручную", "section_code": ""})
        if action in ("xlsx", "match"):
            if action == "match":
                ambiguous = [r for r in rows if r["include"] and
                             r.get("normalization_status") == "review"]
                if ambiguous and str(form.get("review_confirm", "")) != "on":
                    return HTMLResponse(_grand_preview_html(token, rows, draft["region"],
                        "Есть позиции с неоднозначным соответствием базе аналогов. "
                        "Проверьте их и подтвердите проверку перед подбором.",
                        summary=draft.get("normalization"), rim_result=draft.get("rim_result"), metadata=draft.get("metadata")), status_code=422)
                missing = [str(i+1) for i, r in enumerate(rows) if r["include"] and
                           (not r["code"].strip() or not r["name"].strip() or not r["unit"].strip() or
                            _grand_positive_price(r["price"]) is None)]
                if missing:
                    return HTMLResponse(_grand_preview_html(token, rows, draft["region"],
                        "Для подбора нужны код, наименование, единица и положительная цена. "
                        "Исправьте строки: " + ", ".join(missing[:30]),
                        summary=draft.get("normalization"), rim_result=draft.get("rim_result"), metadata=draft.get("metadata")), status_code=422)
            output = draft["directory"] / "OS_from_GrandSmeta.xlsx"
            _grand_export(rows, output, metadata=draft.get("metadata"))
            if action == "xlsx":
                return FileResponse(output, media_type=xlsx_mime, filename="OS_from_GrandSmeta.xlsx")
            if not db_available():
                return HTMLResponse(_grand_preview_html(token, rows, draft["region"],
                     "База аналогов в SQLite не загружена. Скачайте ОС и используйте обычную загрузку с базой.",
                     summary=draft.get("normalization"), rim_result=draft.get("rim_result"), metadata=draft.get("metadata")), status_code=422)
            # Preserve the owner/user and invoke the normal confirmation workflow.
            record = draft["record"]
            record.estimate_path = output
            record.estimate_name = output.name
            app.state.app_state.store[token] = record
            return preview(app.state.app_state, token, selected_sheet=None)
        return HTMLResponse(_grand_preview_html(token, rows, draft["region"],
                                                summary=draft.get("normalization"), rim_result=draft.get("rim_result"), metadata=draft.get("metadata")))


def _grand_positive_price(raw: str) -> Decimal | None:
    try:
        value = Decimal(str(raw).replace(" ", "").replace("\u00a0", "").replace(",", "."))
        return value if value.is_finite() and value > 0 else None
    except (InvalidOperation, ValueError):
        return None


def _grand_apply_form(previous: list[dict], form) -> list[dict]:
    """Only edit existing indexed rows; original GSFX fields stay server-side."""
    result = []
    for i, old in enumerate(previous):
        row = old.copy()
        for key in ("number", "chapter", "code", "name", "unit", "quantity", "price", "section_code"):
            row[key] = str(form.get(f"{key}_{i}", old.get(key, "")))[:2000].strip()
        if any(row[key] != old[key] for key in ("name", "unit", "quantity", "code")):
            row["normalization_status"] = "manual"
            row["normalization_note"] = "Изменено пользователем — проверьте количество и цену"
        row["include"] = f"include_{i}" in form
        result.append(row)
    return result


def _grand_excel_text(value: str) -> str:
    """Prevent uploaded text from becoming an executable Excel formula."""
    return "'" + value if value.lstrip().startswith(("=", "+", "-", "@")) else value


def _grand_export(rows: list[dict], target: Path, *, metadata: dict | None = None) -> None:
    """Expose the standard RIM-compatible OS writer as a stable helper."""
    export_grand_os(rows, target, metadata=metadata)


def _grand_preview_html(token: str, rows: list[dict], region: str, error: str = "",
                        summary: NormalizationSummary | None = None,
                        rim_result: tuple[int, int] | None = None,
                        metadata: dict | None = None) -> str:
    h = escape
    header = """<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>ГРАНД-Смета → ОС</title><style>
    body{font:14px system-ui,Arial;background:#f3f6fa;color:#152a40;margin:0;padding:24px}
    main{max-width:1650px;margin:auto;background:#fff;border-radius:12px;padding:22px}
    h1{font-size:24px}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:12px}
    input[type=text]{box-sizing:border-box;width:100%;padding:7px;border:1px solid #b4c3d0;border-radius:5px}
    .table{overflow:auto;max-height:69vh;border:1px solid #d9e1ed}
    table{min-width:1900px;border-collapse:collapse;width:100%}th{position:sticky;top:0;background:#243b53;color:white;padding:9px;text-align:left}
    td{border-bottom:1px solid #d9e1ed;padding:5px;vertical-align:top}
    td:nth-child(3) input{min-width:280px}.muted{color:#566b82}.warn{background:#fff0d1;padding:10px}
    .err{background:#ffdddd;padding:10px}button{background:#1c5b9e;color:#fff;border:0;border-radius:6px;padding:10px 13px;margin:8px 8px 0 0;cursor:pointer}
    </style></head><body><main><h1>ГРАНД-Смета → исходная ОС</h1>"""
    missing_prices = sum(1 for row in rows if row.get('include', True) and _grand_positive_price(row.get('price', '')) is None)
    result = f'<p class="muted">Позиций: {len(rows)} · Без проверенной цены: {missing_prices} · Регион: {h(region)}</p>'
    if rim_result is None:
        result += '<p class="warn">GSFX не содержит готовую цену работы. Для автоматического заполнения нужен рассчитанный Excel-отчёт формы РИМ, экспортированный ГРАНД-Сметой. Стоимость ресурсов не подменяет цену позиции.</p>'
    else:
        result += f'<p>Сверка с заполненным отчётом РИМ: цены для {rim_result[0]} позиций; без цены {rim_result[1]}. Неподтверждённые позиции остаются для проверки.</p>'
    if summary and summary.available:
        result += f'<p class="muted">Нормализация: только база аналогов (main), {summary.catalog_items} записей. Изменено: {summary.changed}; требует проверки: {summary.review}.</p>'
    else:
        result += '<p class="warn">База аналогов недоступна; наименования и единицы не нормализованы.</p>'
    if error:
        result += '<p class="err">' + h(error) + '</p>'
    result += f'<form method="post" action="/grand/preview"><input type="hidden" name="token" value="{h(token, quote=True)}">'
    result += '<h3>Реквизиты ОС (над телом таблицы)</h3><div class="grid">'
    meta = metadata or {}
    for key, label in (('task_number','Номер задачи'),('object_name','Объект'),
                       ('description','Техническое задание / вид работ'),
                       ('region','Регион'),('coefficient','Коэффициент')):
        value = str(meta.get(key, region if key == 'region' else '') or '')
        result += f'<label>{h(label)}<input type="text" name="{key}" value="{h(value, quote=True)}"></label>'
    result += '</div><p class="muted">N — исходный шифр ГЭСН/ФЕР/ТЕР; O — отдельный /КР (после подбора); P — код раздела.</p>'
    result += '<div class="table"><table><thead><tr>'
    for name in ('В ОС','№ в смете','Наименование работ','Ед. изм.','Кол-во',
                 'Цена единицы без НДС','Итого без НДС','ТЗ ед.','ТЗм ед.',
                 'ТЗ всего','ТЗм всего','Перечень ГЭСН/ФЕР/ТЕР',
                 '/КР после подбора','Код раздела','Статус'):
        result += '<th>' + name + '</th>'
    result += '</tr></thead><tbody>'
    for i, row in enumerate(rows):
        result += '<tr><td><input type="checkbox" name="include_' + str(i) + '" ' + ('checked' if row.get('include', True) else '') + '></td>'
        for key in ('number','name','unit','quantity','price'):
            result += f'<td><input type="text" name="{key}_{i}" value="{h(str(row.get(key, "")), quote=True)}"></td>'
        price = _grand_positive_price(row.get('price', ''))
        try:
            qty = Decimal(str(row.get('quantity', '')).replace(',', '.'))
            total = format(qty * price, '.2f') if qty.is_finite() and price is not None else '—'
        except (ValueError, InvalidOperation):
            total = '—'
        result += '<td>' + h(total) + '</td>'
        for key in ('labor_unit','machinist_labor_unit','labor_total','machinist_labor_total'):
            result += '<td>' + h(str(row.get(key, ''))) + '</td>'
        result += f'<td><input type="text" name="code_{i}" value="{h(str(row.get("code", "")), quote=True)}"></td>'
        result += '<td class="muted">Шифр/КР при наличии аналога, иначе исходный шифр</td>'
        result += f'<td><input type="text" name="section_code_{i}" value="{h(str(row.get("section_code", "")), quote=True)}"></td>'
        info = str(row.get('normalization_note', ''))
        if row.get('price_origin'):
            info += ('; ' if info else '') + str(row['price_origin'])
        result += '<td class="muted">' + h(info) + '</td></tr>'
    result += '</tbody></table></div>'
    if any(row.get('include', True) and row.get('normalization_status') == 'review' for row in rows):
        result += '<p class="warn"><label><input type="checkbox" name="review_confirm">Я проверил(а) неоднозначные соответствия и масштаб цен перед подбором.</label></p>'
    result += '<p class="muted">Если меняете единицу вручную, проверьте и цену, и количество.</p>'
    for action,label in (('save','Сохранить изменения'),('add','Добавить строку'),
                         ('xlsx','Скачать ОС в Excel'),('match','Продолжить к подбору аналогов')):
        result += f'<button name="action" value="{action}" type="submit">{label}</button>'
    return header + result + '</form><p><a href="/">Вернуться к загрузке</a></p></main></body></html>'
