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
    async def grand_import(request: Request, estimate: UploadFile = File(...)) -> HTMLResponse:
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
        drafts[token] = {"record": record, "rows": rows, "directory": directory,
                         "name": estimate.filename or "original.gsfx",
                         "region": parsed.region, "normalization": summary}
        return HTMLResponse(_grand_preview_html(token, rows, parsed.region, summary=summary))

    @app.get("/grand/preview", response_class=HTMLResponse)
    def grand_preview(request: Request, token: str) -> HTMLResponse:
        draft = drafts.get(token)
        if draft is None:
            return HTMLResponse(render_error("Сессия не найдена", "Загрузите GSFX заново"), status_code=404)
        if not access_allowed(draft["record"], getattr(request.state, "user", None)):
            return HTMLResponse(render_error("Доступ запрещён", "Чужая сессия обработки"), status_code=403)
        return HTMLResponse(_grand_preview_html(token, draft["rows"], draft["region"],
                                                 summary=draft.get("normalization")))

    @app.post("/grand/preview", response_class=HTMLResponse)
    async def grand_edit(request: Request) -> HTMLResponse:
        form = await request.form()
        token = str(form.get("token", ""))
        draft = drafts.get(token)
        if draft is None:
            return HTMLResponse(render_error("Сессия не найдена", "Загрузите GSFX заново"), status_code=404)
        if not access_allowed(draft["record"], getattr(request.state, "user", None)):
            return HTMLResponse(render_error("Доступ запрещён", "Чужая сессия обработки"), status_code=403)
        rows = _grand_apply_form(draft["rows"], form)
        draft["rows"] = rows
        action = str(form.get("action", "save"))
        if action == "add":
            rows.append({"number": "", "chapter": "", "code": "", "name": "",
                         "unit": "", "quantity": "", "price": "", "source_id": "", "include": True,
                         "normalization_status": "manual", "normalization_note": "Добавлено вручную"})
        if action in ("xlsx", "match"):
            if action == "match":
                ambiguous = [r for r in rows if r["include"] and
                             r.get("normalization_status") == "review"]
                if ambiguous and str(form.get("review_confirm", "")) != "on":
                    return HTMLResponse(_grand_preview_html(token, rows, draft["region"],
                        "Есть позиции с неоднозначным соответствием базе аналогов. "
                        "Проверьте их и подтвердите проверку перед подбором.",
                        summary=draft.get("normalization")), status_code=422)
                missing = [str(i+1) for i, r in enumerate(rows) if r["include"] and
                           (not r["code"].strip() or not r["name"].strip() or not r["unit"].strip() or
                            _grand_positive_price(r["price"]) is None)]
                if missing:
                    return HTMLResponse(_grand_preview_html(token, rows, draft["region"],
                        "Для подбора нужны код, наименование, единица и положительная цена. "
                        "Исправьте строки: " + ", ".join(missing[:30]),
                        summary=draft.get("normalization")), status_code=422)
            output = draft["directory"] / "OS_from_GrandSmeta.xlsx"
            _grand_export(rows, output)
            if action == "xlsx":
                return FileResponse(output, media_type=xlsx_mime, filename="OS_from_GrandSmeta.xlsx")
            if not db_available():
                return HTMLResponse(_grand_preview_html(token, rows, draft["region"],
                     "База аналогов в SQLite не загружена. Скачайте ОС и используйте обычную загрузку с базой.",
                     summary=draft.get("normalization")), status_code=422)
            # Preserve the owner/user and invoke the normal confirmation workflow.
            record = draft["record"]
            record.estimate_path = output
            record.estimate_name = output.name
            app.state.app_state.store[token] = record
            return preview(app.state.app_state, token, selected_sheet=None)
        return HTMLResponse(_grand_preview_html(token, rows, draft["region"],
                                                summary=draft.get("normalization")))


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
        for key in ("number", "chapter", "code", "name", "unit", "quantity", "price"):
            row[key] = str(form.get(f"{key}_{i}", old[key]))[:2000].strip()
        if any(row[key] != old[key] for key in ("name", "unit", "quantity", "code")):
            row["normalization_status"] = "manual"
            row["normalization_note"] = "Изменено пользователем — проверьте количество и цену"
        row["include"] = f"include_{i}" in form
        result.append(row)
    return result


def _grand_excel_text(value: str) -> str:
    """Prevent uploaded text from becoming an executable Excel formula."""
    return "'" + value if value.lstrip().startswith(("=", "+", "-", "@")) else value


def _grand_export(rows: list[dict], target: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "ОС ГРАНД"
    sheet.append(["Шифр", "Наименование работ", "Ед. изм.", "Цена за ед.",
                  "Количество", "Стоимость всего", "№ позиции GSFX", "Раздел", "SysID",
                  "Наименование GSFX", "Единица GSFX", "Количество GSFX", "Нормализация"])
    for r in rows:
        if not r["include"]:
            continue
        price = _grand_positive_price(r["price"])
        try:
            qty = Decimal(r["quantity"].replace(" ", "").replace(",", "."))
            if not qty.is_finite():
                qty = None
        except (InvalidOperation, ValueError):
            qty = None
        sheet.append([_grand_excel_text(r["code"]), _grand_excel_text(r["name"]),
                      _grand_excel_text(r["unit"]), float(price) if price is not None else None,
                      float(qty) if qty is not None else None,
                      float(qty * price) if qty is not None and price is not None else None,
                      _grand_excel_text(r["number"]), _grand_excel_text(r["chapter"]),
                      _grand_excel_text(r["source_id"]),
                      _grand_excel_text(r.get("original_name", r["name"])),
                      _grand_excel_text(r.get("original_unit", r["unit"])),
                      _grand_excel_text(r.get("original_quantity", r["quantity"])),
                      _grand_excel_text(r.get("normalization_note", ""))])
    sheet.freeze_panes = "B2"
    sheet.auto_filter.ref = sheet.dimensions
    for col, width in {"A": 27, "B": 72, "C": 15, "D": 18, "E": 18, "F": 20,
                       "G": 18, "H": 40, "I": 18, "J": 64, "K": 16, "L": 18, "M": 50}.items():
        sheet.column_dimensions[col].width = width
    from openpyxl.styles import Font, PatternFill
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="243B53")
    for row in sheet.iter_rows(min_row=2):
        row[3].number_format = '#,##0.00'
        row[4].number_format = '#,##0.#######'
        row[5].number_format = '#,##0.00'
        if row[3].value is None:
            row[3].fill = PatternFill("solid", fgColor="FFF0C2")
    workbook.save(target)
    workbook.close()


def _grand_preview_html(token: str, rows: list[dict], region: str, error: str = "",
                        summary: NormalizationSummary | None = None) -> str:
    h = escape
    head = """<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>ГРАНД-Смета → ОС</title>
<style>body{font:14px system-ui,Arial;background:#f3f6fa;color:#152a40;margin:0;padding:25px}main{max-width:1550px;margin:auto;background:#fff;padding:23px;border-radius:14px}h1{margin:0 0 9px}p{line-height:1.5}.table{overflow:auto;max-height:68vh;border:1px solid #d8e0eb}table{border-collapse:collapse;min-width:1320px;width:100%}th{position:sticky;top:0;background:#243b53;color:white;text-align:left;padding:8px;z-index:1}td{border-bottom:1px solid #d8e0eb;padding:4px}input[type=text]{box-sizing:border-box;padding:7px;border:1px solid #c6d3e0;border-radius:5px;width:100%;min-width:65px}td:nth-child(4) input{min-width:250px}td:nth-child(6) input{min-width:95px}td:nth-child(8) input{min-width:105px}button{padding:12px 16px;margin:12px 10px 0 0;border:0;background:#1c5b9e;color:white;border-radius:8px;cursor:pointer}.muted{color:#50677b}.error{padding:12px;background:#ffeded;color:#9c1826;border-radius:8px}.warning{padding:12px;background:#fff3d5;color:#735000;border-radius:8px}</style></head><body><main><h1>ГРАНД-Смета → исходная ОС</h1>"""
    incomplete = sum(1 for r in rows if r["include"] and _grand_positive_price(r["price"]) is None)
    content = (f'<p class="muted">Позиций: {len(rows)} · Без подтверждённой цены: {incomplete} · Регион: {h(region)}</p>'
               '<p class="warning">GSFX не гарантирует готовую цену работы. Введите проверенные цены в предпросмотре или скачайте Excel для заполнения. Пустые цены не заменяются нулём и не допускаются к подбору.</p>')
    if summary is not None:
        if summary.available:
            statuses = {"changed": 0, "review": 0, "matched": 0, "no_match": 0}
            for r in rows:
                status = r.get("normalization_status", "no_match")
                statuses[status] = statuses.get(status, 0) + 1
            content += ('<p class="muted">Источник нормализации — только база аналогов (main): '
                        + str(summary.catalog_items) + ' записей. Автоматически исправлено: '
                        + str(statuses["changed"]) + ' · Требует проверки: '
                        + str(statuses["review"]) + ' · Без совпадения: '
                        + str(statuses["no_match"]) + '.</p>')
        else:
            content += '<p class="warning">База аналогов не загружена или недоступна: нормализация не применена. Смета доступна для ручной правки.</p>'
    if error:
        content += '<p class="error">' + h(error) + '</p>'
    content += f'<form action="/grand/preview" method="post"><input type="hidden" name="token" value="{h(token)}"><div class="table"><table><thead><tr>'
    for label in ('В ОС', '№', 'Шифр', 'Наименование работы', 'Ед.', 'Количество', 'Раздел', 'Цена за ед., ₽', 'Сверка с базой аналогов'):
        content += '<th>' + label + '</th>'
    content += '</tr></thead><tbody>'
    for i, r in enumerate(rows):
        content += '<tr><td><input type="checkbox" name="include_' + str(i) + '" ' + ('checked' if r['include'] else '') + '></td>'
        for key in ('number', 'code', 'name', 'unit', 'quantity', 'chapter', 'price'):
            content += f'<td><input type="text" name="{key}_{i}" value="{h(r[key], quote=True)}"></td>'
        status = r.get("normalization_status", "")
        descriptions = {"changed": "Приведено по базе", "review": "Проверьте",
                        "no_match": "Нет совпадения", "matched": "Совпадает", "manual": "Ручная правка"}
        original = ("Исходно: " + r.get("original_unit", r["unit"]) + ", "
                    + r.get("original_quantity", r["quantity"]))
        detail = r.get("normalization_note", "")
        alternative = r.get("analog_suggestion", "")
        content += ('<td style="min-width:220px"><strong>' + h(descriptions.get(status, ""))
                    + '</strong><div class="muted">' + h(detail) + '</div>'
                    + '<small class="muted">' + h(original) + '</small>')
        if alternative and alternative != r["name"]:
            content += '<div class="muted">В базе: ' + h(alternative) + '</div>'
        if r.get("original_name") and r["original_name"] != r["name"]:
            content += '<div class="muted">В GSFX: ' + h(r["original_name"]) + '</div>'
        content += '</td></tr>'
    content += '</tbody></table></div>'
    if any(r.get("normalization_status") == "review" and r["include"] for r in rows):
        content += ('<p class="warning"><label><input type="checkbox" name="review_confirm"> '
                    'Я проверил(а) неоднозначные единицы и наименования, включая масштаб цен, '
                    'перед запуском подбора.</label></p>')
    content += ('<p class="muted">Цена вводится за указанную единицу измерения. '
                'При ручной смене единицы проверьте и количество, и цену.</p>')
    content += '<button type="submit" name="action" value="save">Сохранить изменения</button><button type="submit" name="action" value="add">Добавить строку</button><button type="submit" name="action" value="xlsx">Скачать ОС в Excel</button><button type="submit" name="action" value="match">Продолжить к подбору аналогов</button></form><p><a href="/">Вернуться к загрузке</a></p></main></body></html>'
    return head + content
