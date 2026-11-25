"""Administrator moderation for quarantined GSFX/OS training examples."""
from __future__ import annotations

from html import escape
import json

from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response

from app.services.lsr_pair_export import export_verified_pairs
from core.storage.connection import connect, default_database_path, init_database
from core.storage.lsr_learning import (
    get_learning_source, get_learning_workbook, list_learning_cases,
    list_learning_source_names, review_learning_case,
)


def _admin(request: Request) -> bool:
    user = getattr(request.state, 'user', None)
    return user is not None and getattr(user, 'role', '') == 'admin'


def install_lsr_learning_routes(app: FastAPI) -> None:
    @app.get('/admin/lsr-learning', response_class=HTMLResponse)
    def learning_index(request: Request):
        if not _admin(request):
            return HTMLResponse('Доступ только администратору', status_code=403)
        connection = connect(default_database_path())
        try:
            init_database(connection, synchronize_corrections=False)
            cases = list_learning_cases(connection)
            source_names = {int(case['id']): list_learning_source_names(connection, int(case['id']))
                            for case in cases}
        finally:
            connection.close()
        lines = []
        for case in cases:
            case_id = int(case['id'])
            issues = json.loads(case['issues_json'])
            note = '; '.join(issues[:3]) or 'Основные показатели совпадают'
            decision = ''
            if case['status'] == 'pending':
                if case['quality'] == 'candidate' and case['has_original_os']:
                    decision += (f'<form method="post" action="/admin/lsr-learning/{case_id}/review">'
                                 '<input type="hidden" name="action" value="approved">'
                                 '<button type="submit">Подтвердить</button></form>')
                decision += (f'<form method="post" action="/admin/lsr-learning/{case_id}/review">'
                             '<input type="hidden" name="action" value="rejected">'
                             '<button type="submit">Отклонить</button></form>')
            source_links = ' '.join(
                f'<a href="/admin/lsr-learning/{case_id}/download'
                + (f'/{number}' if number > 1 else '') + '">'
                + escape(f'GSFX {number}: {filename}' if len(source_names[case_id]) > 1
                         else 'GSFX') + '</a>'
                for number, filename in source_names[case_id])
            lines.append('<tr>' + ''.join(f'<td>{escape(str(v))}</td>' for v in (
                case_id, case['created_at'], case['author_name'], case['os_filename'],
                case['source_count'], case['quality'], case['status'],
                f"{case['matched_count']}/{case['os_count']}", case['checked_prices'], note,
            )) + (f'<td>{source_links} '
                  + (f'<a href="/admin/lsr-learning/{case_id}/download-os">ОС</a>'
                     if case['has_original_os'] else 'Оригинал ОС не сохранён')
                  + decision + '</td></tr>'))
        table = '<table border="1" cellpadding="6"><tr><th>ID</th><th>Дата</th><th>Специалист</th><th>ОС</th><th>ЛСР, шт.</th><th>Качество</th><th>Статус</th><th>Совпало</th><th>Цены</th><th>Причины проверки</th><th>Действия</th></tr>' + ''.join(lines) + '</table>'
        return HTMLResponse('''<!doctype html><html lang="ru"><head><meta charset="utf-8">
            <title>ЛСР ↔ ОС — проверка</title></head><body>
            <h1>Проверенные пары ЛСР и ОС</h1>
            <p><a href="/admin">К админке</a></p>
            <p><a href="/admin/lsr-learning/export">Скачать утверждённые пары (GSFX + XLSX)</a></p>
            <p>Это архив исходных данных для последующего анализа, а не механизм самообучения.
            Все материалы изолированы от базы аналогов. Утвердить можно только
            прошедшие структурную и ценовую проверку и имеющие оба исходных файла.</p>'''
            + table + '</body></html>')

    @app.get('/admin/lsr-learning/export')
    def learning_batch_export(request: Request, after_id: int = 0, limit: int = 50):
        if not _admin(request):
            return HTMLResponse('Доступ только администратору', status_code=403)
        connection = connect(default_database_path())
        try:
            init_database(connection, synchronize_corrections=False)
            try:
                exported = export_verified_pairs(connection, after_id=after_id, limit=limit)
            except ValueError as error:
                return HTMLResponse(escape(str(error)), status_code=422)
        finally:
            connection.close()
        return Response(exported.content, media_type='application/zip', headers={
            'Content-Disposition': 'attachment; filename="estimate_ai_verified_pairs.zip"',
            'Cache-Control': 'private, no-store', 'X-Content-Type-Options': 'nosniff',
        })

    @app.post('/admin/lsr-learning/{case_id}/review')
    def learning_review(request: Request, case_id: int,
                        action: str = Form(...), comment: str = Form('')):
        if not _admin(request):
            return HTMLResponse('Доступ только администратору', status_code=403)
        if action not in ('approved', 'rejected'):
            return HTMLResponse('Неверное решение', status_code=400)
        connection = connect(default_database_path())
        try:
            init_database(connection, synchronize_corrections=False)
            changed = review_learning_case(connection, case_id=case_id,
                action=action, reviewer_user_id=request.state.user.id, comment=comment)
        finally:
            connection.close()
        if not changed:
            return HTMLResponse('Нельзя утвердить эту запись без достоверной сверки или решение уже принято', status_code=409)
        return RedirectResponse('/admin/lsr-learning', status_code=303)

    @app.get('/admin/lsr-learning/{case_id}/download')
    def learning_source(request: Request, case_id: int):
        return _download_lsr(request, case_id, 1)

    @app.get('/admin/lsr-learning/{case_id}/download/{source_index}')
    def learning_additional_source(request: Request, case_id: int, source_index: int):
        return _download_lsr(request, case_id, source_index)

    def _download_lsr(request: Request, case_id: int, source_index: int):
        if not _admin(request):
            return HTMLResponse('Доступ только администратору', status_code=403)
        if source_index < 1:
            return HTMLResponse('Пример не найден', status_code=404)
        connection = connect(default_database_path())
        try:
            init_database(connection, synchronize_corrections=False)
            content = get_learning_source(connection, case_id, source_index)
        finally:
            connection.close()
        if content is None:
            return HTMLResponse('Пример не найден', status_code=404)
        return Response(content, media_type='application/octet-stream',
                        headers={'Content-Disposition': f'attachment; filename="lsr_{case_id}_{source_index:02}.gsfx"',
                                 'Cache-Control': 'private, no-store', 'X-Content-Type-Options': 'nosniff'})

    @app.get('/admin/lsr-learning/{case_id}/download-os')
    def learning_workbook(request: Request, case_id: int):
        if not _admin(request):
            return HTMLResponse('Доступ только администратору', status_code=403)
        connection = connect(default_database_path())
        try:
            init_database(connection, synchronize_corrections=False)
            content = get_learning_workbook(connection, case_id)
        finally:
            connection.close()
        if content is None:
            return HTMLResponse('Оригинал ОС недоступен', status_code=404)
        return Response(content, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        headers={'Content-Disposition': f'attachment; filename="os_{case_id}.xlsx"',
                                 'Cache-Control': 'private, no-store', 'X-Content-Type-Options': 'nosniff'})
