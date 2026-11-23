"""Administrator moderation for quarantined GSFX/OS training examples."""
from __future__ import annotations

from html import escape
import json

from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response

from core.storage.connection import connect, default_database_path, init_database
from core.storage.lsr_learning import (
    get_learning_source, list_learning_cases, review_learning_case,
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
        finally:
            connection.close()
        lines = []
        for case in cases:
            case_id = int(case['id'])
            issues = json.loads(case['issues_json'])
            note = '; '.join(issues[:3]) or 'Основные показатели совпадают'
            decision = ''
            if case['status'] == 'pending':
                if case['quality'] == 'candidate':
                    decision += (f'<form method="post" action="/admin/lsr-learning/{case_id}/review">'
                                 '<input type="hidden" name="action" value="approved">'
                                 '<button type="submit">Подтвердить</button></form>')
                decision += (f'<form method="post" action="/admin/lsr-learning/{case_id}/review">'
                             '<input type="hidden" name="action" value="rejected">'
                             '<button type="submit">Отклонить</button></form>')
            lines.append('<tr>' + ''.join(f'<td>{escape(str(v))}</td>' for v in (
                case_id, case['created_at'], case['author_name'], case['os_filename'],
                case['quality'], case['status'],
                f"{case['matched_count']}/{case['os_count']}", case['checked_prices'], note,
            )) + f'<td><a href="/admin/lsr-learning/{case_id}/download">Скачать GSFX</a>{decision}</td></tr>')
        table = '<table border="1" cellpadding="6"><tr><th>ID</th><th>Дата</th><th>Специалист</th><th>ОС</th><th>Качество</th><th>Статус</th><th>Совпало</th><th>Цены</th><th>Причины проверки</th><th>Действия</th></tr>' + ''.join(lines) + '</table>'
        return HTMLResponse('''<!doctype html><html lang="ru"><head><meta charset="utf-8">
            <title>ЛСР ↔ ОС — проверка</title></head><body>
            <h1>Сверка ЛСР и ОС — обучающие материалы</h1>
            <p><a href="/admin">К админке</a></p>
            <p>Все материалы сначала изолированы от базы аналогов. Утвердить можно только
            прошедшие структурную и ценовую проверку; остальные остаются на разбор.</p>'''
            + table + '</body></html>')

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
        if not _admin(request):
            return HTMLResponse('Доступ только администратору', status_code=403)
        connection = connect(default_database_path())
        try:
            init_database(connection, synchronize_corrections=False)
            content = get_learning_source(connection, case_id)
        finally:
            connection.close()
        if content is None:
            return HTMLResponse('Пример не найден', status_code=404)
        return Response(content, media_type='application/octet-stream',
                        headers={'Content-Disposition': f'attachment; filename="lsr_{case_id}.gsfx"',
                                 'X-Content-Type-Options': 'nosniff'})
