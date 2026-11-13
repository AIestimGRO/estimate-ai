"""GSFX upload/edit/export and position 163 regression checks."""
from io import BytesIO
import pytest
from zipfile import ZipFile
from fastapi.testclient import TestClient
from openpyxl import load_workbook
from app.web.app import create_app
from app.services.read_estimate import load_estimate

@pytest.fixture(autouse=True)
def isolated_auth_database(tmp_path, monkeypatch):
    """Never consult or migrate a real installation's authentication database.

    Production has users, which means guest requests redirect to /login.
    These unauthenticated feature tests intentionally use an isolated empty
    database, matching the established tests/test_web.py fixture.
    """
    monkeypatch.setenv("ESTIMATE_AI_DB_PATH", str(tmp_path / "test_auth.sqlite"))


XML = '''<Document Generator="GrandSmeta"><Properties LocNum="L-1"/><Chapters>
<Chapter Caption="Раздел"><Position Number="163" SysID="abc" Code="43-1"
Caption="Погрузка в автотранспортное средство" Units="1т груза"><Quantity Result="1225"/>
</Position><Position Number="164" SysID="def" Code="02-15-1" Caption="Перевозка"
Units="1т груза"><Quantity Result="1225"/></Position></Chapter></Chapters></Document>'''

def _gsfx():
    out = BytesIO()
    with ZipFile(out, 'w') as z:
        z.writestr('Properties.txt', 'GSFX')
        z.writestr('Data.xml', XML)
    return out.getvalue()

def _form(token, rows, *, price='103,54', action='xlsx'):
    data = {'token': token, 'action': action}
    for i, row in enumerate(rows):
        data[f'include_{i}'] = 'on'
        for key in ('number', 'chapter', 'code', 'name', 'unit', 'quantity', 'price'):
            data[f'{key}_{i}'] = price if key == 'price' else row[key]
    return data

def test_grand_upload_preview_edit_export_and_normal_reader(tmp_path):
    with TestClient(create_app(base_dir=tmp_path)) as client:
        assert 'ГРАНД-Сметы' in client.get('/').text
        response = client.post('/grand/import', files={'estimate': ('example.gsfx', _gsfx())})
        assert response.status_code == 200
        assert 'Погрузка в автотранспортное средство' in response.text
        token = next(iter(client.app.state.app_state.grand_drafts))
        rows = client.app.state.app_state.grand_drafts[token]['rows']
        assert len(rows) == 2 and rows[0]['number'] == '163'
        assert client.post('/grand/preview', data=_form(token, rows, price='', action='match')).status_code == 422
        output = client.post('/grand/preview', data=_form(token, rows))
        assert output.status_code == 200
        path = tmp_path / 'output.xlsx'
        path.write_bytes(output.content)
        parsed = load_estimate(path)
        assert len(parsed.positioned_rows) == 2
        assert parsed.positioned_rows[0][1].code == '43-1'
        assert parsed.positioned_rows[0][1].base_price == 103.54
        wb = load_workbook(path, data_only=True)
        assert wb.active['E2'].value == 1225
        assert wb.active['G2'].value == '163'
        wb.close()

def test_bad_gsfx_and_add_row(tmp_path):
    with TestClient(create_app(base_dir=tmp_path)) as client:
        assert client.post('/grand/import', files={'estimate': ('bad.xlsx', b'abc')}).status_code == 400
        assert client.post('/grand/import', files={'estimate': ('bad.gsfx', b'abc')}).status_code == 422
        client.post('/grand/import', files={'estimate': ('good.gsfx', _gsfx())})
        token = next(iter(client.app.state.app_state.grand_drafts))
        rows = client.app.state.app_state.grand_drafts[token]['rows']
        result = client.post('/grand/preview', data=_form(token, rows, action='add'))
        assert result.status_code == 200
        assert len(client.app.state.app_state.grand_drafts[token]['rows']) == 3

def test_grand_import_normalizes_from_analog_db_not_tkp_and_exports_trace(tmp_path, monkeypatch):
    from importlib import import_module
    from sqlite3 import connect
    mod = import_module('app.web.app')
    db_path = tmp_path / 'catalog.sqlite'
    with connect(db_path) as db:
        db.executescript('''
            CREATE TABLE catalog_sources(id INTEGER PRIMARY KEY, name TEXT);
            CREATE TABLE catalog_items(source_id INTEGER, code TEXT, work_name TEXT, unit TEXT);
            CREATE TABLE tkp_items(item_code TEXT, item_name TEXT, unit TEXT);
            INSERT INTO catalog_sources VALUES (1, 'main'), (2, 'foreign');
            INSERT INTO catalog_items VALUES
              (1, '43-1', 'Погрузка в автотранспортное средство', 'т'),
              (2, '43-1', 'Другая работа', 'км');
            INSERT INTO tkp_items VALUES ('43-1','Не использовать','тыс. т');
        ''')
    old = db_path.read_bytes()
    monkeypatch.setattr(mod, 'default_database_path', lambda: db_path)
    with TestClient(create_app(base_dir=tmp_path / 'sessions')) as client:
        response = client.post('/grand/import', files={'estimate': ('grand.gsfx', _gsfx())})
        assert response.status_code == 200
        assert 'Источник нормализации — только база аналогов (main)' in response.text
        token = next(iter(client.app.state.app_state.grand_drafts))
        rows = client.app.state.app_state.grand_drafts[token]['rows']
        assert rows[0]['unit'] == 'т'
        assert rows[0]['quantity'] == '1225'
        assert rows[0]['price'] == ''
        assert rows[0]['original_unit'] == '1т груза'
        assert rows[1]['unit'] == '1т груза'
        output = client.post('/grand/preview', data=_form(token, rows))
        assert output.status_code == 200
        path = tmp_path / 'normalized.xlsx'
        path.write_bytes(output.content)
        wb = load_workbook(path, data_only=True)
        assert wb.active['C2'].value == 'т'
        assert wb.active['D2'].value == 103.54
        assert wb.active['K2'].value == '1т груза'
        wb.close()
        assert len(load_estimate(path).positioned_rows) == 2
    assert db_path.read_bytes() == old

def test_grand_review_requires_explicit_confirmation(tmp_path, monkeypatch):
    from importlib import import_module
    from sqlite3 import connect
    mod = import_module('app.web.app')
    db_path = tmp_path / 'catalog.sqlite'
    with connect(db_path) as db:
        db.executescript('''
            CREATE TABLE catalog_sources(id INTEGER PRIMARY KEY, name TEXT);
            CREATE TABLE catalog_items(source_id INTEGER, code TEXT, work_name TEXT, unit TEXT);
            INSERT INTO catalog_sources VALUES (1, 'main');
            INSERT INTO catalog_items VALUES
              (1, '43-1', 'Погрузка в автотранспортное средство', 'т'),
              (1, '43-1', 'Погрузка в автотранспортное средство', '100 т');
        ''')
    monkeypatch.setattr(mod, 'default_database_path', lambda: db_path)
    with TestClient(create_app(base_dir=tmp_path / 'upload')) as client:
        assert client.post('/grand/import', files={'estimate': ('grand.gsfx', _gsfx())}).status_code == 200
        token = next(iter(client.app.state.app_state.grand_drafts))
        rows = client.app.state.app_state.grand_drafts[token]['rows']
        assert rows[0]['normalization_status'] == 'review'
        data = _form(token, rows, action='match')
        blocked = client.post('/grand/preview', data=data)
        assert blocked.status_code == 422
        assert 'подтвердите проверку' in blocked.text
        data['review_confirm'] = 'on'
        attempted = client.post('/grand/preview', data=data)
        assert attempted.status_code == 422
        assert 'База аналогов в SQLite не загружена' in attempted.text

def test_grand_authentication_and_cross_user_draft_access(tmp_path):
    """Real auth guard stays enabled; users must not read each other's drafts."""
    from core.storage.accounts import ROLE_SPECIALIST, create_session, create_user
    from core.storage.connection import connect, init_database
    from app.web.workspace import SESSION_COOKIE

    connection = connect(tmp_path / "test_auth.sqlite")
    try:
        init_database(connection)
        first_id = create_user(
            connection,
            full_name="GSFX User One",
            login="gsfx.one",
            password="gsfx-test-one",
            role=ROLE_SPECIALIST,
        )
        second_id = create_user(
            connection,
            full_name="GSFX User Two",
            login="gsfx.two",
            password="gsfx-test-two",
            role=ROLE_SPECIALIST,
        )
        first_session = create_session(connection, first_id)
        second_session = create_session(connection, second_id)
    finally:
        connection.close()

    with TestClient(create_app(base_dir=tmp_path / "auth_uploads")) as client:
        guest = client.post(
            "/grand/import",
            files={"estimate": ("sample.gsfx", _gsfx())},
            follow_redirects=False,
        )
        assert guest.status_code == 303
        assert guest.headers["location"].startswith("/login")

        client.cookies.set(SESSION_COOKIE, first_session)
        imported = client.post(
            "/grand/import",
            files={"estimate": ("sample.gsfx", _gsfx())},
        )
        assert imported.status_code == 200
        token = next(iter(client.app.state.app_state.grand_drafts))
        assert client.get("/grand/preview", params={"token": token}).status_code == 200

        client.cookies.set(SESSION_COOKIE, second_session)
        assert client.get("/grand/preview", params={"token": token}).status_code == 403
        attempt = client.post(
            "/grand/preview",
            data={"token": token, "action": "xlsx"},
        )
        assert attempt.status_code == 403

        client.cookies.set(SESSION_COOKIE, first_session)
        assert client.get("/grand/preview", params={"token": token}).status_code == 200
