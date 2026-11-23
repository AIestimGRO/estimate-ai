"""Optional source-LSR upload never trains from unverified specialist data."""
from io import BytesIO
import json
from pathlib import Path
from zipfile import ZipFile

from fastapi.testclient import TestClient
from openpyxl import Workbook

from app.services.lsr_evidence import evaluate_lsr_pair
from app.web.app import create_app
from core.storage.accounts import ROLE_SPECIALIST, create_user
from core.storage.connection import connect, init_database


CODE_A = 'ГЭСН01-01-009-01'
CODE_B = 'ГЭСН01-01-009-02'


def _workbook_bytes(book):
    data = BytesIO()
    book.save(data)
    book.close()
    return data.getvalue()


def os_bytes(*, wrong_quantity=False, wrong_price=False, wrong_code=False):
    wb = Workbook()
    ws = wb.active
    ws.title = 'ОС ГРАНД'
    for col, title in ((2, '№ п/п в смете'), (3, 'Наименование работ'),
                       (4, 'Единица измерения'), (5, 'Кол-во'),
                       (6, 'Цена единицы работ'), (14, 'Перечень ГЭСН/ФЕР/ТЕР'),
                       (15, '/КР')):
        ws.cell(27, col, title)
    ws['B29'] = '162'
    ws['C29'] = 'Разработка грунта'
    ws['D29'] = 'м3'
    ws['E29'] = 613 if wrong_quantity else 612.5
    ws['F29'] = 120 if wrong_price else 100
    ws['N29'] = CODE_B if wrong_code else CODE_A
    ws['O29'] = ws['N29'].value
    ws['B30'] = '163'
    ws['C30'] = 'Погрузка в автотранспортное средство'
    ws['D30'] = 'т'
    ws['E30'] = 10
    ws['F30'] = 50
    ws['N30'] = '43-1'
    ws['O30'] = '43-1'
    return _workbook_bytes(wb)


def gsfx_bytes(*, different_code=False, no_direct_price=False):
    code = CODE_B if different_code else CODE_A
    price = '' if no_direct_price else '<PriceCurr MT="50"/>'
    xml = f'''<Document Generator="GrandSmeta">
      <Properties LocNum="TEST-1"/>
      <Chapters><Chapter Caption="Раздел">
        <Position Number="162" Code="{code}" Caption="Разработка грунта" Units="1000 м3">
          <Quantity Result="0,6125"/><PriceCurr MT="100000"/>
        </Position>
        <Position Number="163" Code="43-1" Caption="Погрузка в автотранспортное средство" Units="1т груза">
          <Quantity Result="10"/>{price}
        </Position>
      </Chapter></Chapters></Document>'''
    data = BytesIO()
    with ZipFile(data, 'w') as zipf:
        zipf.writestr('Properties.txt', 'GSFX')
        zipf.writestr('Data.xml', xml)
    return data.getvalue()


def test_pair_checks_code_quantity_unit_price_and_missing_price(tmp_path):
    gsfx = tmp_path / 'source.gsfx'
    os = tmp_path / 'estimate.xlsx'
    gsfx.write_bytes(gsfx_bytes())
    os.write_bytes(os_bytes())
    report = evaluate_lsr_pair(os, gsfx, sheet_title='ОС ГРАНД')
    assert report.quality == 'candidate'
    assert report.matched == 2
    assert report.checked_prices == 2
    assert not report.issues
    for opts, expected in (
        ({'wrong_quantity': True}, 'физическое количество'),
        ({'wrong_price': True}, 'Цена × количество'),
        ({'wrong_code': True}, 'однозначной позиции'),
    ):
        os.write_bytes(os_bytes(**opts))
        bad = evaluate_lsr_pair(os, gsfx, sheet_title='ОС ГРАНД')
        assert bad.quality == 'review'
        assert expected in ' '.join(bad.issues)
    os.write_bytes(os_bytes())
    gsfx.write_bytes(gsfx_bytes(no_direct_price=True))
    unchecked = evaluate_lsr_pair(os, gsfx, sheet_title='ОС ГРАНД')
    assert unchecked.quality == 'review'  # unverified prices cannot pass the gate
    assert unchecked.checked_prices == 1


def test_lsr_evidence_is_quarantined_until_admin_approval(tmp_path, monkeypatch):
    db_path = tmp_path / 'isolated.sqlite'
    monkeypatch.setenv('ESTIMATE_AI_DB_PATH', str(db_path))
    with TestClient(create_app(base_dir=tmp_path / 'work')) as client:
        assert 'name="lsr"' in client.get('/').text
        setup = client.post('/setup', data={
            'full_name': 'Admin', 'login': 'admin', 'password': 'verylongpassword',
        }, follow_redirects=False)
        assert setup.status_code == 303
        with connect(db_path) as connection:
            specialist_id = create_user(connection, full_name='Specialist',
                         login='specialist', password='verylongpassword', role=ROLE_SPECIALIST)
        login = client.post('/login', data={
            'login': 'specialist', 'password': 'verylongpassword',
        }, follow_redirects=False)
        assert login.status_code == 303
        catalog = Workbook()
        catalog.active.title = 'Каталог'
        catalog.active['B4'] = 'T-1'
        catalog.active['C4'] = 'Разработка грунта'
        catalog.active['D4'] = 'м3'
        catalog.active['G4'] = 100
        catalog.active['N4'] = CODE_A
        uploaded_os = os_bytes()
        uploaded_gsfx = gsfx_bytes()
        started = client.post('/run', files={
            'catalog': ('catalog.xlsx', _workbook_bytes(catalog)),
            'estimate': ('os.xlsx', uploaded_os),
            'lsr': ('source.gsfx', uploaded_gsfx),
        })
        assert started.status_code == 200
        assert 'Структурная сверка пройдена' in started.text
        import re
        token = re.search(r'name="token" value="([a-f0-9]+)"', started.text).group(1)
        sheet = re.search(r'name="sheet" value="([^"]+)"', started.text).group(1)
        confirmed = client.post('/confirm', data={
            'token': token, 'sheet': sheet, 'region': '', 'coefficient': '1',
        }, follow_redirects=False)
        assert confirmed.status_code == 303
        assert 'ЛСР № 1: сохранена отдельно' in client.get(f'/jobs/{token}').text
        with connect(db_path) as connection:
            cases = connection.execute('SELECT * FROM lsr_learning_cases').fetchall()
            assert len(cases) == 1
            row = cases[0]
            assert row['author_user_id'] == specialist_id
            assert row['quality'] == 'candidate'
            assert row['status'] == 'pending'
            assert bytes(row['gsfx_blob']) == uploaded_gsfx
            assert bytes(row['os_blob']) == uploaded_os
            assert len(json.loads(row['os_rows_json'])) == 2
            assert connection.execute('PRAGMA integrity_check').fetchone()[0] == 'ok'
        # A specialist cannot view or approve the private evidence queue.
        assert client.get('/admin/lsr-learning', follow_redirects=False).status_code == 403
        assert client.get('/admin/lsr-learning/export', follow_redirects=False).status_code == 403
        client.post('/login', data={'login': 'admin', 'password': 'verylongpassword'})
        response = client.get('/admin/lsr-learning')
        assert response.status_code == 200
        assert 'os.xlsx' in response.text
        assert client.get('/admin/lsr-learning/export').status_code == 422  # pending cases excluded
        approved = client.post('/admin/lsr-learning/1/review', data={'action': 'approved'},
                               follow_redirects=False)
        assert approved.status_code == 303
        assert client.get('/admin/lsr-learning/1/download-os').content == uploaded_os
        downloaded = client.get('/admin/lsr-learning/export')
        assert downloaded.status_code == 200
        assert downloaded.headers['content-type'] == 'application/zip'
        with ZipFile(BytesIO(downloaded.content)) as archive:
            assert archive.testzip() is None
            assert archive.read('pair_000001/original_lsr.gsfx') == uploaded_gsfx
            assert archive.read('pair_000001/original_os.xlsx') == uploaded_os
            metadata = json.loads(archive.read('pair_000001/verification.json'))
            manifest = json.loads(archive.read('manifest.json'))
        assert metadata['checked_prices'] == 2
        assert metadata['os_filename'] == 'os.xlsx'
        assert manifest['total_pairs'] == 1
        assert manifest['pairs'][0]['review_status'] == 'approved_by_admin'
        with connect(db_path) as connection:
            assert connection.execute('SELECT status FROM lsr_learning_cases WHERE id=1').fetchone()[0] == 'approved'
            connection.execute('UPDATE lsr_learning_cases SET os_blob=NULL WHERE id=1')
            connection.commit()
        # Replaying a pairing cannot change the administrator's decision.
        with connect(db_path) as connection:
            from core.storage.lsr_learning import save_learning_case
            record = client.app.state.app_state.store[token]
            case_id, created = save_learning_case(connection, gsfx_path=record.lsr_path,
                         os_path=record.estimate_path, gsfx_filename='source.gsfx',
                         os_filename='os.xlsx', sheet_title='ОС ГРАНД',
                         author_user_id=specialist_id, job_token=token,
                         report=record.evidence_report)
            assert (case_id, created) == (1, False)
            assert connection.execute('SELECT status FROM lsr_learning_cases WHERE id=1').fetchone()[0] == 'approved'
            assert bytes(connection.execute('SELECT os_blob FROM lsr_learning_cases WHERE id=1').fetchone()[0]) == uploaded_os
            connection.execute("UPDATE lsr_learning_cases SET os_blob = ? WHERE id=1", (b'tampered',))
            connection.commit()
        assert client.get('/admin/lsr-learning/export').status_code == 422  # hash mismatch blocks export


def test_unrelated_files_cannot_be_approved(tmp_path, monkeypatch):
    db = tmp_path / 'isolated.sqlite'
    monkeypatch.setenv('ESTIMATE_AI_DB_PATH', str(db))
    with TestClient(create_app(base_dir=tmp_path / 'sessions')) as client:
        client.post('/setup', data={'full_name': 'Admin', 'login': 'admin',
                                    'password': 'verylongpassword'})
        client.post('/login', data={'login': 'admin', 'password': 'verylongpassword'})
        gsfx = tmp_path / 'source.gsfx'
        os = tmp_path / 'os.xlsx'
        gsfx.write_bytes(gsfx_bytes())
        os.write_bytes(os_bytes(wrong_quantity=True))
        report = evaluate_lsr_pair(os, gsfx, sheet_title='ОС ГРАНД')
        assert report.quality == 'review'
        with connect(db) as connection:
            init_database(connection)
            from core.storage.lsr_learning import save_learning_case
            case_id, created = save_learning_case(connection,
                gsfx_path=gsfx, os_path=os, gsfx_filename='source.gsfx',
                os_filename='os.xlsx', sheet_title='ОС ГРАНД',
                author_user_id=1, job_token='test', report=report)
            assert created
        assert client.post(f'/admin/lsr-learning/{case_id}/review',
                           data={'action': 'approved'}).status_code == 409
        assert client.post(f'/admin/lsr-learning/{case_id}/review',
                           data={'action': 'rejected'}, follow_redirects=False).status_code == 303


def test_schema_migration_adds_learning_without_touching_existing_catalog():
    connection = connect(':memory:')
    try:
        init_database(connection, synchronize_corrections=False)
        with connection:
            connection.execute("INSERT INTO catalog_sources(name) VALUES ('main')")
            connection.execute('DROP TABLE lsr_learning_cases')
            connection.execute('DELETE FROM schema_migrations')
            connection.execute('INSERT INTO schema_migrations(version) VALUES (15)')
        init_database(connection, synchronize_corrections=False)
        assert connection.execute('SELECT max(version) FROM schema_migrations').fetchone()[0] == 17
        assert connection.execute("SELECT COUNT(*) FROM catalog_sources WHERE name='main'").fetchone()[0] == 1
        assert connection.execute('SELECT COUNT(*) FROM lsr_learning_cases').fetchone()[0] == 0
        assert connection.execute('PRAGMA integrity_check').fetchone()[0] == 'ok'
        assert connection.execute('PRAGMA foreign_key_check').fetchall() == []
    finally:
        connection.close()


def test_v16_migration_preserves_existing_case_table_and_adds_original_os():
    connection = connect(':memory:')
    try:
        init_database(connection, synchronize_corrections=False)
        with connection:
            connection.execute('ALTER TABLE lsr_learning_cases DROP COLUMN os_blob')
            connection.execute('DELETE FROM schema_migrations')
            connection.execute('INSERT INTO schema_migrations(version) VALUES (16)')
        init_database(connection, synchronize_corrections=False)
        cols = {row['name'] for row in connection.execute('PRAGMA table_info(lsr_learning_cases)')}
        assert 'os_blob' in cols
        assert connection.execute('SELECT MAX(version) FROM schema_migrations').fetchone()[0] == 17
        assert connection.execute('PRAGMA integrity_check').fetchone()[0] == 'ok'
    finally:
        connection.close()
