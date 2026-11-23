"""Single-file RIM regression cases using synthetic, non-client GSFX data."""
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from app.services.grand_rim_calculator import calculate_gsfx_prices, _group_for_code, _reference
from app.services.grand_normalization import (
    CatalogExample, CatalogIndex, normalize_rows, parse_unit,
)
from app.web.app import create_app


def make_estimate(tmp_path: Path, *, region='14', zone='4', addon='19',
                  resource_price='2728,89', operator_price='1318') -> Path:
    xml = f'''<Document Generator="GrandSmeta">
      <RegionInfo RegionName="ФСНБ-2022 (с Изм. 1-19)" AdmRegionCode="{region}"
                  AdmRegionZoneCode="{zone}" AdmRegionZone="Республика Саха (Якутия) (4 зона)"/>
      <FRSN_Info AddOnNumber="{addon}"/>
      <Chapters><Chapter Caption="Раздел">
        <Position Number="162" Code="ГЭСН01-01-009-01" Caption="Разработка грунта"
                  Units="1000 м3">
          <Quantity Result="0,6125"/>
          <Resources>
            <Tzm Quantity="13"/>
            <Mch Quantity="13" Code="91.01.05-087">
              <PriceCurr Value="{resource_price}"/>
              <Tzm2022 Quantity="1" Price="{operator_price}"/>
            </Mch>
          </Resources>
          <Koefficients>
            <K Options="Base Curr EmAll EmQty OzpTz ZpmTz" Value_OZ="1,7" Value_EM="1,7"/>
            <K Options="Base Curr EmAll EmQty OzpTz ZpmTz" Value_OZ="1,1" Value_EM="1,1"/>
          </Koefficients>
        </Position>
      </Chapter></Chapters>
    </Document>'''
    path = tmp_path / 'estimate.gsfx'
    with ZipFile(path, 'w') as archive:
        archive.writestr('Properties.txt', 'GSFX')
        archive.writestr('Data.xml', xml)
    return path


def test_rim_calculates_normative_wages_and_coefficients_without_xlsx(tmp_path):
    results = calculate_gsfx_prices(make_estimate(tmp_path))
    assert len(results) == 1
    result = results[0]
    assert result is not None
    assert result.nr_sp_code == '001.1'
    # RKS: 97% overhead + 46% profit; position coefficients 1.7 × 1.1.
    assert result.total == Decimal('88321.23')
    assert result.unit_price == Decimal('88321.23') / Decimal('0.6125')
    assert result.machinist_labor_per_norm_unit == Decimal('24.31')


def test_rim_does_not_guess_region_norm_revision_or_machinist_salary(tmp_path):
    assert calculate_gsfx_prices(make_estimate(tmp_path, zone='9')) == (None,)
    assert calculate_gsfx_prices(make_estimate(tmp_path, addon='18')) == (None,)
    assert calculate_gsfx_prices(make_estimate(tmp_path, operator_price='')) == (None,)
    assert calculate_gsfx_prices(make_estimate(tmp_path, resource_price='')) == (None,)


@pytest.mark.parametrize('code, group', [
    ('ГЭСН01-01-009-01', '001.1'),
    ('ГЭСН22-01-024-13', '018.0'),
    ('ГЭСН44-01-026-02', '038.0'),
    ('ГЭСН46-03-006-01', '040.1'),
    ('ГЭСН01-02-055-01', None),
    ('ГЭСНм08-01-080-01', None),
])
def test_nr_sp_reference_maps_only_unambiguous_work_categories(code, group):
    assert _group_for_code(code, _reference()) == group


def test_rim_rejects_unverified_estimate_wide_coefficients(tmp_path):
    path = make_estimate(tmp_path)
    with ZipFile(path) as archive:
        data = archive.read('Data.xml').decode('utf-8')
    data = data.replace('<Chapters>',
        '<Koefficients><K Options="Base Curr OzpTz" Value_OZ="1,25"/></Koefficients><Chapters>')
    with ZipFile(path, 'w') as archive:
        archive.writestr('Properties.txt', 'GSFX')
        archive.writestr('Data.xml', data)
    assert calculate_gsfx_prices(path) == (None,)


def test_labor_hours_follow_the_same_quantity_scaling_as_work_price():
    source = parse_unit('1000 м3')
    target = parse_unit('м3')
    assert source is not None and target is not None
    index = CatalogIndex({'ГЭСН01-01-009-01': (
        CatalogExample('Разработка грунта', 'м3', target),
    )}, 1)
    row = {'code': 'ГЭСН01-01-009-01', 'name': 'Разработка грунта',
           'unit': '1000 м3', 'quantity': '0.6125', 'price': '144197.93',
           'labor_unit': '0', 'machinist_labor_unit': '24.31',
           'machinist_labor_total': '14.89'}
    summary = normalize_rows([row], index)
    assert summary.changed == 1
    assert row['unit'] == 'м3'
    assert row['quantity'] == '612.5'
    assert row['price'] == '144.19793'
    assert row['machinist_labor_unit'] == '0.02431'
    assert row['machinist_labor_total'] == '14.89'


def test_single_gsfx_upload_auto_prices_and_exports_valid_os(tmp_path, monkeypatch):
    monkeypatch.setenv('ESTIMATE_AI_DB_PATH', str(tmp_path / 'isolated_auth.sqlite'))
    source = make_estimate(tmp_path)
    with TestClient(create_app(base_dir=tmp_path / 'uploads')) as client:
        response = client.post('/grand/import', files={
            'estimate': ('estimate.gsfx', source.read_bytes()),
        })
        assert response.status_code == 200
        token = next(iter(client.app.state.app_state.grand_drafts))
        rows = client.app.state.app_state.grand_drafts[token]['rows']
        assert len(rows) == 1
        assert rows[0]['price'] == format(Decimal('88321.23') / Decimal('0.6125'), 'f')
        assert 'Расчёт РИМ из GSFX' in rows[0]['price_origin']
        assert rows[0]['machinist_labor_unit'] == '24.31'
        exported = client.post('/grand/preview', data={
            'token': token, 'include_0': 'on', 'action': 'xlsx',
        })
        assert exported.status_code == 200
        with ZipFile(BytesIO(exported.content)) as archive:
            assert archive.testzip() is None
        workbook = load_workbook(BytesIO(exported.content), read_only=True, data_only=True)
        try:
            sheet = workbook.active
            assert sheet['B29'].value == '162'
            assert sheet['F29'].value is not None
            assert sheet['H29'].value == 88321.23
            assert sheet['K29'].value == 24.31
            assert sheet['N27'].value == 'Перечень ГЭСН/ФЕР/ТЕР'
            assert sheet['O27'].value == '/КР'
        finally:
            workbook.close()
