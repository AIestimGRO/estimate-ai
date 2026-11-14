"""Export GrandSmeta OS using the established 16-column RIM/OS structure.

Positioned fields: header row 27, numbering 28, work rows from 29.
N is the original code, O the independent /KR column, P the section code.
Catalog matching is run later by the regular matching engine.
"""
from __future__ import annotations
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

def _num(value):
    if value in (None, ''):
        return None
    try:
        parsed = Decimal(str(value).replace('\u00a0', '').replace(' ', '').replace(',', '.'))
        return parsed if parsed.is_finite() else None
    except (ValueError, InvalidOperation):
        return None

def _text(value: object) -> str:
    result = str(value or '')
    return "'" + result if result.lstrip().startswith(('=', '+', '-', '@')) else result

def export_grand_os(rows: list[dict], target: Path, *, metadata: dict | None = None) -> None:
    info = metadata or {}
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'ОС ГРАНД'
    title_fill = PatternFill('solid', fgColor='243B53')
    label_fill = PatternFill('solid', fgColor='DEEAF4')
    alert_fill = PatternFill('solid', fgColor='FFF0C2')
    sheet.sheet_view.showGridLines = False
    sheet['M2'] = 'УТВЕРЖДЕНО:'
    sheet['M3'] = 'Оценка стоимости (по заполненной форме РИМ)'
    task_no = str(info.get('task_number') or '').strip()
    sheet['M9'] = '№ задачи: ' + (task_no or 'не указан — заполните')
    if not task_no:
        sheet['M9'].fill = alert_fill
    sheet['B10'] = 'Исходные данные:'
    for nr, label in ((11,'РД'),(12,'КСГ'),(13,'ТЗ'),(14,'ВОР, нВОР'),(15,'договор ГП')):
        sheet.cell(nr, 2, nr-10)
        sheet.cell(nr, 3, label)
    sheet.merge_cells('B16:P16')
    sheet['B16'] = 'Оценка стоимости (сметно-коммерческим методом).\n' + _text(info.get('object_name') or 'Объект не указан')
    sheet['B16'].font = Font(size=12, bold=True)
    sheet['B16'].alignment = Alignment(wrap_text=True, vertical='center')
    sheet.row_dimensions[16].height = 43
    sheet.merge_cells('B17:P17')
    sheet['B17'] = 'Вид работ: ' + _text(info.get('work_type') or 'СМР')
    sheet.merge_cells('B18:P18')
    sheet['B18'] = 'Техническое задание: ' + _text(info.get('description') or '')
    sheet['B18'].alignment = Alignment(wrap_text=True)
    sheet.row_dimensions[18].height = 30
    for i, label in ((20,'Регион:'),(21,'Коэффициент:'),(22,'Сметы в ценах:'),
                     (23,'Начало работ:'),(24,'Окончание работ:')):
        sheet.cell(i, 3, label).font = Font(bold=True, color='243B53')
    sheet['D20'] = _text(info.get('region') or '')
    coef = _num(info.get('coefficient'))
    if coef is not None and coef > 0:
        sheet['D21'] = float(coef)
    sheet['D22'] = _text(info.get('price_basis') or '')
    sheet['D23'] = _text(info.get('start_date') or '')
    sheet['D24'] = _text(info.get('end_date') or '')
    sheet['C26'] = 'ИТОГО:'
    sheet['C26'].font = Font(bold=True)
    headers = [
        '№ п/п', '№ п/п в смете', 'Наименование работ',
        'Единица измерения', 'Кол-во', 'Цена единицы работ, руб. без НДС',
        'Цена средняя', 'Итого стоимость, руб. без НДС', 'Стоимость средняя',
        'ТЗ на ед., чел-час', 'ТЗм на ед., чел-час',
        'ТЗ всего, чел-час', 'ТЗм всего, чел-час',
        'Перечень ГЭСН/ФЕР/ТЕР', '/КР', 'Код раздела'
    ]
    for col, value in enumerate(headers, 1):
        c = sheet.cell(27, col, value)
        c.fill = title_fill
        c.font = Font(color='FFFFFF', bold=True, size=9)
        c.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        c.border = Border(bottom=Side(style='thin', color='CFD8E3'))
    sheet.row_dimensions[27].height = 56
    for col in range(1, 17):
        c = sheet.cell(28, col, str(col))
        c.fill = label_fill
        c.font = Font(color='526579', size=8)
        c.alignment = Alignment(horizontal='center')
    sheet.row_dimensions[28].height = 19
    trace = workbook.create_sheet('Сверка с GSFX')
    trace.append(['№ в смете','Шифр','Наименование GSFX','Ед. GSFX',
                  'Кол-во GSFX','Цена до обработки','Цена ОС','Источник цены',
                  'Статус','Комментарий','В ОС'])
    row_num = 28
    for row in rows:
        price, qty = _num(row.get('price')), _num(row.get('quantity'))
        total = (price*qty).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP) if price is not None and qty is not None else None
        if row.get('include', True):
            row_num += 1
            index = row_num
            values = [
                index-28, _text(row.get('number')),
                _text(row.get('name')), _text(row.get('unit')),
                float(qty) if qty is not None else None,
                float(price) if price is not None and price > 0 else None,
                None, float(total) if total is not None else None, None,
                *[float(_num(row.get(k))) if _num(row.get(k)) is not None else None for k in
                  ('labor_unit','machinist_labor_unit','labor_total','machinist_labor_total')],
                _text(row.get('code')), _text(row.get('code')),
                _text(row.get('section_code')),
            ]
            for col, value in enumerate(values, 1):
                c = sheet.cell(index, col, value)
                c.alignment = Alignment(vertical='top', wrap_text=col in (3,14,15))
                c.border = Border(bottom=Side(style='hair', color='CFD8E3'))
                if 5 <= col <= 13:
                    c.number_format = '#,##0.00'
            sheet.cell(index,9, f'=IF(G{index}="","",ROUND(E{index}*G{index},2))')
            if price is None:
                sheet.cell(index,6).fill = alert_fill
        trace.append([
            _text(row.get('number')), _text(row.get('code')),
            _text(row.get('original_name', row.get('name'))),
            _text(row.get('original_unit', row.get('unit'))),
            _text(row.get('original_quantity', row.get('quantity'))),
            _text(row.get('original_price')), float(price) if price is not None else None,
            _text(row.get('price_origin', 'GSFX: нет рассчитанной цены')),
            _text(row.get('normalization_status')), _text(row.get('normalization_note')),
            'ДА' if row.get('include', True) else 'НЕТ',
        ])
    last = max(29, row_num)
    for col in ('H','I','L','M'):
        c = sheet[f'{col}26']
        c.value = f'=SUM({col}29:{col}{last})'
        c.font = Font(bold=True)
        c.number_format = '#,##0.00'
    for col,width in {
      'A':9,'B':18,'C':65,'D':17,'E':15,'F':23,'G':19,'H':23,'I':19,
      'J':19,'K':19,'L':19,'M':19,'N':28,'O':28,'P':18
    }.items():
        sheet.column_dimensions[col].width = width
    sheet.freeze_panes = 'C29'
    sheet.auto_filter.ref = f'A27:P{last}'
    sheet.print_title_rows = '1:28'
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.fitToWidth = 1
    trace.freeze_panes = 'C2'
    trace.auto_filter.ref = f'A1:K{max(trace.max_row,2)}'
    for i,width in enumerate((18,27,64,19,19,19,19,38,20,65,10),1):
        trace.column_dimensions[get_column_letter(i)].width = width
    for c in trace[1]:
        c.fill = title_fill
        c.font = Font(bold=True,color='FFFFFF')
        c.alignment = Alignment(wrap_text=True)
    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = 'auto'
    except AttributeError:
        pass
    workbook.save(target)
    workbook.close()
