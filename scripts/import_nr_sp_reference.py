"""Normalize the published FSNB-2022 NR/SP attachment into reviewable JSON.

Usage: python scripts/import_nr_sp_reference.py published_reference.xlsx
The input is the publicly published recommended mapping for supplement 19.
"""
from __future__ import annotations

import json
from pathlib import Path
import sys

from openpyxl import load_workbook


def convert(source: Path, target: Path) -> None:
    wb = load_workbook(source, read_only=True, data_only=True)
    try:
        ws = wb.active
        rows = {}
        for line in ws.iter_rows(min_row=8, max_col=8, values_only=True):
            number, code, name, applicability, normal, northern_equal, far_north, profit = line
            if not code or not all(isinstance(rate, (int, float)) for rate in (normal, northern_equal, far_north, profit)):
                continue
            if str(code) in rows:
                raise ValueError(f'Duplicate NR/SP code: {code}')
            rows[str(code)] = {
                'work_type': str(name or ''),
                'applicability': str(applicability or ''),
                'overhead_ordinary': int(normal),
                'overhead_equivalent_north': int(northern_equal),
                'overhead_far_north': int(far_north),
                'profit': int(profit),
            }
        if len(rows) < 100:
            raise ValueError('The supplied reference does not contain enough normative work groups')
    finally:
        wb.close()
    result = {
        'edition': 'FSNB-2022 supplement 19; recommended NR/SP mapping, 24.08.2026',
        'provenance': 'https://fgiscs.minstroyrf.ru/frsn/reference/letters',
        'overhead_method': 'Minstroy 812/pr (2020, amendments 636/pr and 611/pr)',
        'profit_method': 'Minstroy 774/pr (2020, amendment 317/pr)',
        'groups': rows,
    }
    target.write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    if len(sys.argv) != 2:
        raise SystemExit('Usage: python scripts/import_nr_sp_reference.py reference.xlsx')
    convert(Path(sys.argv[1]), Path(__file__).resolve().parents[1] / 'core/data/nr_sp_fsnb2022_dop19.json')
