"""Generate fictional downloadable classroom fixtures without creating app data."""
import ast
import csv
import io
from pathlib import Path
from openpyxl import Workbook

root = Path(__file__).resolve().parents[1]
tree = ast.parse((root / 'backend/main.py').read_text(encoding='utf-8'))
samples = {n.targets[0].id: ast.literal_eval(n.value) for n in tree.body
           if isinstance(n, ast.Assign) and isinstance(n.targets[0], ast.Name)
           and n.targets[0].id in ('SAMPLE_OLD', 'SAMPLE_NEW')}
folder = root / 'examples'
folder.mkdir(exist_ok=True)
for side in ('old', 'new'):
    data = samples['SAMPLE_' + side.upper()]
    (folder / f'sample-{side}.csv').write_text(data, encoding='utf-8-sig', newline='')
    wb = Workbook()
    wb.active.title = '虚构演示资料'
    for row in csv.reader(io.StringIO(data)):
        wb.active.append(row)
    wb.save(folder / f'sample-{side}.xlsx')
