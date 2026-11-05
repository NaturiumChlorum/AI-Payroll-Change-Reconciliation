"""Exercise the running demo with fictional sample CSV/XLSX over HTTP."""
import io
from pathlib import Path
import httpx
from openpyxl import load_workbook

root = Path(__file__).resolve().parents[1]
with httpx.Client(base_url='http://127.0.0.1:8000', trust_env=False, timeout=30) as client:
    assert client.get('/api/health').json()['status'] == 'ok'
    assert client.get('/').status_code == 200
    old = client.post('/api/upload', files={'file': ('sample-old.csv', (root / 'examples/sample-old.csv').read_bytes())}).json()
    new = client.post('/api/upload', files={'file': ('sample-new.xlsx', (root / 'examples/sample-new.xlsx').read_bytes())}).json()
    run = client.post('/api/compare', json={'old_id': old['id'], 'new_id': new['id'], 'old_key': '员工编号', 'new_key': '人员编码', 'fields': [{'old': h, 'new': h} for h in old['headers'][1:]]}).json()
    assert run['counts'] == {'新增': 1, '缺失': 1, '重复': 1, '变更': 2, '空主键': 1}
    assert run['unchanged'] == 2
    item = next(i for i in run['items'] if i['key'] == 'D002')
    url = f'/api/runs/{run["id"]}'
    result = client.put(url + f'/items/{item["id"]}/review', json={'status': '需补充资料', 'reviewer': '演示测试员', 'note': '课堂验证：12000 → 15000，需补充虚构变更通知与生效日期。'})
    assert result.status_code == 200
    saved = client.get(url).json()
    assert next(i for i in saved['items'] if i['key'] == 'D002')['review']['status'] == '需补充资料'
    report = client.get(url + '/export')
    assert report.status_code == 200
    wb = load_workbook(io.BytesIO(report.content))
    assert len(wb.sheetnames) == 3
    folder = root / 'docs'
    folder.mkdir(exist_ok=True)
    (folder / 'verified-sample-report.xlsx').write_bytes(report.content)
    assert any(log['action'] == '导出报告' for log in client.get('/api/audit').json())
    print('PASS: health, webpage, CSV+XLSX upload, mapping, comparison, review, reload, XLSX export, audit')
    print('Verified run:', run['id'])
