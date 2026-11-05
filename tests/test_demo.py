import importlib
import io
import json
import pytest
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setenv('DEMO_DB_PATH', str(tmp_path / 'test.sqlite3'))
    from backend import main
    importlib.reload(main)
    return TestClient(main.app)


def setup_run(client):
    sample = client.post('/api/sample').json()
    a, b = sample['old'], sample['new']
    payload = {'old_id': a['id'], 'new_id': b['id'], 'old_key': '员工编号', 'new_key': '人员编码',
               'fields': [{'old': h, 'new': h} for h in a['headers'][1:]]}
    r = client.post('/api/compare', json=payload)
    assert r.status_code == 200, r.text
    return r.json(), payload


def test_sample_review_export_persistence(client):
    run, payload = setup_run(client)
    assert run['counts'] == {'新增': 1, '缺失': 1, '重复': 1, '变更': 2, '空主键': 1}
    assert run['unchanged'] == 2
    changed = next(i for i in run['items'] if i['key'] == 'D002')
    assert changed['changes'][0]['old'] == '12000'
    assert changed['changes'][0]['new'] == '15000'
    assert changed['old'][0]['row'] == 3
    duplicate = next(i for i in run['items'] if i['kind'] == '重复')
    assert len(duplicate['new']) == 2 and duplicate['changes'] == []
    url = f'/api/runs/{run["id"]}'
    review = {'status': '已核实', 'reviewer': '演示员', 'note': '=1+1'}
    assert client.put(url + f'/items/{changed["id"]}/review', json=review).status_code == 200
    reopened = client.get(url).json()
    assert next(i for i in reopened['items'] if i['id'] == changed['id'])['review']['note'] == '=1+1'
    exported = client.get(url + '/export')
    assert exported.status_code == 200
    wb = load_workbook(io.BytesIO(exported.content), data_only=False)
    assert wb.sheetnames == ['差异与复核', '批次说明', '操作日志']
    formula_note = [c for row in wb.worksheets[0] for c in row if c.value == '=1+1']
    assert formula_note and all(c.data_type == 's' for c in formula_note)
    actions = [a['action'] for a in client.get('/api/audit').json()]
    assert '人工复核' in actions and '导出报告' in actions
    # Repeating comparison does not mutate imported records or inherit review status.
    again = client.post('/api/compare', json=payload).json()
    assert again['counts'] == run['counts']
    assert all(i['review']['status'] == '待复核' for i in again['items'])


def test_csv_and_xlsx_mapping(client):
    csv = 'id,old\n001,NA\n002,100\n'
    a = client.post('/api/upload', files={'file': ('old.csv', csv.encode(), 'text/csv')}).json()
    wb = Workbook()
    wb.active.append(['code', 'new'])
    wb.active.append(['001', 'NA'])
    wb.active.append(['002', '200'])
    stream = io.BytesIO()
    wb.save(stream)
    b = client.post('/api/upload', files={'file': ('new.xlsx', stream.getvalue())}).json()
    assert a['preview'][0]['values']['id'] == '001'
    assert a['preview'][0]['values']['old'] == 'NA'
    r = client.post('/api/compare', json={'old_id': a['id'], 'new_id': b['id'], 'old_key': 'id', 'new_key': 'code', 'fields': [{'old': 'old', 'new': 'new'}]})
    assert r.status_code == 200
    assert r.json()['counts'] == {'变更': 1} and r.json()['unchanged'] == 1
    assert r.json()['items'][0]['new'][0]['row'] == 3


@pytest.mark.parametrize('name,content', [('bad.txt', b'abc'), ('empty.csv', b''), ('duplicate.csv', b'id,id\n1,2'), ('ragged.csv', b'id,name\n1'), ('broken.xlsx', b'invalid'), ('big.csv', b'x' * (2*1024*1024+1)), ('empty-header.csv', b'id,\n1,2')])
def test_invalid_uploads(client, name, content):
    assert client.post('/api/upload', files={'file': (name, content)}).status_code == 400


def test_formula_rejected(client):
    wb = Workbook()
    wb.active.append(['id', 'value'])
    wb.active.append(['001', '=1+1'])
    stream = io.BytesIO()
    wb.save(stream)
    r = client.post('/api/upload', files={'file': ('formula.xlsx', stream.getvalue())})
    assert r.status_code == 400 and '公式' in r.json()['detail']


def test_validation(client):
    run, payload = setup_run(client)
    payload['fields'] *= 2
    assert client.post('/api/compare', json=payload).status_code == 400
    assert client.get('/api/runs/not-found').status_code == 404
    target = f'/api/runs/{run["id"]}/items/1/review'
    assert client.put(target, json={'status': '自动批准', 'reviewer': 'a', 'note': 'b'}).status_code == 422
    assert client.put(target, json={'status': '已核实', 'reviewer': ' ', 'note': ' '}).status_code == 400


def test_identical_and_duplicates_both_sides(client):
    a = client.post('/api/upload', files={'file': ('a.csv', b'id,v\n01,1\n02,2\n02,3')}).json()
    payload = {'old_id': a['id'], 'new_id': a['id'], 'old_key': 'id', 'new_key': 'id', 'fields': [{'old': 'v', 'new': 'v'}]}
    run = client.post('/api/compare', json=payload).json()
    assert run['unchanged'] == 1 and run['counts'] == {'重复': 1}
    assert len(run['items'][0]['old']) == len(run['items'][0]['new']) == 2
