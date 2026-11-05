"""Local classroom demo. Source records are immutable; reviews are separate."""
import csv
import io
import json
import os
import sqlite3
import uuid
import zipfile
from collections import Counter, defaultdict
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Literal

import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook, load_workbook
from pydantic import BaseModel, Field

ROOT = Path(__file__).resolve().parents[1]
DB = Path(os.environ.get('DEMO_DB_PATH', str(ROOT / 'data/demo.sqlite3')))
DB.parent.mkdir(parents=True, exist_ok=True)
app = FastAPI(title='AI辅助薪酬服务变更核对系统 · 课堂演示', version='1.0.0')


@contextmanager
def connection():
    c = sqlite3.connect(DB, timeout=15)
    c.row_factory = sqlite3.Row
    try:
        with c:
            yield c
    finally:
        c.close()


with connection() as c:
    c.executescript('''
    CREATE TABLE IF NOT EXISTS uploads(id TEXT PRIMARY KEY, body TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS runs(id TEXT PRIMARY KEY, created TEXT, body TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS reviews(run_id TEXT, item_id TEXT, status TEXT, reviewer TEXT, note TEXT,
       updated TEXT, PRIMARY KEY(run_id,item_id));
    CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY AUTOINCREMENT, time TEXT, action TEXT,
       run_id TEXT, detail TEXT);
    ''')


def now():
    return datetime.now(timezone.utc).isoformat()


def audit(c, action, run_id, detail):
    c.execute('INSERT INTO audit(time,action,run_id,detail) VALUES(?,?,?,?)',
              (now(), action, run_id, json.dumps(detail, ensure_ascii=False)))


def fail(message):
    raise HTTPException(400, message)


def parse_file(name, raw):
    ext = Path(name).suffix.lower()
    if ext not in ('.csv', '.xlsx'):
        fail('仅支持 CSV 或 XLSX 文件。')
    if not raw or len(raw) > 2 * 1024 * 1024:
        fail('文件为空或超过 2 MB。')
    sheet = 'CSV'
    try:
        if ext == '.csv':
            try:
                content = raw.decode('utf-8-sig')
            except UnicodeDecodeError:
                content = raw.decode('gb18030')
            rows = list(csv.reader(io.StringIO(content, newline=''), strict=True))
        else:
            with zipfile.ZipFile(io.BytesIO(raw)) as z:
                if sum(x.file_size for x in z.infolist()) > 20 * 1024 * 1024:
                    fail('XLSX 解压后过大，请缩小文件。')
            wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=False)
            ws = wb.worksheets[0]
            sheet = ws.title
            rows = []
            for row in ws.iter_rows():
                if len(rows) > 5000 or len(row) > 50:
                    fail('最多支持 5000 条记录、50 列。')
                if any(cell.data_type == 'f' for cell in row):
                    fail('演示版不计算公式，请将公式粘贴为值后上传。')
                rows.append(['' if cell.value is None else str(cell.value) for cell in row])
            wb.close()
        if len(rows) < 2:
            fail('需要首行表头及至少一条数据。')
        headers = [str(x).strip() for x in rows[0]]
        if not headers or len(headers) > 50 or len(rows) > 5001:
            fail('最多支持 5000 条记录、50 列。')
        if any(not x for x in headers) or len(set(headers)) != len(headers):
            fail('表头不能为空或重复。')
        if any(len(r) != len(headers) for r in rows[1:]):
            fail('存在列数与表头不一致的记录。')
        if any(len(str(v)) > 2000 for r in rows for v in r):
            fail('单元格不能超过 2000 字符。')
        # Explicit string semantics: preserve leading zeroes and NA literal.
        frame = pd.DataFrame(rows[1:], columns=headers, dtype=str).fillna('')
        records = [{'row': i + 2, 'values': r} for i, r in enumerate(frame.to_dict('records'))]
        return {'id': uuid.uuid4().hex, 'name': Path(name.replace('\\', '/')).name,
                'sheet': sheet, 'headers': headers, 'records': records, 'count': len(records)}
    except HTTPException:
        raise
    except Exception:
        fail('无法读取文件，请检查编码、文件格式及内容。CSV 支持 UTF-8 / GB18030。')


def store_upload(c, obj):
    c.execute('INSERT INTO uploads VALUES(?,?)', (obj['id'], json.dumps(obj, ensure_ascii=False)))
    audit(c, '导入资料', '', {'file': obj['name'], 'records': obj['count'], 'id': obj['id']})


def public_upload(obj):
    return {**{k: v for k, v in obj.items() if k != 'records'}, 'preview': obj['records'][:5]}


@app.get('/api/health')
def health():
    return {'status': 'ok', 'assistance': 'rule-based-demo'}


@app.post('/api/upload')
async def upload(file: UploadFile = File(...)):
    raw = await file.read(2 * 1024 * 1024 + 1)
    await file.close()
    obj = parse_file(file.filename or '', raw)
    with connection() as c:
        store_upload(c, obj)
    return public_upload(obj)


SAMPLE_OLD = '员工编号,姓名,部门,月度金额,服务状态\nD001,演示甲,运营组,8000,在服\nD002,演示乙,技术组,12000,在服\nD003,演示丙,运营组,9000,在服\nD004,演示丁,客服组,7000,在服\nD005,演示戊,财务组,10000,在服\nD006,演示己,客服组,7500,在服\n'
SAMPLE_NEW = '人员编码,姓名,部门,月度金额,服务状态\nD001,演示甲,运营组,8000,在服\nD002,演示乙,技术组,15000,在服\nD003,演示丙,项目组,9000,暂停\nD005,演示戊,财务组,10000,在服\nD005,演示戊,财务组,10500,在服\nD006,演示己,客服组,7500,在服\nD007,演示庚,运营组,8500,在服\n,演示待核,运营组,6000,在服\n'


@app.post('/api/sample')
def sample():
    a = parse_file('虚构样例_变更前.csv', SAMPLE_OLD.encode())
    b = parse_file('虚构样例_变更后.csv', SAMPLE_NEW.encode())
    with connection() as c:
        store_upload(c, a)
        store_upload(c, b)
    return {'old': public_upload(a), 'new': public_upload(b)}


@app.get('/api/sample/{side}')
def sample_download(side: Literal['old', 'new']):
    data = SAMPLE_OLD if side == 'old' else SAMPLE_NEW
    return StreamingResponse(io.BytesIO(data.encode('utf-8-sig')), media_type='text/csv',
                             headers={'Content-Disposition': f'attachment; filename="sample-{side}.csv"'})


class Mapping(BaseModel):
    old: str
    new: str


class Compare(BaseModel):
    old_id: str
    new_id: str
    old_key: str
    new_key: str
    fields: list[Mapping] = Field(min_length=1, max_length=50)


def get_upload(c, id):
    row = c.execute('SELECT body FROM uploads WHERE id=?', (id,)).fetchone()
    if not row:
        raise HTTPException(404, '导入资料不存在，请重新导入。')
    return json.loads(row['body'])


def reconcile(a, b, req):
    groups = []
    items = []

    def item(kind, key, old, new, changes, explanation, suggestion):
        items.append({'id': str(len(items) + 1), 'kind': kind, 'key': key,
                      'old': old, 'new': new, 'changes': changes,
                      'explanation': explanation, 'suggestion': suggestion})

    for source, keyfield, side in [(a, req.old_key, 'old'), (b, req.new_key, 'new')]:
        grouped = defaultdict(list)
        for record in source['records']:
            key = record['values'][keyfield].strip()
            if not key:
                item('空主键', f'空主键·{side}·{record["row"]}', [record] if side == 'old' else [],
                     [record] if side == 'new' else [], [], '记录缺少匹配标识，无法确定对应关系。',
                     '核实员工编号并在原始资料中补充；本系统不修改源文件。')
            else:
                grouped[key].append(record)
        groups.append(grouped)
    oldg, newg = groups
    unchanged = 0
    for key in sorted(set(oldg) | set(newg)):
        old, new = oldg[key], newg[key]
        if len(old) > 1 or len(new) > 1:
            item('重复', key, old, new, [], f'变更前 {len(old)} 条、变更后 {len(new)} 条，主键不能唯一匹配。',
                 '逐行核对重复原因；未选择任意一条进行金额比较，需先确认唯一记录。')
        elif not old:
            item('新增', key, [], new, [], '该主键仅出现在变更后资料。', '核实新增服务人员及生效资料，确认是否属于本期范围。')
        elif not new:
            item('缺失', key, old, [], [], '该主键仅出现在变更前资料。', '核实是否退出服务、漏报或版本不完整，不能直接判定为离职。')
        else:
            changes = [{'field': f.old, 'new_field': f.new, 'old': old[0]['values'][f.old],
                        'new': new[0]['values'][f.new]} for f in req.fields
                       if old[0]['values'][f.old] != new[0]['values'][f.new]]
            if changes:
                item('变更', key, old, new, changes, f'{len(changes)} 个映射字段的文本值发生变化。',
                     '对照变更通知及生效日期逐项核实；差异本身不代表错误或已获批准。')
            else:
                unchanged += 1
    return items, unchanged


@app.post('/api/compare')
def compare(req: Compare):
    with connection() as c:
        a, b = get_upload(c, req.old_id), get_upload(c, req.new_id)
        if req.old_key not in a['headers'] or req.new_key not in b['headers']:
            fail('主键不在选定文件的表头中。')
        if len({f.old for f in req.fields}) != len(req.fields) or len({f.new for f in req.fields}) != len(req.fields):
            fail('字段映射必须一一对应，不能重复。')
        if any(f.old not in a['headers'] or f.new not in b['headers'] for f in req.fields):
            fail('字段映射包含不存在的列。')
        items, unchanged = reconcile(a, b, req)
        run = {'id': uuid.uuid4().hex, 'created': now(), 'old': public_upload(a), 'new': public_upload(b),
               'mapping': req.model_dump(), 'items': items, 'unchanged': unchanged,
               'counts': dict(Counter(i['kind'] for i in items)), 'assistance': '规则辅助（演示），未调用大模型'}
        c.execute('INSERT INTO runs VALUES(?,?,?)', (run['id'], run['created'], json.dumps(run, ensure_ascii=False)))
        audit(c, '完成只读核对', run['id'], {'mapping': req.model_dump(), 'counts': run['counts']})
    return get_run(run['id'])


@app.get('/api/runs')
def runs():
    with connection() as c:
        return [dict(r) for r in c.execute('SELECT id,created FROM runs ORDER BY created DESC LIMIT 50')]


@app.get('/api/runs/{run_id}')
def get_run(run_id: str):
    with connection() as c:
        row = c.execute('SELECT body FROM runs WHERE id=?', (run_id,)).fetchone()
        if not row:
            raise HTTPException(404, '核对批次不存在。')
        run = json.loads(row['body'])
        reviews = {r['item_id']: dict(r) for r in c.execute('SELECT * FROM reviews WHERE run_id=?', (run_id,))}
    for item in run['items']:
        item['review'] = reviews.get(item['id'], {'status': '待复核', 'reviewer': '', 'note': '', 'updated': ''})
    return run


class Review(BaseModel):
    status: Literal['待复核', '已核实', '需补充资料', '存在疑问']
    reviewer: str = Field(min_length=1, max_length=50)
    note: str = Field(min_length=1, max_length=1000)


@app.put('/api/runs/{run_id}/items/{item_id}/review')
def review(run_id: str, item_id: str, req: Review):
    run = get_run(run_id)
    item = next((i for i in run['items'] if i['id'] == item_id), None)
    if not item:
        raise HTTPException(404, '差异项不存在。')
    if not req.note.strip() or not req.reviewer.strip():
        fail('请填写复核人及复核备注。')
    with connection() as c:
        c.execute('INSERT OR REPLACE INTO reviews VALUES(?,?,?,?,?,?)',
                  (run_id, item_id, req.status, req.reviewer.strip(), req.note.strip(), now()))
        audit(c, '人工复核', run_id, {'item': item_id, 'before': item['review'], 'after': req.model_dump()})
    return get_run(run_id)


@app.get('/api/audit')
def logs(run_id: str = ''):
    with connection() as c:
        rows = c.execute('SELECT * FROM audit WHERE run_id=? ORDER BY id DESC LIMIT 200', (run_id,)) if run_id else c.execute('SELECT * FROM audit ORDER BY id DESC LIMIT 200')
        return [dict(r) for r in rows]


@app.get('/api/runs/{run_id}/export')
def export(run_id: str):
    run = get_run(run_id)
    wb = Workbook()
    ws = wb.active
    ws.title = '差异与复核'
    ws.append(['差异ID', '类型', '主键', '字段', '新文件字段', '旧值', '新值', '旧来源', '新来源', '规则解释', '核对建议', '复核状态', '复核人', '备注'])
    def source(side, records):
        return '; '.join(f'{run[side]["name"]} / {run[side]["sheet"]} / 第{r["row"]}行' for r in records)
    for i in run['items']:
        changes = i['changes'] or [{'field': '整条记录', 'new_field': '', 'old': json.dumps([r['values'] for r in i['old']], ensure_ascii=False), 'new': json.dumps([r['values'] for r in i['new']], ensure_ascii=False)}]
        for change in changes:
            ws.append([i['id'], i['kind'], i['key'], change['field'], change['new_field'], change['old'], change['new'], source('old', i['old']), source('new', i['new']), i['explanation'], i['suggestion'], i['review']['status'], i['review']['reviewer'], i['review']['note']])
    summary = wb.create_sheet('批次说明')
    for row in [('批次', run_id), ('时间 UTC', run['created']), ('辅助模式', run['assistance']), ('未变更主键数', run['unchanged']), ('差异项数', len(run['items'])), ('计数口径', '一项对应一个主键；空主键按源行单独计数；重复主键不再做字段比较。'), ('比较规则', '主键去首尾空格，字段按原始文本精确比较；仅首张工作表；CSV行号指逻辑记录序号（含表头）。'), ('范围', '仅课堂演示；人工复核状态不构成薪酬审批。'), ('映射', json.dumps(run['mapping'], ensure_ascii=False))]:
        summary.append(row)
    trace = wb.create_sheet('操作日志')
    trace.append(['时间 UTC', '操作', '详情'])
    with connection() as c:
        audit(c, '导出报告', run_id, {'format': 'xlsx'})
    for row in reversed(logs(run_id)):
        trace.append([row['time'], row['action'], row['detail']])
    from openpyxl.styles import Font, PatternFill, Alignment
    for sheet in wb:
        sheet.freeze_panes = 'A2'
        sheet.auto_filter.ref = sheet.dimensions
        for row in sheet:
            for cell in row:
                if isinstance(cell.value, str):
                    cell.data_type = 's'  # never execute user content as spreadsheet formulas
                cell.alignment = Alignment(vertical='top', wrap_text=True)
        for cell in sheet[1]:
            cell.font = Font(color='FFFFFF', bold=True)
            cell.fill = PatternFill('solid', fgColor='155E63')
        for col in sheet.columns:
            sheet.column_dimensions[col[0].column_letter].width = 26
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return StreamingResponse(stream, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', headers={'Content-Disposition': f'attachment; filename="reconciliation-{run_id[:8]}.xlsx"'})


DIST = ROOT / 'frontend/dist'
if DIST.exists():
    app.mount('/', StaticFiles(directory=DIST, html=True), name='frontend')
