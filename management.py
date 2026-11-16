"""Project-scoped resource and expense ledgers; quality uses live acceptance data."""
import datetime
import math
import os
import sqlite3
import json
import uuid
from contextlib import contextmanager

FIELDS = {
    'resources': ['name', 'source', 'category', 'quantity', 'unit', 'acquired', 'scope', 'use_start', 'use_end', 'owner', 'storage_location', 'note'],
    'costs': ['name', 'category', 'amount', 'currency', 'date', 'scope', 'owner', 'note'],
}

@contextmanager
def connect(root):
    os.makedirs(root, exist_ok=True)
    db = sqlite3.connect(os.path.join(root, '资源成本管理.db'))
    try:
        with db:
            db.execute('CREATE TABLE IF NOT EXISTS ledger (id TEXT PRIMARY KEY, kind TEXT NOT NULL, data TEXT NOT NULL)')
            yield db
    finally:
        db.close()

def snapshot(root):
    import salary_costs
    salary_sync = salary_costs.sync(root)
    with connect(root) as db:
        rows = db.execute('SELECT id, kind, data FROM ledger ORDER BY rowid DESC').fetchall()
    result = {kind: [dict(json.loads(data), id=id) for id, k, data in rows if k == kind] for kind in FIELDS}
    import receivables
    result['collections'] = receivables.snapshot(root)
    result['salary_sync'] = salary_sync
    return result

def mutate(root, payload):
    if payload.get('kind') == 'salary_sync':
        import salary_costs
        return salary_costs.sync(root)
    if payload.get('kind') == 'receivables':
        import receivables
        return receivables.mutate(root, payload)
    if payload.get('kind') == 'costs' and payload.get('action') == 'import':
        return import_costs(root, payload)
    if payload.get('kind') == 'changes':
        return _mutate_change(root, payload)
    kind = payload.get('kind')
    if kind not in FIELDS:
        raise ValueError('未知明细类型')
    action = payload.get('action')
    row = payload.get('row') or {}
    record_id = str(payload.get('id') or '')
    if action not in ('save', 'delete'):
        raise ValueError('未知操作')
    if kind == 'costs' and record_id:
        with connect(root) as db:
            saved = db.execute("SELECT data FROM ledger WHERE id=? AND kind='costs'", (record_id,)).fetchone()
            if saved and json.loads(saved[0]).get('source') == 'personnel_salary_v1':
                raise ValueError('工资成本由人员管理自动同步，请修改组织架构工价或人员出勤后刷新')
    if action == 'save':
        row = {key: str(row.get(key, '')).strip() for key in FIELDS[kind]}
        for key in (['name', 'source', 'category', 'unit', 'scope'] if kind == 'resources' else ['name', 'category', 'currency', 'date']):
            if not row[key]:
                raise ValueError('请填写必填信息')
        key = 'quantity' if kind == 'resources' else 'amount'
        try:
            value = float(row[key])
        except (ValueError, TypeError):
            raise ValueError('数量或金额必须为有效数字')
        if not math.isfinite(value) or value <= 0:
            raise ValueError('数量或金额必须大于零')
        row[key] = value
        for key in ('acquired', 'use_start', 'use_end', 'date'):
            if row.get(key):
                datetime.date.fromisoformat(row[key])
        if row.get('use_start') and row.get('use_end') and row['use_end'] < row['use_start']:
            raise ValueError('使用结束日期不能早于开始日期')
    with connect(root) as db:
        if record_id and not db.execute('SELECT 1 FROM ledger WHERE id=? AND kind=?', (record_id, kind)).fetchone():
            raise ValueError('记录已不存在，请刷新')
        if action == 'delete':
            if not record_id:
                raise ValueError('缺少记录编号')
            db.execute('DELETE FROM ledger WHERE id=? AND kind=?', (record_id, kind))
        else:
            record_id = record_id or uuid.uuid4().hex
            db.execute('INSERT OR REPLACE INTO ledger(id,kind,data) VALUES(?,?,?)', (record_id, kind, json.dumps(row, ensure_ascii=False)))
    return {'ok': True, 'msg': '已保存' if action == 'save' else '已删除'}

CHANGE_FIELDS = ['code','name','status','date','system','original','description','materials','cost_rule','cost_status','amount','currency','schedule_impact','responsibility','owner','signed_by','signed_date','source_doc','source_id','note']
CHANGE_STATUSES = ['待确认','待签字','已确认','实施中','已完成','已取消']

def _mutate_change(root, payload):
    action=payload.get('action'); record_id=str(payload.get('id') or '')
    if action not in ('save','delete'):raise ValueError('未知操作')
    row={key:str((payload.get('row') or {}).get(key,'')).strip() for key in CHANGE_FIELDS}
    if action=='save':
        if not row['code'] or not row['name'] or not row['description']:raise ValueError('请填写变更编号、名称及内容')
        if row['status'] not in CHANGE_STATUSES:raise ValueError('变更状态无效')
        for key in ('date','signed_date'):
            if row[key]:datetime.date.fromisoformat(row[key])
        if row['status'] in ('已确认','实施中','已完成') and (not row['signed_by'] or not row['signed_date']):raise ValueError('请先填写双方签署代表及签署日期')
        if row['cost_status'] not in ('待核算','部分不增费，其余待核算','已确认','不涉及费用'):raise ValueError('费用确认状态无效')
        if row['currency'] not in ('IDR','CNY','USD','THB'):raise ValueError('币种无效')
        if row['amount']!='':
            try: value=float(row['amount'])
            except (ValueError,TypeError):raise ValueError('变更金额必须为有效数字')
            if not math.isfinite(value):raise ValueError('变更金额必须为有限数字')
            row['amount']=value
        if row['source_id'] and (len(row['source_id'])!=32 or any(c not in '0123456789abcdef' for c in row['source_id'])):raise ValueError('来源资料编号无效')
    with connect(root) as db:
        db.execute('CREATE TABLE IF NOT EXISTS change_ledger_history (id TEXT PRIMARY KEY, change_id TEXT, action TEXT, before_json TEXT, after_json TEXT, created_at TEXT)')
        old=db.execute("SELECT data FROM ledger WHERE id=? AND kind='changes'",(record_id,)).fetchone() if record_id else None
        if record_id and not old:raise ValueError('记录已不存在，请刷新')
        before=json.loads(old[0]) if old else {}
        if action=='delete':
            if not record_id:raise ValueError('缺少记录编号')
            db.execute("DELETE FROM ledger WHERE id=? AND kind='changes'",(record_id,))
            after={}
        else:
            for other_id,raw in db.execute("SELECT id,data FROM ledger WHERE kind='changes'"):
                if other_id!=record_id and json.loads(raw).get('code','').casefold()==row['code'].casefold():raise ValueError('本项目变更编号已存在')
            record_id=record_id or uuid.uuid4().hex
            row['created_at']=before.get('created_at') or datetime.datetime.now().isoformat(timespec='seconds')
            row['updated_at']=datetime.datetime.now().isoformat(timespec='seconds')
            db.execute('INSERT OR REPLACE INTO ledger(id,kind,data) VALUES(?,?,?)',(record_id,'changes',json.dumps(row,ensure_ascii=False)))
            after=row
        db.execute('INSERT INTO change_ledger_history VALUES(?,?,?,?,?,?)',(uuid.uuid4().hex,record_id,action,json.dumps(before,ensure_ascii=False),json.dumps(after,ensure_ascii=False),datetime.datetime.now().isoformat(timespec='seconds')))
    return {'ok':True,'id':record_id,'msg':'变更已保存' if action=='save' else '变更已删除，历史记录保留'}

FIELDS['changes'] = CHANGE_FIELDS


COST_HEADERS = ['支出内容','开支类别','金额','币种','发生日期','使用范围 / 子系统','经办人','备注 / 凭证编号']

def import_costs(root, payload):
    import csv, io
    content = payload.get('content')
    if not isinstance(content, str) or len(content) > 2000000:
        raise ValueError('请选择不超过 2 MB 的 CSV 文件')
    try:
        reader = csv.reader(io.StringIO(content.lstrip('\ufeff')), strict=True)
        header = next(reader, [])
        if [v.strip() for v in header] != COST_HEADERS:
            raise ValueError('表头不匹配，请使用下载的模板并保留表头顺序')
        records = []
        for cells in reader:
            line = reader.line_num
            if not any(v.strip() for v in cells): continue
            if len(records) >= 5000: raise ValueError('每次最多导入 5000 条')
            if len(cells) != len(COST_HEADERS): raise ValueError(f'第 {line} 行：列数不正确')
            row = dict(zip(FIELDS['costs'], (v.strip() for v in cells)))
            try:
                if not row['name']: raise ValueError('支出内容必填')
                if row['category'] not in ('人工','耗材','交通','话费','商务','设备','其他'): raise ValueError('开支类别无效')
                if row['currency'] not in ('CNY','THB','USD','IDR'): raise ValueError('币种无效')
                row['amount'] = float(row['amount'])
                if not math.isfinite(row['amount']) or row['amount'] <= 0: raise ValueError('金额须为大于零的数字')
                row['date'] = datetime.date.fromisoformat(row['date']).isoformat()
            except (ValueError, TypeError) as error:
                raise ValueError(f'第 {line} 行：{error}')
            records.append(row)
    except csv.Error as error: raise ValueError('CSV 格式错误：'+str(error))
    if not records: raise ValueError('文件中没有支出记录')
    def key(row):
        return tuple(str(float(row.get(k,0))) if k == 'amount' else str(row.get(k,'')).strip() for k in FIELDS['costs'])
    with connect(root) as db:
        db.execute('BEGIN IMMEDIATE')
        seen = {key(json.loads(r[0])) for r in db.execute("SELECT data FROM ledger WHERE kind='costs'")}
        pending = []; duplicates = 0
        for row in records:
            token = key(row)
            if token in seen: duplicates += 1; continue
            seen.add(token); pending.append(row)
        if payload.get('dry_run') is not True:
            db.executemany('INSERT INTO ledger VALUES(?,?,?)', [(uuid.uuid4().hex,'costs',json.dumps(row,ensure_ascii=False)) for row in pending])
    totals = {}
    for row in pending: totals[row['currency']] = totals.get(row['currency'],0) + row['amount']
    return {'ok':True,'count':len(pending),'duplicates':duplicates,'totals':totals,'preview':pending[:20], 'msg':f'新增 {len(pending)} 条，跳过完全重复记录 {duplicates} 条'}
