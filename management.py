"""Project-scoped resource and expense ledgers; quality uses live acceptance data."""
import datetime
import math
import os
import sqlite3
import json
import uuid
from contextlib import contextmanager

FIELDS = {
    'resources': ['name', 'source', 'category', 'quantity', 'unit', 'acquired', 'scope', 'use_start', 'use_end', 'owner', 'note'],
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
    with connect(root) as db:
        rows = db.execute('SELECT id, kind, data FROM ledger ORDER BY rowid DESC').fetchall()
    return {kind: [dict(json.loads(data), id=id) for id, k, data in rows if k == kind] for kind in FIELDS}

def mutate(root, payload):
    kind = payload.get('kind')
    if kind not in FIELDS:
        raise ValueError('未知明细类型')
    action = payload.get('action')
    row = payload.get('row') or {}
    record_id = str(payload.get('id') or '')
    if action not in ('save', 'delete'):
        raise ValueError('未知操作')
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
