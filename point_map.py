"""Project-scoped drawing geometry and installation records.

The seed is read-only source geometry. Field edits and manually added devices
are persisted separately, so drawing updates cannot erase site progress.
"""
import datetime
import copy
from functools import lru_cache
import json
import math
import os
import sqlite3
import uuid
from contextlib import contextmanager


@contextmanager
def connect(root):
    os.makedirs(root, exist_ok=True)
    db = sqlite3.connect(os.path.join(root, '点位图状态.db'), timeout=20)
    db.row_factory = sqlite3.Row
    try:
        db.executescript('''
          CREATE TABLE IF NOT EXISTS records (
            id TEXT PRIMARY KEY, data TEXT NOT NULL,
            revision INTEGER NOT NULL, updated_at TEXT NOT NULL);
          CREATE TABLE IF NOT EXISTS history (
            seq INTEGER PRIMARY KEY AUTOINCREMENT, id TEXT NOT NULL,
            before_data TEXT, after_data TEXT NOT NULL, updated_at TEXT NOT NULL);
        ''')
        db.commit()
        yield db
    finally:
        db.close()


@lru_cache(maxsize=8)
def _read_seed(path,mtime,size):
    with open(path,encoding='utf-8') as stream:return json.load(stream)

def seed(root):
    path=os.path.join(root,'点位图基础数据.json')
    if not os.path.exists(path):
        return {'meta':{'title':'点位图','notes':['当前项目尚未导入图纸，可在图中新增点位和线路。']},'devices':[],'lines':[],'background':[]}
    stamp=os.stat(path);data=_read_seed(os.path.abspath(path),stamp.st_mtime_ns,stamp.st_size)
    if data.get('meta',{}).get('cad_underlay'):
        data={k:v for k,v in data.items() if k not in ('architecture','background')}
    return copy.deepcopy(data)


def snapshot(root):
    result = seed(root)
    with connect(root) as db:
        records = {r['id']: (json.loads(r['data']), r['revision'], r['updated_at'])
                   for r in db.execute('SELECT * FROM records')}
    known = set()
    for kind in ('devices', 'lines'):
        for item in result[kind]:
            known.add(item['id'])
            data, rev, stamp = records.get(item['id'], ({}, 0, ''))
            item.update(installed=False, commissioned=False, laid=False,
                        measured_length_m=None, note='', revision=rev, updated_at=stamp)
            item.update(data)
        for id, (data, rev, stamp) in records.items():
            if id not in known and data.get('custom') and data.get('record_kind') == kind:
                result[kind].append(dict(data, id=id, revision=rev, updated_at=stamp))
    return dict(result, ok=True)


def _text(value, maximum):
    if not isinstance(value, str) or len(value) > maximum:
        raise ValueError('文本内容过长或格式无效')
    return value.strip()


def _points(value, line=False):
    if not isinstance(value, list) or not (2 if line else 1) <= len(value) <= 500:
        raise ValueError('线路至少包含两个坐标点' if line else '点位坐标无效')
    for pair in value:
        if not isinstance(pair, list) or len(pair) != 2:
            raise ValueError('坐标格式无效')
        if any(isinstance(n, bool) or not isinstance(n, (float, int)) or not math.isfinite(n) or abs(n) > 1e9 for n in pair):
            raise ValueError('坐标无效')
    return value


def mutate(root, payload):
    current = snapshot(root)
    creating = payload.get('action') == 'create'
    if payload.get('action') not in ('create', 'update'):
        raise ValueError('未知操作')
    patch = payload.get('patch')
    if not isinstance(patch, dict):
        raise ValueError('更新内容无效')
    item = None
    kind = payload.get('kind')
    if creating:
        if kind not in ('devices', 'lines'):
            raise ValueError('点位类型无效')
        name = _text(patch.get('name', ''), 120)
        category = _text(patch.get('category', ''), 60)
        if not name or not category:
            raise ValueError('请输入编号和设备或线路类型')
        if any(i.get('name') == name for i in current[kind]):
            raise ValueError('编号已存在')
        points = _points(patch.get('points'), kind == 'lines')
        id = 'manual-' + uuid.uuid4().hex
        data = dict(id=id, name=name, category=category, record_kind=kind,
                    custom=True, source='现场补充', installed=False,
                    commissioned=False, laid=False, measured_length_m=None,
                    note='', points=points, zone=_text(patch.get('zone', '现场补充'), 80))
        if kind == 'devices':
            data.update(x=points[0][0], y=points[0][1])
        else:
            data['drawing_length_m'] = round(sum(math.dist(a, b) for a, b in zip(points, points[1:])), 3)
        revision = 0
    else:
        id = str(payload.get('id') or '')
        item = next((i for k in ('devices', 'lines') for i in current[k] if i['id'] == id), None)
        if item is None:
            raise ValueError('点位或线路不存在，请刷新')
        if payload.get('revision') != item['revision']:
            raise ValueError('该记录已被其他窗口修改，请刷新后重试')
        revision = item['revision']
        data = {k: item[k] for k in ('installed', 'commissioned', 'laid', 'measured_length_m', 'note', 'name', 'category', 'zone') if k in item}
        if item.get('custom'):
            data = {k: v for k, v in item.items() if k not in ('revision', 'updated_at')}
        kind = 'lines' if any(i['id'] == id for i in current['lines']) else 'devices'
    for key, limit in (('name', 120), ('category', 60), ('zone', 80)):
        if key in patch:
            value = _text(patch[key], limit)
            if not value:
                raise ValueError('编号、类别和区域不能为空')
            if key == 'name' and any(i['id'] != id and i.get('name') == value for i in current[kind]):
                raise ValueError('编号已存在')
            data[key] = value
    for key in ('installed', 'commissioned', 'laid'):
        if key in patch:
            if not isinstance(patch[key], bool):
                raise ValueError('状态必须是勾选值')
            if (kind == 'lines') != (key == 'laid'):
                raise ValueError('状态不适用于此类记录')
            data[key] = patch[key]
    if data.get('commissioned') and not data.get('installed'):
        raise ValueError('设备安装完成后才能标记调试完成')
    if 'measured_length_m' in patch:
        value = patch['measured_length_m']
        if kind != 'lines':
            raise ValueError('只有线路可以填写长度')
        if value is not None and (isinstance(value, bool) or not isinstance(value, (float, int)) or not math.isfinite(value) or not 0 <= value <= 1e7):
            raise ValueError('确认长度必须为非负有限数字，未确认请留空')
        data['measured_length_m'] = value
    if 'note' in patch:
        data['note'] = _text(patch['note'], 2000)
    stamp = datetime.datetime.now(datetime.timezone.utc).isoformat(timespec='seconds')
    encoded = json.dumps(data, ensure_ascii=False)
    with connect(root) as db:
        with db:
            db.execute('BEGIN IMMEDIATE')
            before = db.execute('SELECT * FROM records WHERE id=?', (id,)).fetchone()
            if (before['revision'] if before else 0) != revision:
                raise ValueError('该记录已被其他窗口修改，请刷新后重试')
            db.execute('INSERT OR REPLACE INTO records VALUES(?,?,?,?)', (id, encoded, revision+1, stamp))
            db.execute('INSERT INTO history(id,before_data,after_data,updated_at) VALUES(?,?,?,?)',
                       (id, before['data'] if before else None, encoded, stamp))
    return {'ok': True, 'msg': '已保存', 'id': id, 'revision': revision+1}

from point_design import design_snapshot as snapshot, design_mutate as mutate, export_dxf
