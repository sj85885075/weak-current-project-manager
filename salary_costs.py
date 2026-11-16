"""Reconcile project-local earned wages with the cost ledger."""
import datetime
import hashlib
import json
import os
import unicodedata
from collections import Counter
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

SOURCE = 'personnel_salary_v1'
CURRENCIES = ('IDR', 'CNY', 'USD', 'THB')


def read_people(root):
    import openpyxl
    import org_deliver
    path = os.path.join(root, '26_项目人员组织架构.xlsx')
    if not os.path.exists(path):
        return []
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = book['人员信息'].iter_rows(values_only=True)
        headers = list(next(rows, ()))
        keys = dict(zip(org_deliver.ORG_HEADERS, org_deliver.ORG_KEYS))
        people = [{keys[h]: row[i] if i < len(row) else None
                   for i, h in enumerate(headers) if h in keys}
                  for row in rows if any(v is not None for v in row)]
    finally:
        book.close()
    org_deliver.attach_attendance('', people, os.path.join(root, '29_项目人员管理表.xlsx'))
    return people


def desired_rows(people):
    def key(value):
        return unicodedata.normalize('NFKC', str(value or '').strip()).casefold()
    names = Counter(key(p.get('name')) for p in people)
    ids = Counter(str(p.get('pid') or '') for p in people)
    result, warnings = {}, []
    for p in people:
        if p.get('status') == '已删除':
            continue
        name, pid = str(p.get('name') or ''), str(p.get('pid') or '')
        reason = ''
        if not pid or ids[pid] > 1:
            reason = '人员ID缺失或重复'
        elif not key(name) or names[key(name)] > 1 or p.get('attendance_warning'):
            reason = '姓名或出勤关联待核对'
        try:
            rate = Decimal(str(p.get('daily_rate')))
            if not rate.is_finite() or rate < 0:
                raise InvalidOperation
        except (InvalidOperation, ValueError):
            rate = None
            reason = reason or '日工价未填写或无效'
        currency = str(p.get('salary_currency') or '')
        if currency not in CURRENCIES:
            reason = reason or '请在组织架构填写工资币种'
        if reason:
            warnings.append(name + '：' + reason)
            continue
        days = int(p.get('work_days') or 0)
        amount = (rate * days).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        if amount <= 0:
            continue
        record_id = 'salary_' + hashlib.sha256(pid.encode('utf-8')).hexdigest()[:32]
        result[record_id] = {
            'name': name + '工资（本项目累计）', 'category': '人工',
            'amount': float(amount), 'currency': currency,
            'date': p.get('last_work_date') or datetime.date.today().isoformat(),
            'scope': p.get('group') or '人员工资', 'owner': name,
            'note': f'人员工资自动同步；人员ID：{pid}；日工价 {rate} {currency} × 已保存出勤 {days} 天；'
                    '按当前工价计算应发工资，已发工资不重复计入成本。'
                    f'出勤期间：{p.get("first_work_date", "")} 至 {p.get("last_work_date", "")}。',
            'source': SOURCE, 'source_person_id': pid,
        }
    return result, warnings


def sync(root):
    import management
    desired, warnings = desired_rows(read_people(root))
    added = updated = removed = 0
    with management.connect(root) as db:
        db.execute('CREATE TABLE IF NOT EXISTS salary_cost_history '
                   '(id INTEGER PRIMARY KEY, record_id TEXT, before_json TEXT, after_json TEXT, created_at TEXT)')
        existing = {id: json.loads(raw) for id, raw in db.execute("SELECT id,data FROM ledger WHERE kind='costs'")}
        owned = {id: row for id, row in existing.items() if row.get('source') == SOURCE}
        for id in desired:
            if id in existing and id not in owned:
                raise ValueError('工资同步编号冲突，请核对成本记录')
        for id in owned.keys() | desired.keys():
            before, after = owned.get(id), desired.get(id)
            if before == after:
                continue
            db.execute('INSERT INTO salary_cost_history(record_id,before_json,after_json,created_at) VALUES(?,?,?,?)',
                       (id, json.dumps(before, ensure_ascii=False), json.dumps(after, ensure_ascii=False),
                        datetime.datetime.now().isoformat(timespec='seconds')))
            if after is None:
                db.execute("DELETE FROM ledger WHERE id=? AND kind='costs'", (id,))
                removed += 1
            else:
                db.execute('INSERT INTO ledger(id,kind,data) VALUES(?,?,?) ON CONFLICT(id) DO UPDATE SET data=excluded.data',
                           (id, 'costs', json.dumps(after, ensure_ascii=False)))
                if before is None:
                    added += 1
                else:
                    updated += 1
    totals = {}
    for row in desired.values():
        totals[row['currency']] = totals.get(row['currency'], 0) + row['amount']
    return {'ok': True, 'msg': f'工资已同步：新增{added}条，更新{updated}条，移除{removed}条',
            'count': len(desired), 'totals': totals, 'warnings': warnings}
