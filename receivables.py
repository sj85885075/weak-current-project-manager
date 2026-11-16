"""Project-scoped collection schedule and receipt ledger. Monetary values use decimal strings."""
import calendar
import datetime as dt
import json
import sqlite3
import uuid
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from contextlib import contextmanager

STAGES=[
    dict(code='advance',name='预付款',percent=50,days=14,months=0,condition='合同签署后 14 个日历日内'),
    dict(code='second',name='第二期款',percent=45,days=14,months=0,condition='全部设备到货、物理安装完成且 BAST 签署后 14 个日历日内（填写全部条件满足日期）'),
    dict(code='retention',name='尾款 / 保留金',percent=5,days=0,months=3,condition='配置完成、测试通过、培训完成且最终验收 / 移交报告签署后 3 个月（填写全部条件满足日期）'),
]

@contextmanager
def connect(root):
    db=sqlite3.connect(Path(root)/'资源成本管理.db')
    try:
        with db:
            db.execute('CREATE TABLE IF NOT EXISTS ledger (id TEXT PRIMARY KEY, kind TEXT NOT NULL, data TEXT NOT NULL)')
            db.execute('CREATE TABLE IF NOT EXISTS collection_history (id TEXT PRIMARY KEY, action TEXT, before_json TEXT, after_json TEXT, created_at TEXT)')
            yield db
    finally: db.close()

def money(value,positive=False):
    try:
        v=Decimal(str(value))
        if not v.is_finite() or v<0 or (positive and v<=0) or v>Decimal('999999999999999'):raise ValueError()
        if v!=v.quantize(Decimal('.01')):raise ValueError()
    except (InvalidOperation,ValueError,TypeError):raise ValueError('金额须为有效非负数字，最多两位小数；合同及回款金额须大于零')
    return v.quantize(Decimal('.01'))

def date(value):
    if not value:return ''
    try:return dt.date.fromisoformat(str(value)).isoformat()
    except (ValueError,TypeError):raise ValueError('日期格式须为 YYYY-MM-DD')

def deadline(value,stage):
    if not value:return ''
    d=dt.date.fromisoformat(value)
    if stage['months']:
        m=d.month-1+stage['months'];y=d.year+m//12;m=m%12+1
        return d.replace(year=y,month=m,day=min(d.day,calendar.monthrange(y,m)[1])).isoformat()
    return (d+dt.timedelta(days=stage['days'])).isoformat()

def read(db):
    row=db.execute("SELECT data FROM ledger WHERE id='collection-plan' AND kind='collection_plan'").fetchone()
    plan=json.loads(row[0]) if row else None
    receipts=[dict(json.loads(raw),id=i) for i,raw in db.execute("SELECT id,data FROM ledger WHERE kind='collection_receipt' ORDER BY rowid DESC")]
    return plan,receipts

def calculate(plan,receipts,today=None):
    if not plan:return dict(plan=None,receipts=[],stages=[],version=0)
    # Client business timezone is UTC+7.
    today=today or dt.datetime.now(dt.timezone(dt.timedelta(hours=7))).date().isoformat()
    total=money(plan['amount']);allocated=Decimal(0);stages=[]
    for index,stage in enumerate(STAGES):
        amount=total-allocated if index==len(STAGES)-1 else (total*Decimal(stage['percent'])/100).quantize(Decimal('.01'),rounding=ROUND_HALF_UP)
        allocated+=amount;received=sum((money(r['amount']) for r in receipts if r['stage']==stage['code']),Decimal(0))
        trigger=plan.get('dates',{}).get(stage['code'],'');due=deadline(trigger,stage);balance=amount-received
        status='已收齐' if balance<=0 else '逾期未核销' if due and due<today else '部分回款' if received else '待回款登记' if due else '待触发'
        stages.append(dict(stage,amount=str(amount),received=str(received),balance=str(balance),trigger_date=trigger,due_date=due,status=status))
    received=sum((money(r['amount']) for r in receipts),Decimal(0))
    overdue=sum((money(s['balance']) for s in stages if s['due_date'] and s['due_date']<today and money(s['balance'])>0),Decimal(0))
    return dict(plan=plan,receipts=receipts,stages=stages,version=plan['version'],totals=dict(planned=str(total),received=str(received),balance=str(total-received),overdue=str(overdue)))

def snapshot(root):
    with connect(root) as db:return calculate(*read(db))

def mutate(root,payload):
    with connect(root) as db:
        db.execute('BEGIN IMMEDIATE')
        plan,receipts=read(db);before=dict(plan=plan,receipts=receipts)
        if payload.get('expected_version')!=(plan['version'] if plan else 0):raise ValueError('回款数据已变化，请刷新后重试')
        action=payload.get('action');row=payload.get('row') or {}
        if action=='save_plan':
            amount=money(row.get('amount'),True);currency=str(row.get('currency',''))
            if currency not in ('IDR','CNY','USD','THB'):raise ValueError('币种无效')
            if receipts and currency!=plan['currency']:raise ValueError('已有回款记录，不能直接变更币种')
            dates={s['code']:date((row.get('dates') or {}).get(s['code'],'')) for s in STAGES}
            if dates['advance'] and any(v and v<dates['advance'] for k,v in dates.items() if k!='advance'):raise ValueError('验收节点不能早于合同签署日期')
            if dates['second'] and dates['retention'] and dates['retention']<dates['second']:raise ValueError('最终验收不能早于设备安装及 BAST 节点')
            plan=dict(amount=str(amount),currency=currency,dates=dates,source=str(row.get('source',''))[:1000],note=str(row.get('note',''))[:4000],version=(plan['version'] if plan else 0))
            preview=calculate(plan,receipts)
            if any(Decimal(s['balance'])<0 for s in preview['stages']):raise ValueError('修改后的分期金额小于已登记回款，请先核对回款记录')
        elif action in ('save_receipt','delete_receipt'):
            if not plan:raise ValueError('请先配置合同回款计划')
            rid=str(payload.get('id') or '');existing=next((r for r in receipts if r['id']==rid),None)
            if rid and not existing:raise ValueError('该回款记录已不存在')
            if action=='delete_receipt':
                if not existing:raise ValueError('请选择回款记录')
                db.execute("DELETE FROM ledger WHERE id=? AND kind='collection_receipt'",(rid,))
            else:
                stage=str(row.get('stage',''));amount=money(row.get('amount'),True);received_date=date(row.get('date'))
                if stage not in {s['code'] for s in STAGES} or not received_date:raise ValueError('请选择期次并填写实收日期')
                receipt=dict(stage=stage,amount=str(amount),date=received_date,reference=str(row.get('reference',''))[:300],note=str(row.get('note',''))[:4000])
                remaining=[r for r in receipts if r['id']!=rid]
                trial=calculate(plan,remaining+[dict(receipt,id=rid)])
                if any(Decimal(s['balance'])<0 for s in trial['stages']):raise ValueError('累计回款超过该期应收金额，请核对期次或金额')
                if not rid and any(all(r.get(k,'')==v for k,v in receipt.items()) for r in receipts):raise ValueError('完全相同的回款记录已存在')
                rid=rid or uuid.uuid4().hex
                db.execute('INSERT OR REPLACE INTO ledger VALUES(?,?,?)',(rid,'collection_receipt',json.dumps(receipt,ensure_ascii=False)))
        else:raise ValueError('未知回款操作')
        plan['version']+=1
        db.execute('INSERT OR REPLACE INTO ledger VALUES(?,?,?)',('collection-plan','collection_plan',json.dumps(plan,ensure_ascii=False)))
        after_plan,after_receipts=read(db)
        db.execute('INSERT INTO collection_history VALUES(?,?,?,?,?)',(uuid.uuid4().hex,action,json.dumps(before,ensure_ascii=False),json.dumps(dict(plan=after_plan,receipts=after_receipts),ensure_ascii=False),dt.datetime.now(dt.timezone.utc).isoformat()))
    return {'ok':True,'msg':'回款计划已保存' if action=='save_plan' else '回款记录已删除' if action=='delete_receipt' else '回款已登记'}
