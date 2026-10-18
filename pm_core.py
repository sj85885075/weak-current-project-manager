# -*- coding: utf-8 -*-
"""项目管理核心数据层（SQLite）。

本模块不替换现场既有 Excel 台账；它保存项目参数、软考过程库、WBS、
实际完成量和变更留痕，供前端和报表统一读取。
"""
import datetime
import base64
import csv
import io
import json
import os
import sqlite3
import uuid
import math
import openpyxl

GROUPS = [
    ('启动', '明确项目授权、目标、范围边界和主要相关方。'),
    ('规划', '建立可执行、可度量、可批准的项目管理基准。'),
    ('执行', '组织资源完成施工、安装、调试、测试和沟通。'),
    ('监控', '比较计划与实际，闭环质量、风险、问题和变更。'),
    ('收尾', '完成验收、移交、归档和项目总结。'),
]
DOMAINS = ['项目整合管理', '项目范围管理', '项目进度管理', '项目成本管理', '项目质量管理',
           '项目资源管理', '项目沟通管理', '项目风险管理', '项目采购管理', '项目相关方管理']
PROCESSES = [
    ('制定项目章程', '启动', '项目整合管理', 2, '项目章程、合同要点', '项目章程'),
    ('识别主要相关方', '启动', '项目相关方管理', 1, '合同、组织信息', '相关方登记册'),
    ('项目启动会议', '启动', '项目沟通管理', 1, '项目章程、相关方清单', '启动会议纪要'),
    ('范围规划与WBS', '规划', '项目范围管理', 2, '项目参数、图纸', '范围说明书、WBS'),
    ('进度计划', '规划', '项目进度管理', 2, 'WBS、工程量、日历', '项目进度计划'),
    ('资源与成本计划', '规划', '项目资源管理', 2, '工程量、工效、人员', '资源计划、成本计划'),
    ('质量、风险与沟通计划', '规划', '项目质量管理', 2, '合同、规范、图纸', '质量/风险/沟通计划'),
    ('采购计划', '规划', '项目采购管理', 2, '清单、交期、预算', '采购计划'),
    ('施工与安装', '执行', '项目范围管理', 10, '施工计划、材料、人员', '施工记录、安装记录'),
    ('系统调试与测试', '执行', '项目质量管理', 5, '安装记录、测试标准', '调试记录、测试报告'),
    ('进度与工程量统计', '监控', '项目进度管理', 1, '日报、实际完成量', '进度报告'),
    ('质量检查与问题闭环', '监控', '项目质量管理', 1, '检查记录、问题单', '质量检查记录'),
    ('风险与变更控制', '监控', '项目风险管理', 1, '风险/变更申请', '风险登记册、变更记录'),
    ('系统验收与移交', '收尾', '项目整合管理', 3, '测试报告、竣工资料', '验收报告、移交清单'),
    ('资料归档与项目总结', '收尾', '项目整合管理', 1, '全部交付文件', '项目总结报告、竣工资料目录'),
]


def _now():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _root(base_dir, proj):
    if proj == 'titan': return os.path.join(base_dir, '泰坦IDC')
    if proj == 'bg': return base_dir
    return os.path.join(base_dir, '自定义项目', proj)


def _db_path(base_dir, proj):
    return os.path.join(_root(base_dir, proj), '项目管理核心数据.db')


def connect(base_dir, proj):
    path = _db_path(base_dir, proj)
    con = sqlite3.connect(path)
    con.row_factory = sqlite3.Row
    con.execute('PRAGMA foreign_keys = ON')
    return con


def ensure(base_dir, proj, project_name, systems):
    con = connect(base_dir, proj)
    con.executescript('''
        CREATE TABLE IF NOT EXISTS project_profile (
          project_key TEXT PRIMARY KEY, project_name TEXT NOT NULL, project_no TEXT,
          project_type TEXT, owner_unit TEXT, contractor TEXT, manager TEXT,
          technical_lead TEXT, location TEXT, contract_amount REAL DEFAULT 0,
          plan_start TEXT, plan_finish TEXT, work_hours REAL DEFAULT 8,
          work_days INTEGER DEFAULT 6, personnel_count INTEGER DEFAULT 0,
          custom_quantities TEXT DEFAULT '{}', updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS process_groups (code TEXT PRIMARY KEY, description TEXT NOT NULL, sort_no INTEGER NOT NULL);
        CREATE TABLE IF NOT EXISTS knowledge_domains (name TEXT PRIMARY KEY, sort_no INTEGER NOT NULL);
        CREATE TABLE IF NOT EXISTS process_library (
          id TEXT PRIMARY KEY, name TEXT NOT NULL, process_group TEXT NOT NULL,
          domain TEXT NOT NULL, default_days REAL DEFAULT 1, inputs TEXT, outputs TEXT,
          mandatory INTEGER DEFAULT 1, allow_skip INTEGER DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS wbs_tasks (
          id TEXT PRIMARY KEY, parent_id TEXT, wbs_code TEXT NOT NULL, system_name TEXT,
          task_name TEXT NOT NULL, process_group TEXT, domain TEXT, owner TEXT,
          plan_start TEXT, plan_finish TEXT, plan_qty REAL DEFAULT 0, actual_qty REAL DEFAULT 0,
          unit TEXT, status TEXT DEFAULT '未开始', predecessor TEXT, sort_no INTEGER DEFAULT 0,
          created_at TEXT NOT NULL, updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS actual_logs (
          id TEXT PRIMARY KEY, task_id TEXT NOT NULL, work_date TEXT NOT NULL, qty REAL NOT NULL,
          note TEXT, operator TEXT, created_at TEXT NOT NULL,
          FOREIGN KEY(task_id) REFERENCES wbs_tasks(id)
        );
        CREATE TABLE IF NOT EXISTS change_history (
          id TEXT PRIMARY KEY, entity_type TEXT NOT NULL, entity_id TEXT NOT NULL,
          action TEXT NOT NULL, before_json TEXT, after_json TEXT, operator TEXT, created_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_wbs_system ON wbs_tasks(system_name);
        CREATE INDEX IF NOT EXISTS idx_log_task_date ON actual_logs(task_id, work_date);
    ''')
    con.execute('INSERT OR IGNORE INTO project_profile(project_key, project_name, updated_at) VALUES(?,?,?)', (proj, project_name, _now()))
    for i, (code, description) in enumerate(GROUPS, 1):
        con.execute('INSERT OR IGNORE INTO process_groups VALUES(?,?,?)', (code, description, i))
    for i, name in enumerate(DOMAINS, 1):
        con.execute('INSERT OR IGNORE INTO knowledge_domains VALUES(?,?)', (name, i))
    for i, p in enumerate(PROCESSES, 1):
        con.execute('INSERT OR IGNORE INTO process_library VALUES(?,?,?,?,?,?,?,?,?)',
                    ('PROC-%02d' % i, p[0], p[1], p[2], p[3], p[4], p[5], 1, 0))
    columns = {r[1] for r in con.execute('PRAGMA table_info(wbs_tasks)')}
    if 'deleted' not in columns:
        con.execute('ALTER TABLE wbs_tasks ADD COLUMN deleted INTEGER DEFAULT 0')
    for name, ddl in [
        ('task_type', "TEXT DEFAULT '工作包'"),
        ('description', "TEXT DEFAULT ''"),
        ('deliverable', "TEXT DEFAULT ''"),
        ('acceptance_criteria', "TEXT DEFAULT ''"),
        ('weight', 'REAL DEFAULT 1'),
        ('source_type', "TEXT DEFAULT ''"),
        ('source_key', "TEXT DEFAULT ''"),
        ('last_synced_at', "TEXT DEFAULT ''"),
    ]:
        if name not in columns:
            con.execute('ALTER TABLE wbs_tasks ADD COLUMN %s %s' % (name, ddl))
    con.commit()
    return con


def _dict(row):
    return dict(row) if row else {}


def profile(base_dir, proj, project_name, systems):
    con = ensure(base_dir, proj, project_name, systems)
    ret = _dict(con.execute('SELECT * FROM project_profile WHERE project_key=?', (proj,)).fetchone())
    try: ret['custom_quantities'] = json.loads(ret.get('custom_quantities') or '{}')
    except Exception: ret['custom_quantities'] = {}
    con.close()
    return ret


def save_profile(base_dir, proj, project_name, systems, data, operator=''):
    con = ensure(base_dir, proj, project_name, systems)
    before = _dict(con.execute('SELECT * FROM project_profile WHERE project_key=?', (proj,)).fetchone())
    allowed = ['project_name','project_no','project_type','owner_unit','contractor','manager','technical_lead','location','contract_amount','plan_start','plan_finish','work_hours','work_days','personnel_count','custom_quantities']
    vals = {k: data.get(k, before.get(k)) for k in allowed}
    try:
        for key in ['plan_start','plan_finish']:
            if vals.get(key): datetime.date.fromisoformat(str(vals[key]))
        if vals.get('plan_start') and vals.get('plan_finish') and vals['plan_finish'] < vals['plan_start']:
            raise ValueError()
    except (ValueError, TypeError):
        con.close(); return {'ok': False, 'err': '计划日期无效，完成日期不能早于开始日期'}
    try:
        for key in ('contract_amount','personnel_count','work_hours','work_days'):
            value=float(vals.get(key) or 0)
            if not math.isfinite(value) or value<0:raise ValueError()
        if not 0<float(vals['work_hours'])<=24 or not 1<=float(vals['work_days'])<=7 or float(vals['work_days'])%1 or float(vals['personnel_count'] or 0)%1:raise ValueError()
    except (ValueError,TypeError):
        con.close();return {'ok':False,'err':'项目数值无效：工时须在 0 至 24 小时内，工作天数为 1 至 7 的整数，金额和人数不得为负'}
    if isinstance(vals.get('custom_quantities'), (dict, list)):
        vals['custom_quantities'] = json.dumps(vals['custom_quantities'], ensure_ascii=False)
    sets = ','.join(k+'=?' for k in allowed) + ',updated_at=?'
    con.execute('UPDATE project_profile SET '+sets+' WHERE project_key=?', [vals[k] for k in allowed]+[_now(),proj])
    after = _dict(con.execute('SELECT * FROM project_profile WHERE project_key=?', (proj,)).fetchone())
    _log(con, 'profile', proj, '更新项目基础参数', before, after, operator)
    con.commit(); con.close()
    return {'ok': True, 'msg': '项目基础参数已保存，并保留修改历史'}


def _dates(start, days):
    if not start: return '', ''
    try:
        d = datetime.date.fromisoformat(str(start)[:10]); return d.isoformat(), (d + datetime.timedelta(days=max(int(days)-1, 0))).isoformat()
    except ValueError: return '', ''


def generate_wbs(base_dir, proj, project_name, systems, quantity_map=None, operator=''):
    con = ensure(base_dir, proj, project_name, systems)
    exists = con.execute('SELECT COUNT(*) FROM wbs_tasks WHERE deleted=0').fetchone()[0]
    if exists:
        con.close(); return {'ok': False, 'err': 'WBS 已存在；为保护已录入的实际数据，未自动覆盖。请在 WBS 页面逐项调整。'}
    p = _dict(con.execute('SELECT * FROM project_profile WHERE project_key=?', (proj,)).fetchone())
    start = p.get('plan_start') or datetime.date.today().isoformat()
    workforce = max(float(p.get('personnel_count') or 1), 1)
    quantity_map = quantity_map or {}
    tasks = []
    cursor = 0
    for n, system in enumerate(systems, 1):
        raw = quantity_map.get(system, 0)
        try: qty = float(raw or 0)
        except Exception: qty = 0
        # 一人每日默认 10 单位，0 工程量仍生成管理任务但不虚构工期
        days = max(1, int((qty / max(workforce * 10, 1)) + 0.999)) if qty else 1
        ps, pe = _dates((datetime.date.fromisoformat(start)+datetime.timedelta(days=cursor)).isoformat(), days)
        tid = str(uuid.uuid4())
        tasks.append((tid, None, '%02d' % n, system, system, '执行', '项目范围管理', p.get('manager') or '', ps, pe, qty, 0, '项' if qty else '', '未开始', '', n, _now(), _now()))
        cursor += days
    for row in tasks:
        con.execute('''INSERT INTO wbs_tasks(id,parent_id,wbs_code,system_name,task_name,process_group,domain,owner,plan_start,plan_finish,plan_qty,actual_qty,unit,status,predecessor,sort_no,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', row)
    _log(con, 'wbs', proj, '根据项目参数自动生成WBS', {}, {'tasks':len(tasks)}, operator)
    con.commit(); con.close()
    return {'ok': True, 'msg': '已生成 %d 个子系统 WBS 任务；后续录入实际量会自动更新进度。' % len(tasks)}


IMPORT_HEADERS = ['WBS编码','上级WBS编码','任务类型','所属子系统','工作包名称','工作说明','责任人','计划开始','计划完成','计划量','单位','权重','前置WBS编码','交付物','验收标准']
HEADER_ALIASES = {
    'WBS编码':'wbs_code','wbs_code':'wbs_code','编码':'wbs_code',
    '上级WBS编码':'parent_code','parent_code':'parent_code','上级编码':'parent_code',
    '任务类型':'task_type','task_type':'task_type','类型':'task_type',
    '所属子系统':'system_name','system_name':'system_name','子系统':'system_name','施工分项':'system_name',
    '工作包名称':'task_name','task_name':'task_name','任务名称':'task_name','名称':'task_name',
    '工作说明':'description','description':'description','范围说明':'description',
    '责任人':'owner','责任班组':'owner','owner':'owner',
    '计划开始':'plan_start','plan_start':'plan_start','开始日期':'plan_start',
    '计划完成':'plan_finish','plan_finish':'plan_finish','完成日期':'plan_finish',
    '计划量':'plan_qty','plan_qty':'plan_qty','工程量':'plan_qty',
    '单位':'unit','unit':'unit',
    '权重':'weight','weight':'weight',
    '前置WBS编码':'predecessor','predecessor':'predecessor','前置任务':'predecessor',
    '交付物':'deliverable','deliverable':'deliverable',
    '验收标准':'acceptance_criteria','acceptance_criteria':'acceptance_criteria',
}
TASK_TYPES = ('汇总节点','工作包','活动','里程碑')


def wbs_template():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'WBS导入'
    ws.append(IMPORT_HEADERS)
    examples = [
        ['02','', '汇总节点','7F→2F布线','7F→2F布线','','7F布线班组','2026-09-24','2026-09-28',576,'条',1,'','576条链路交付','下级工作包全部完成'],
        ['02.01','02','汇总节点','7F→2F布线','7F-01列链路','','7F布线班组','2026-09-24','2026-09-25',0,'',1,'','','下级工作包全部完成'],
        ['02.01.01-R1','02.01','工作包','7F→2F布线','7F-01-20机柜／路由1链路交付','敷设、整理、双端标签和自检','7F布线班组','2026-09-24','2026-09-24',18,'条',1,'','链路安装记录','数量、点位、标签及外观检查合格'],
    ]
    for row in examples: ws.append(row)
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = 'A1:O%d' % ws.max_row
    widths = [14,16,12,18,34,36,18,14,14,12,10,10,18,28,34]
    for i,w in enumerate(widths,1): ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
    for cell in ws[1]:
        cell.font=openpyxl.styles.Font(name='微软雅黑',bold=True,color='FFFFFF')
        cell.fill=openpyxl.styles.PatternFill('solid',fgColor='1F4E78')
        cell.alignment=openpyxl.styles.Alignment(horizontal='center',vertical='center')
    note=wb.create_sheet('填写说明')
    notes=[('规则','说明'),('WBS编码','必填且项目内唯一；建议使用 02.01.01-R1。'),('上级WBS编码','顶层留空；非顶层必须引用文件内或系统已有编码。'),('任务类型','汇总节点、工作包、活动、里程碑。汇总节点的进度由叶子任务汇总。'),('计划量','汇总节点可填控制总量，但项目统计只汇总叶子任务，避免重复计算。'),('导入方式','同编码任务合并更新，保留任务ID和已登记实际量；新编码新增。'),('日期','使用 YYYY-MM-DD；完成日期不得早于开始日期。'),('前置WBS编码','可填写多个编码，以英文逗号分隔。')]
    for row in notes: note.append(row)
    note.column_dimensions['A'].width=18; note.column_dimensions['B'].width=85
    for cell in note[1]: cell.font=openpyxl.styles.Font(name='微软雅黑',bold=True,color='FFFFFF');cell.fill=openpyxl.styles.PatternFill('solid',fgColor='1F4E78')
    out=io.BytesIO();wb.save(out);return out.getvalue()


def _read_wbs_rows(filename, content_b64):
    try: raw=base64.b64decode(content_b64, validate=True)
    except Exception: raise ValueError('导入文件编码无效')
    if len(raw)>8*1024*1024: raise ValueError('导入文件不能超过 8MB')
    suffix=os.path.splitext(str(filename or ''))[1].lower()
    if suffix=='.xlsx':
        wb=openpyxl.load_workbook(io.BytesIO(raw),read_only=True,data_only=True)
        ws=wb['WBS导入'] if 'WBS导入' in wb.sheetnames else wb.active
        matrix=list(ws.iter_rows(values_only=True));wb.close()
    elif suffix=='.csv':
        text=raw.decode('utf-8-sig');matrix=list(csv.reader(io.StringIO(text)))
    else: raise ValueError('仅支持 .xlsx 或 .csv 文件')
    if not matrix: raise ValueError('导入文件为空')
    fields=[]
    for h in matrix[0]: fields.append(HEADER_ALIASES.get(str(h or '').strip(),''))
    if 'wbs_code' not in fields or 'task_name' not in fields: raise ValueError('必须包含 WBS编码 和 工作包名称 两列')
    rows=[]
    for row_no, values in enumerate(matrix[1:],2):
        data={fields[i]:values[i] for i in range(min(len(fields),len(values))) if fields[i]}
        if not any(v not in (None,'') for v in data.values()): continue
        data['_row']=row_no;rows.append(data)
    return rows


def import_wbs(base_dir, proj, project_name, systems, filename, content_b64, dry_run=True, operator=''):
    con=ensure(base_dir,proj,project_name,systems)
    try:
        raw_rows=_read_wbs_rows(filename,content_b64);errors=[];normalized=[];seen=set()
        existing={r['wbs_code']:dict(r) for r in con.execute('SELECT * FROM wbs_tasks WHERE deleted=0')}
        allowed_codes=set(existing)
        for item in raw_rows:
            row=item['_row'];code=str(item.get('wbs_code') or '').strip();name=str(item.get('task_name') or '').strip()
            if not code: errors.append('第%d行：WBS编码不能为空'%row)
            elif code in seen: errors.append('第%d行：WBS编码重复：%s'%(row,code))
            else: seen.add(code);allowed_codes.add(code)
            if not name: errors.append('第%d行：工作包名称不能为空'%row)
            task_type=str(item.get('task_type') or '工作包').strip()
            if task_type not in TASK_TYPES: errors.append('第%d行：任务类型必须为%s'%(row,'、'.join(TASK_TYPES)))
            parent=str(item.get('parent_code') or '').strip()
            ps=str(item.get('plan_start') or '').strip()[:10];pe=str(item.get('plan_finish') or '').strip()[:10]
            try:
                if ps: datetime.date.fromisoformat(ps)
                if pe: datetime.date.fromisoformat(pe)
                if ps and pe and pe<ps: raise ValueError()
            except ValueError: errors.append('第%d行：计划日期无效或完成日期早于开始日期'%row)
            try:
                qty=float(item.get('plan_qty') or 0);weight=float(item.get('weight') or 1)
                if not math.isfinite(qty) or qty<0 or not math.isfinite(weight) or weight<0: raise ValueError()
            except (ValueError,TypeError): errors.append('第%d行：计划量和权重必须为非负数字'%row);qty=0;weight=1
            normalized.append({'row':row,'wbs_code':code,'parent_code':parent,'task_type':task_type,'system_name':str(item.get('system_name') or '').strip(),'task_name':name,'description':str(item.get('description') or '').strip(),'owner':str(item.get('owner') or '').strip(),'plan_start':ps,'plan_finish':pe,'plan_qty':qty,'unit':str(item.get('unit') or '').strip(),'weight':weight,'predecessor':str(item.get('predecessor') or '').strip(),'deliverable':str(item.get('deliverable') or '').strip(),'acceptance_criteria':str(item.get('acceptance_criteria') or '').strip()})
        for item in normalized:
            if item['parent_code'] and item['parent_code'] not in allowed_codes: errors.append('第%d行：上级WBS编码不存在：%s'%(item['row'],item['parent_code']))
            if item['parent_code']==item['wbs_code']: errors.append('第%d行：任务不能以自身作为上级'%item['row'])
            for pred in [x.strip() for x in item['predecessor'].replace('；',',').replace(';',',').split(',') if x.strip()]:
                if pred not in allowed_codes: errors.append('第%d行：前置WBS编码不存在：%s'%(item['row'],pred))
        parent_map={x['wbs_code']:x['parent_code'] for x in normalized}
        for code in parent_map:
            trail=set();cur=code
            while cur in parent_map and parent_map[cur]:
                cur=parent_map[cur]
                if cur in trail or cur==code: errors.append('WBS层级存在循环：%s'%code);break
                trail.add(cur)
        result={'ok':not errors,'dry_run':bool(dry_run),'rows':len(normalized),'new':sum(1 for x in normalized if x['wbs_code'] not in existing),'update':sum(1 for x in normalized if x['wbs_code'] in existing),'errors':errors[:100]}
        if errors or dry_run:return result
        now=_now();code_to_id={k:v['id'] for k,v in existing.items()}
        for item in normalized:
            if item['wbs_code'] not in code_to_id: code_to_id[item['wbs_code']]=str(uuid.uuid4())
        for sort_no,item in enumerate(normalized,1):
            tid=code_to_id[item['wbs_code']];parent_id=code_to_id.get(item['parent_code']) if item['parent_code'] else None
            old=existing.get(item['wbs_code'])
            values=(parent_id,item['system_name'],item['task_name'],item['owner'],item['plan_start'],item['plan_finish'],item['plan_qty'],item['unit'],item['predecessor'],sort_no,item['task_type'],item['description'],item['deliverable'],item['acceptance_criteria'],item['weight'],now,tid)
            if old:
                con.execute('UPDATE wbs_tasks SET parent_id=?,system_name=?,task_name=?,owner=?,plan_start=?,plan_finish=?,plan_qty=?,unit=?,predecessor=?,sort_no=?,task_type=?,description=?,deliverable=?,acceptance_criteria=?,weight=?,updated_at=? WHERE id=?',values)
                _log(con,'wbs_task',tid,'导入合并WBS',old,{**item,'id':tid},operator)
            else:
                con.execute('''INSERT INTO wbs_tasks(id,parent_id,wbs_code,system_name,task_name,process_group,domain,owner,plan_start,plan_finish,plan_qty,actual_qty,unit,status,predecessor,sort_no,created_at,updated_at,deleted,task_type,description,deliverable,acceptance_criteria,weight) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',(tid,parent_id,item['wbs_code'],item['system_name'],item['task_name'],'执行','项目范围管理',item['owner'],item['plan_start'],item['plan_finish'],item['plan_qty'],0,item['unit'],'未开始',item['predecessor'],sort_no,now,now,0,item['task_type'],item['description'],item['deliverable'],item['acceptance_criteria'],item['weight']))
                _log(con,'wbs_task',tid,'导入新增WBS',{},item,operator)
        con.commit();result['ok']=True;result['msg']='WBS导入完成：新增%d项，更新%d项；已保留历史实际量。'%(result['new'],result['update']);return result
    except ValueError as e:
        return {'ok':False,'errors':[str(e)],'rows':0,'new':0,'update':0,'dry_run':bool(dry_run)}
    finally: con.close()


def _log(con, etype, eid, action, before, after, operator):
    con.execute('INSERT INTO change_history VALUES(?,?,?,?,?,?,?,?)', (str(uuid.uuid4()), etype, eid, action, json.dumps(before, ensure_ascii=False, default=str), json.dumps(after, ensure_ascii=False, default=str), operator or '', _now()))


def sync_subsystem_tasks(base_dir, proj, project_name, systems, source_map):
    """把子系统主计量项同步到一级WBS；不新增实际流水，避免双重记账。"""
    con=ensure(base_dir,proj,project_name,systems);updated=0
    try:
        con.execute('BEGIN IMMEDIATE')
        for system,data in (source_map or {}).items():
            row=con.execute("SELECT * FROM wbs_tasks WHERE deleted=0 AND parent_id IS NULL AND (source_key=? OR system_name=?) ORDER BY CASE WHEN source_key=? THEN 0 ELSE 1 END,sort_no LIMIT 1",(system,system,system)).fetchone()
            if not row:
                # 新增子系统也必须建立工作包；保留现有编码和历史实际流水。
                codes={r[0] for r in con.execute('SELECT wbs_code FROM wbs_tasks')}
                number=1
                while '%02d' % number in codes:number+=1
                tid=str(uuid.uuid4());now=_now()
                sort_no=con.execute('SELECT COALESCE(MAX(sort_no),0)+1 FROM wbs_tasks').fetchone()[0]
                con.execute('''INSERT INTO wbs_tasks
                    (id,wbs_code,system_name,task_name,process_group,domain,sort_no,created_at,updated_at)
                    VALUES(?,?,?,?,?,?,?,?,?)''',
                    (tid,'%02d' % number,system,system,'执行','项目范围管理',sort_no,now,now))
                row=con.execute('SELECT * FROM wbs_tasks WHERE id=?',(tid,)).fetchone()
                _log(con,'wbs_task',tid,'新增子系统同步工作包',{},dict(row),'自动同步')
            qty=float(data.get('plan_qty') or 0);actual=float(data.get('actual_qty') or 0)
            status=str(data.get('status') or ('已完成' if qty and actual>=qty else ('施工中' if actual else '未开始')))
            display_name=str(data.get('display_name') or system)
            desired={'system_name':display_name,'task_name':display_name,'owner':str(data.get('owner') or ''),'plan_start':str(data.get('plan_start') or ''),'plan_finish':str(data.get('plan_finish') or ''),'plan_qty':qty,'actual_qty':actual,'unit':str(data.get('unit') or ''),'status':status,'source_type':'subsystem','source_key':system}
            if all(row[k]==v for k,v in desired.items()):continue
            con.execute('''UPDATE wbs_tasks SET system_name=?,task_name=?,owner=?,plan_start=?,plan_finish=?,plan_qty=?,actual_qty=?,unit=?,status=?,source_type='subsystem',source_key=?,last_synced_at=?,updated_at=? WHERE id=?''',
                        (display_name,display_name,str(data.get('owner') or ''),str(data.get('plan_start') or ''),str(data.get('plan_finish') or ''),qty,actual,str(data.get('unit') or ''),status,system,_now(),_now(),row['id']))
            updated+=1
        con.commit();return {'ok':True,'updated':updated}
    finally:con.close()


def record_actual(base_dir, proj, project_name, systems, task_id, work_date, qty, note='', operator=''):
    con = ensure(base_dir, proj, project_name, systems)
    task = _dict(con.execute('SELECT * FROM wbs_tasks WHERE id=? AND deleted=0', (task_id,)).fetchone())
    if not task: con.close(); return {'ok': False, 'err': '未找到 WBS 任务'}
    if task.get('source_type')=='subsystem':
        con.close(); return {'ok':False,'err':'该任务由子系统进度自动同步，请在子系统页面录入实际完成量'}
    if con.execute('SELECT 1 FROM wbs_tasks WHERE parent_id=? AND deleted=0 LIMIT 1',(task_id,)).fetchone():
        con.close(); return {'ok':False,'err':'汇总节点不能直接录入实际量，请在下级工作包或活动中录入'}
    try: qty = float(qty); datetime.date.fromisoformat(work_date)
    except Exception: con.close(); return {'ok': False, 'err': '完成量必须是数字'}
    if not math.isfinite(qty) or qty < 0: con.close(); return {'ok': False, 'err': '完成量不能为负数'}
    before = dict(task)
    con.execute('INSERT INTO actual_logs VALUES(?,?,?,?,?,?,?)', (str(uuid.uuid4()), task_id, work_date, qty, note, operator, _now()))
    actual = con.execute('SELECT COALESCE(SUM(qty),0) FROM actual_logs WHERE task_id=?', (task_id,)).fetchone()[0]
    total = float(task.get('plan_qty') or 0)
    status = '已完成' if total and actual >= total else ('施工中' if actual else '未开始')
    con.execute('UPDATE wbs_tasks SET actual_qty=?,status=?,updated_at=? WHERE id=?', (actual, status, _now(), task_id))
    after = _dict(con.execute('SELECT * FROM wbs_tasks WHERE id=? AND deleted=0', (task_id,)).fetchone())
    _log(con, 'wbs_task', task_id, '录入实际完成量', before, after, operator)
    con.commit(); con.close()
    return {'ok': True, 'msg': '实际完成量已记录；WBS 与进度统计已更新。'}


def task_actual_logs(base_dir, proj, project_name, systems, task_id):
    """返回单个工作包的实际完成流水，供工作包归档导出。"""
    con = ensure(base_dir, proj, project_name, systems)
    try:
        return [_dict(r) for r in con.execute(
            'SELECT work_date,qty,note,operator,created_at FROM actual_logs WHERE task_id=? ORDER BY work_date,created_at',
            (task_id,)
        ).fetchall()]
    finally:
        con.close()


def remove_subsystem_task(base_dir, proj, project_name, systems, source_key, operator=''):
    con=ensure(base_dir,proj,project_name,systems)
    try:
        rows=con.execute("SELECT * FROM wbs_tasks WHERE deleted=0 AND source_type='subsystem' AND source_key=?",(source_key,)).fetchall()
        for row in rows:
            before=_dict(row);con.execute('UPDATE wbs_tasks SET deleted=1,updated_at=? WHERE id=?',(_now(),row['id']))
            _log(con,'wbs_task',row['id'],'删除子系统同步工作包',before,{},operator)
        con.commit();return len(rows)
    finally:con.close()


def snapshot(base_dir, proj, project_name, systems):
    con = ensure(base_dir, proj, project_name, systems)
    p = _dict(con.execute('SELECT * FROM project_profile WHERE project_key=?', (proj,)).fetchone())
    try: p['custom_quantities'] = json.loads(p.get('custom_quantities') or '{}')
    except Exception: p['custom_quantities'] = {}
    tasks = [_dict(r) for r in con.execute('SELECT * FROM wbs_tasks WHERE deleted=0 ORDER BY sort_no,wbs_code').fetchall()]
    by_id={x['id']:x for x in tasks};children={x['id']:[] for x in tasks}
    for x in tasks:
        if x.get('parent_id') in children: children[x['parent_id']].append(x['id'])
    leaves=[x for x in tasks if not children[x['id']]]
    def descendants(tid):
        if not children.get(tid): return [by_id[tid]]
        result=[]
        for cid in children[tid]: result.extend(descendants(cid))
        return result
    for x in tasks:
        ds=descendants(x['id']);units={str(d.get('unit') or '') for d in ds if d.get('unit')}
        x['level']=str(x.get('wbs_code') or '').count('.')+1
        x['children_count']=len(children[x['id']]);x['is_summary']=bool(children[x['id']])
        if x.get('source_type')=='subsystem':
            x['display_plan_qty']=float(x.get('plan_qty') or 0);x['display_actual_qty']=float(x.get('actual_qty') or 0);x['display_unit']=x.get('unit') or ''
        elif children[x['id']]:
            x['display_plan_qty']=round(sum(float(d.get('plan_qty') or 0) for d in ds),2)
            x['display_actual_qty']=round(sum(float(d.get('actual_qty') or 0) for d in ds),2)
            x['display_unit']=next(iter(units)) if len(units)==1 else ('多单位' if units else '')
            x['status']='已完成' if x['display_plan_qty'] and x['display_actual_qty']>=x['display_plan_qty'] else ('施工中' if x['display_actual_qty'] else '未开始')
        else:
            x['display_plan_qty']=float(x.get('plan_qty') or 0);x['display_actual_qty']=float(x.get('actual_qty') or 0);x['display_unit']=x.get('unit') or ''
    total = sum(float(x.get('plan_qty') or 0) for x in leaves)
    actual = sum(float(x.get('actual_qty') or 0) for x in leaves)
    today = datetime.date.today().isoformat()
    start, finish = p.get('plan_start') or '', p.get('plan_finish') or ''
    plan_rate = 0
    if start and finish:
        try:
            s, e, t = datetime.date.fromisoformat(start[:10]), datetime.date.fromisoformat(finish[:10]), datetime.date.fromisoformat(today)
            plan_rate = min(100, max(0, (t-s).days / max((e-s).days, 1) * 100))
        except ValueError: pass
    # 任务自己的计划日期优先；未设置时使用项目日期。
    planned_qty = 0
    weighted_plan=0;weighted_actual=0;total_weight=0
    for task in leaves:
        ts = task.get('plan_start') or start
        te = task.get('plan_finish') or finish
        qty=float(task.get('plan_qty') or 0);weight=float(task.get('weight') or 1)
        elapsed=0
        if ts and te:
            try:
                a, b, now = datetime.date.fromisoformat(ts), datetime.date.fromisoformat(te), datetime.date.fromisoformat(today)
                elapsed = min(1, max(0, ((now-a).days+1) / max((b-a).days+1, 1)))
                planned_qty += qty * elapsed
            except ValueError: pass
        if qty>0 and weight>0:
            total_weight+=weight;weighted_plan+=elapsed*weight;weighted_actual+=min(1,float(task.get('actual_qty') or 0)/qty)*weight
    plan_rate = weighted_plan / total_weight * 100 if total_weight else 0
    actual_rate = weighted_actual / total_weight * 100 if total_weight else 0
    unit_totals={}
    for x in leaves:
        unit=x.get('unit') or '未设置';bucket=unit_totals.setdefault(unit,{'plan_qty':0,'actual_qty':0,'task_count':0})
        bucket['plan_qty']+=float(x.get('plan_qty') or 0);bucket['actual_qty']+=float(x.get('actual_qty') or 0);bucket['task_count']+=1
    for bucket in unit_totals.values():
        bucket['plan_qty']=round(bucket['plan_qty'],2);bucket['actual_qty']=round(bucket['actual_qty'],2);bucket['rate']=round(bucket['actual_qty']/bucket['plan_qty']*100,1) if bucket['plan_qty'] else 0
    dashboard = {'task_count':len(tasks),'work_package_count':len(leaves), 'total_qty':round(total,2), 'actual_qty':round(actual,2), 'remaining_qty':round(max(total-actual,0),2), 'actual_rate':round(actual_rate,1), 'plan_rate':round(plan_rate,1), 'variance':round(actual_rate-plan_rate,1), 'status':'正常' if actual_rate >= plan_rate else '进度落后','unit_totals':unit_totals}
    groups = [_dict(r) for r in con.execute('SELECT * FROM process_groups ORDER BY sort_no').fetchall()]
    domains = [_dict(r) for r in con.execute('SELECT * FROM knowledge_domains ORDER BY sort_no').fetchall()]
    processes = [_dict(r) for r in con.execute('SELECT * FROM process_library ORDER BY process_group,domain,name').fetchall()]
    history = [_dict(r) for r in con.execute('SELECT * FROM change_history ORDER BY created_at DESC LIMIT 30').fetchall()]
    con.close()
    return {'ok':True,'profile':p,'dashboard':dashboard,'tasks':tasks,'groups':groups,'domains':domains,'processes':processes,'history':history}


def update_task(base_dir, proj, project_name, systems, task_id, data, operator=''):
    con = ensure(base_dir, proj, project_name, systems)
    try:
        before = _dict(con.execute('SELECT * FROM wbs_tasks WHERE id=? AND deleted=0', (task_id,)).fetchone())
        if not before: return {'ok': False, 'err': '未找到 WBS 任务'}
        allowed=['plan_start','plan_finish','plan_qty','unit','task_name','owner','predecessor','task_type','description','deliverable','acceptance_criteria','weight']
        vals = {k: data.get(k, before.get(k)) for k in allowed}
        for k in ['plan_start','plan_finish']:
            vals[k] = str(vals[k] or '')
            if vals[k]: datetime.date.fromisoformat(vals[k])
        if vals['plan_start'] and vals['plan_finish'] and vals['plan_finish'] < vals['plan_start']:
            return {'ok': False, 'err': '计划完成不能早于计划开始'}
        vals['plan_qty'] = float(vals['plan_qty'] or 0)
        vals['weight'] = float(vals['weight'] or 0)
        if not math.isfinite(vals['plan_qty']) or vals['plan_qty'] < 0 or not math.isfinite(vals['weight']) or vals['weight'] < 0:
            return {'ok': False, 'err': '计划量和权重必须是非负有限数字'}
        if not str(vals['task_name'] or '').strip(): return {'ok': False, 'err': '任务名称不能为空'}
        if vals['task_type'] not in TASK_TYPES:return {'ok':False,'err':'任务类型无效'}
        actual = float(before['actual_qty'] or 0)
        status = '已完成' if vals['plan_qty'] > 0 and actual >= vals['plan_qty'] else ('施工中' if actual else '未开始')
        con.execute('UPDATE wbs_tasks SET plan_start=?,plan_finish=?,plan_qty=?,unit=?,task_name=?,owner=?,predecessor=?,task_type=?,description=?,deliverable=?,acceptance_criteria=?,weight=?,status=?,updated_at=? WHERE id=?',
                    [vals[k] for k in ['plan_start','plan_finish','plan_qty','unit','task_name','owner','predecessor','task_type','description','deliverable','acceptance_criteria','weight']]+[status,_now(),task_id])
        after = _dict(con.execute('SELECT * FROM wbs_tasks WHERE id=?', (task_id,)).fetchone())
        _log(con, 'wbs_task', task_id, '修改计划参数', before, after, operator)
        con.commit()
        return {'ok': True, 'msg': '计划已更新，统计已联动'}
    except (ValueError, TypeError):
        return {'ok': False, 'err': '请输入有效的日期和计划量'}
    finally: con.close()


def delete_task(base_dir, proj, project_name, systems, task_id, operator=''):
    con = ensure(base_dir, proj, project_name, systems)
    try:
        before = _dict(con.execute('SELECT * FROM wbs_tasks WHERE id=? AND deleted=0', (task_id,)).fetchone())
        if not before: return {'ok': False, 'err': '未找到 WBS 任务'}
        con.execute('UPDATE wbs_tasks SET deleted=1,updated_at=? WHERE id=?', (_now(),task_id))
        _log(con, 'wbs_task', task_id, '删除计划任务（保留实际记录）', before, {'deleted':1}, operator)
        con.commit()
        return {'ok': True, 'msg': '任务已移出计划与统计，历史实际记录已保留'}
    finally: con.close()
