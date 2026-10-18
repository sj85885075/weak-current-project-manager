# -*- coding: utf-8 -*-
"""
BG / 泰坦IDC 项目管理系统集成小工具 · 本地服务
启动: python server.py  (或双击 启动工具.bat)
功能: 提供前端 UI 数据读写接口, 把 UI 输入同步到各项目 29/30/31/32/05 号 Excel 文件(保留格式)
支持项目: bg=BG弱电改造项目, titan=泰坦IDC数据中心光缆施工工程
"""
import io, sys, os, json, datetime, threading, time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs, quote
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from subsys_defs import SUBSYSTEMS, SUBS_MAP
import titan_defs
import org_deliver
import pm_core
import management
import point_map
import runtime_cache
import http_delivery
import point_statistics
import point_import
import materials
import construction_flow
from auth import Auth

# 兼容 PyInstaller 打包：frozen 模式用 exe 所在目录
if getattr(sys, 'frozen', False):
    BASE = os.path.dirname(sys.executable)
    # noconsole 打包下 stdout/stderr 为 None，重定向到日志文件，避免 print 崩溃
    _logf = open(os.path.join(BASE, 'tool_run.log'), 'a', encoding='utf-8')
    sys.stdout = _logf
    sys.stderr = _logf
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass
DIR = os.path.dirname(BASE)  # 文件体系根目录

PORT = int(os.environ.get('BG_TOOL_PORT', '8765'))
AUTH = Auth(BASE)

HDR_FILL = PatternFill('solid', fgColor='1F3A5F')
BODY_FONT = Font(name='微软雅黑', size=9)
HDR_FONT = Font(name='微软雅黑', size=9, bold=True, color='FFFFFF')
BORDER = Border(*[Side(style='thin', color='B8C4D0')] * 4)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
WRAP = Alignment(vertical='center', wrap_text=True)

STATUS_LIST = ['未开始', '图纸审核', '材料送审', '待采购', '采购中', '待施工', '施工中', '调试中', '试运行', '验收中', '已完成', '暂停']
SUBS = [s['name'] for s in SUBSYSTEMS]
BG_MILES = [
    ('M1', '合同签订', 'D+0'), ('M2', '提交送审计划表', 'D+5'), ('M3', '提交样品清单', 'D+7'),
    ('M4', '长交期设备下单', 'D+15'), ('M5', '深化设计批准', 'D+35'), ('M6', '全部设备到货', 'D+60'),
    ('M7', '布线完成', 'D+60'), ('M8', '设备安装完成', 'D+72'), ('M9', '系统联调完成', 'D+82'),
    ('M10', '试运行开始（2自然月）', 'D+88'), ('M11', '实施交付', 'D+90（90工作日）'),
    ('M12', '试运行结束/终验', '交付后2自然月'),
]
SUBS_OF = {
    'bg': SUBS,
    'titan': [s['name'] for s in titan_defs.SUBSYSTEMS],
}
PROJ_NAMES = {'bg': 'BG弱电改造项目', 'titan': titan_defs.PROJ_NAME}
CUSTOM_PROJECTS_FILE = os.path.join(BASE, '自定义项目.json')

def load_custom_projects():
    try:
        with open(CUSTOM_PROJECTS_FILE,'r',encoding='utf-8') as f:data=json.load(f)
        return data if isinstance(data,list) else []
    except (OSError,json.JSONDecodeError):return []

for _project in load_custom_projects():
    _key=str(_project.get('key') or '')
    if _key and _key not in SUBS_OF:
        SUBS_OF[_key]=list(_project['subsystems']) if isinstance(_project.get('subsystems'),list) else list(SUBS)
        PROJ_NAMES[_key]=str(_project.get('name') or _key)
BASE_SUBS = {proj: tuple(subs) for proj, subs in SUBS_OF.items()}
DATA_LOCK = threading.RLock()
BOOTSTRAP_CACHE=runtime_cache.SnapshotCache(limit=8,ttl=30)
PM_CACHE=runtime_cache.SnapshotCache(limit=8,ttl=30)

def project_root(proj):
    if proj=='titan':return os.path.join(DIR,'泰坦IDC')
    if proj=='bg':return DIR
    return os.path.join(DIR,'自定义项目',proj)

def project_list():
    return [{'key':k,'name':PROJ_NAMES[k],'custom':k not in ('bg','titan')} for k in SUBS_OF]

def create_project(data):
    import re,shutil,uuid
    name=str(data.get('name') or '').strip()
    if not name:raise ValueError('项目名称不能为空')
    if any(v==name for v in PROJ_NAMES.values()):raise ValueError('项目名称已存在')
    key='p_'+uuid.uuid4().hex[:10]
    root=project_root(key);os.makedirs(root,exist_ok=False)
    try:
        for code in ('29_项目人员管理表.xlsx','30_项目进度工程量跟踪表.xlsx','31_子系统状态总览表.xlsx','32_子系统施工步骤与明细表.xlsx','05_项目问题及待办事项.xlsx','34_施工进度检查表.xlsx'):
            src=os.path.join(DIR,code)
            if os.path.exists(src):shutil.copy2(src,os.path.join(root,code))
        # 仅复用工作簿结构和格式，新项目不继承任何业务记录或子系统。
        for code in ('29_项目人员管理表.xlsx','30_项目进度工程量跟踪表.xlsx','31_子系统状态总览表.xlsx','32_子系统施工步骤与明细表.xlsx','05_项目问题及待办事项.xlsx','34_施工进度检查表.xlsx'):
            path=os.path.join(root,code)
            if not os.path.exists(path):continue
            book=load_workbook(path)
            for sheet in book.worksheets:
                if sheet.max_row>1:sheet.delete_rows(2,sheet.max_row-1)
            book.save(path);book.close()
        projects=load_custom_projects();entry={'key':key,'name':name,'subsystems':[],'created_at':datetime.datetime.now().isoformat(timespec='seconds')};projects.append(entry)
        tmp=CUSTOM_PROJECTS_FILE+'.tmp'
        with open(tmp,'w',encoding='utf-8') as f:json.dump(projects,f,ensure_ascii=False,indent=2)
        os.replace(tmp,CUSTOM_PROJECTS_FILE)
        SUBS_OF[key]=[];PROJ_NAMES[key]=name;BASE_SUBS[key]=tuple()
        profile={'project_no':str(data.get('project_no') or ''),'project_type':str(data.get('project_type') or '系统集成项目'),'plan_start':str(data.get('plan_start') or ''),'plan_finish':str(data.get('plan_finish') or '')}
        pm_core.save_profile(DIR,key,name,SUBS_OF[key],profile,'创建项目')
        return {'ok':True,'msg':'项目已创建','project':{'key':key,'name':name,'custom':True}}
    except Exception:
        if os.path.isdir(root):shutil.rmtree(root,ignore_errors=True)
        raise

def refresh_subsystems():
    for proj, defaults in BASE_SUBS.items():
        custom = [name for name, cfg in sub_settings(proj).items() if cfg.get('custom') and name not in defaults]
        SUBS_OF[proj][:] = list(defaults) + custom

def titan_plan_keys():
    return ['fen','b7','b9','odf','duan','ces'] + ['custom_'+str(i) for i in range(6, len(SUBS_OF['titan']))]

BG_QTY_SUBSYSTEMS = {'网络':'网络系统','监控':'视频监控','门禁':'门禁系统','UPS/锂电':'UPS及锂电','配电/接地':'配电接地','动环':'动环监控','消防':'消防系统','机柜':'机房机柜'}

def canonical_sub(proj, name):
    name = str(name or '')
    return BG_QTY_SUBSYSTEMS.get(name, name) if proj != 'titan' else name


def pm_snapshot(proj):
    """项目管理核心库的统一只读视图。"""
    if proj not in SUBS_OF: proj = 'bg'
    source=bootstrap(proj);profile=pm_core.profile(DIR,proj,PROJ_NAMES[proj],SUBS_OF[proj])
    def plan_date(value):
        text=str(value or '').strip()
        try:return datetime.date.fromisoformat(text[:10]).isoformat()
        except ValueError:pass
        if '+' in text:
            try:
                prefix,offset=text.split('+',1);days=int(offset)
                if prefix.upper()=='T':base=titan_defs.T0
                elif prefix.upper()=='D' and profile.get('plan_start'):base=datetime.date.fromisoformat(str(profile['plan_start'])[:10])
                else:return ''
                return (base+datetime.timedelta(days=days)).isoformat()
            except (ValueError,TypeError):return ''
        return ''
    source_map={}
    settings=source.get('sub_settings') or {}
    for sub in SUBS_OF[proj]:
        state=(source.get('sys_status') or {}).get(sub,{})
        rows=[x for x in source.get('qty',[]) if x.get('sub')==sub]
        metric=str(state.get('metric') or '')
        primary=next((x for x in rows if str(x.get('item') or '')==metric),None)
        if primary is None and rows:primary=max(rows,key=lambda x:float(x.get('total') or 0))
        primary=primary or {}
        source_map[sub]={'display_name':(settings.get(sub) or {}).get('name') or sub,'plan_qty':primary.get('total') or state.get('total') or 0,'actual_qty':primary.get('cum') or state.get('cum') or 0,'unit':primary.get('unit') or state.get('unit') or '','owner':state.get('resp') or primary.get('resp') or '','plan_start':plan_date(state.get('ps')),'plan_finish':plan_date(state.get('pe')),'status':state.get('status') or primary.get('status') or '未开始'}
    pm_core.sync_subsystem_tasks(DIR,proj,PROJ_NAMES[proj],SUBS_OF[proj],source_map)
    result=pm_core.snapshot(DIR, proj, PROJ_NAMES[proj], SUBS_OF[proj])
    for task in result.get('tasks',[]):
        if task.get('source_type')=='subsystem':
            setting=settings.get(task.get('source_key'),{})
            task.update({k:setting.get(k,'') for k in ('actual_start','actual_finish')})
    return result


_build_pm_snapshot=pm_snapshot
def pm_snapshot(proj):
    if proj in point_statistics.ACTIVE:return _build_pm_snapshot(proj)
    signature=runtime_cache.project_signature(project_root(proj),(CUSTOM_PROJECTS_FILE,))
    key=(os.path.abspath(DIR),proj)
    cached=PM_CACHE.get(key,signature)
    if cached is not None:return cached
    result=_build_pm_snapshot(proj)
    after=runtime_cache.project_signature(project_root(proj),(CUSTOM_PROJECTS_FILE,))
    return PM_CACHE.put(key,after,result) if after==signature else result


def _paths(proj):
    """各项目数据文件路径（两项目完全独立）"""
    if proj == 'titan':
        d = os.path.join(DIR, '泰坦IDC')
        return {k: os.path.join(d, f) for k, f in {
            'p29': '29_项目人员管理表.xlsx', 'p30': '30_项目进度工程量跟踪表.xlsx',
            'p31': '31_子系统状态总览表.xlsx', 'p32': '32_子系统施工步骤与明细表.xlsx',
            'p05': '05_项目问题及待办事项.xlsx',
            'p33': os.path.join('工程审计', '33_工程审计登记表.xlsx'),
            'p34': '34_施工进度检查表.xlsx'}.items()}
    root=project_root(proj)
    return {
        'p29': os.path.join(root, '29_项目人员管理表.xlsx'),
        'p30': os.path.join(root, '30_项目进度工程量跟踪表.xlsx'),
        'p31': os.path.join(root, '31_子系统状态总览表.xlsx'),
        'p32': os.path.join(root, '32_子系统施工步骤与明细表.xlsx'),
        'p05': os.path.join(root, '05_项目问题及待办事项.xlsx'),
        'p33': os.path.join(root, '工程审计', '33_工程审计登记表.xlsx'),
        'p34': os.path.join(root, '34_施工进度检查表.xlsx'),
    }


AUDIT_TYPES = ['甲方检查', '内部检查', '内部验收', '终验收']
AUDIT_HEADERS = ['日期', '检查类型', '问题编号', '问题描述', '图片文件', '是否已处理', '处理时间', '处理说明', '登记人', '登记时间']
DAILY_HEADERS = ['记录编号', '日期', '类型', '子系统', '工作内容', '完成量', '单位', '结果', '责任人', '位置及备注', '关联报告', '创建时间']
RISK_HEADERS = ['风险编号','所属子系统','风险描述','风险类别','发生概率','影响程度','风险等级','触发条件','应对策略','应对措施','责任人','计划复查日期','状态','转化问题编号','经验总结','影响区域/范围','影响估算','估算单位/类型','风险识别日期']
TODO_HEADERS = ['编号','事项','优先级','责任人','截止时间','状态','备注','是否已处理','处理方式','所属子系统','影响区域/范围']

def _daily_path(proj):
    root = project_root(proj)
    return os.path.join(root, '35_项目每日记录表.xlsx')

def ensure_daily_file(proj):
    p = _daily_path(proj)
    if os.path.exists(p): return p
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = '每日记录'; rows = []
    if proj == 'titan':
        rows = [['DAY-OTDR-20260911', '2026-09-11', '测试', '光纤系统', '288芯 OTDR 光纤测试', 142, '芯/链路', '完成', '测试工程师', 'OS2 单模；142 条全部 PASS，失败 0、预警 0；累计 43,213 ft。ODF：CGK3A_2F-07-12-48_ODF4U-LC288；CGK3A_7F-07-03-33_ODF4U-LC288。', '288芯测试报告/288芯OTDR测试报告_2026-09-11.pdf', '2026-09-11 19:50:18']]
    rebuild_sheet(ws, DAILY_HEADERS, rows, [18,12,10,15,30,12,12,10,12,62,42,20], center_cols={1,2,3,4,6,7,8,9,12})
    wb.save(p); return p

def load_daily(proj):
    wb = load_workbook(ensure_daily_file(proj))
    return {'ok': True, 'records': read_rows(wb['每日记录'], len(DAILY_HEADERS), date_col=-1), 'headers': DAILY_HEADERS}

def save_daily(data):
    proj = data.get('proj', 'bg'); item = data.get('record') or {}
    if proj not in SUBS_OF: proj = 'bg'
    if not item.get('date') or not item.get('title'): return {'ok': False, 'err': '日期和工作内容不能为空'}
    import math
    datetime.date.fromisoformat(item['date'])
    qty=float(item.get('qty') or 0)
    if not math.isfinite(qty) or qty<0:return {'ok':False,'err':'完成量必须是非负数字'}
    item['qty']=qty
    wb = load_workbook(ensure_daily_file(proj)); ws = wb['每日记录']; no = 'DAY-' + datetime.datetime.now().strftime('%Y%m%d%H%M%S%f')
    vals = [no,item.get('date'),item.get('type') or '施工',item.get('system') or '其他',item.get('title'),item.get('qty') or 0,item.get('unit') or '',item.get('result') or '完成',item.get('owner') or '',item.get('note') or '',item.get('report') or '',datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
    r = ws.max_row + 1
    for c,v in enumerate(vals,1):
        cell=ws.cell(row=r,column=c,value=v); cell.font=BODY_FONT; cell.border=BORDER; cell.alignment=CENTER if c in {1,2,3,4,6,7,8,9,12} else WRAP
    wb.save(_daily_path(proj)); return {'ok': True, 'msg': '每日记录已同步至 35_项目每日记录表.xlsx'}


def ensure_audit_file(proj):
    """确保工程审计登记表存在（不存在则创建）"""
    p = _paths(proj)['p33']
    if os.path.exists(p):
        return p
    os.makedirs(os.path.dirname(p), exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '审计登记'
    rebuild_sheet(ws, AUDIT_HEADERS, [], [12, 12, 12, 40, 30, 11, 12, 30, 10, 18],
                  center_cols={1, 2, 3, 6, 7, 9, 10})
    wb.save(p)
    return p


def audit_dir(proj, date=None):
    """审计原始文件目录（图片按日期分子文件夹）"""
    root = os.path.join(project_root(proj), '工程审计')
    if date:
        root = os.path.join(root, str(date))
    os.makedirs(root, exist_ok=True)
    return root


def save_audit(data):
    """追加一条审计问题登记（含图片文件清单）"""
    proj = data.get('proj', 'bg')
    p = ensure_audit_file(proj)
    wb = load_workbook(p)
    ws = wb['审计登记']
    rows = read_rows(ws, len(AUDIT_HEADERS), date_col=-1)
    max_no = 0
    for r in rows:
        v = str(r[2] or '')
        if v.startswith('AUD-'):
            try:
                max_no = max(max_no, int(v.split('-')[1]))
            except Exception:
                pass
    date = data.get('date') or datetime.date.today().strftime('%Y-%m-%d')
    atype = data.get('type') or '内部检查'
    desc = data.get('desc') or ''
    files = data.get('files') or []
    handled = '已处理' if data.get('handled') else '未处理'
    htime = data.get('handle_time') or ''
    hnote = data.get('handle_note') or ''
    who = data.get('who') or ''
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    max_no += 1
    no = 'AUD-%03d' % max_no
    r = ws.max_row + 1
    row = [date, atype, no, desc, ';'.join(files), handled, htime, hnote, who, now]
    for c, v in enumerate(row, 1):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = BODY_FONT; cell.border = BORDER
        cell.alignment = CENTER if c in (1, 2, 3, 6, 7, 9, 10) else WRAP
    wb.save(p)
    return {'ok': True, 'no': no, 'msg': '审计问题 %s 已登记' % no}


def save_audit_img(data):
    """保存审计图片（base64）到 工程审计/{date}/ 目录"""
    proj = data.get('proj', 'bg')
    date = data.get('date') or datetime.date.today().strftime('%Y-%m-%d')
    atype = data.get('type') or '内部检查'
    name = data.get('name') or 'photo.png'
    b64 = data.get('data') or ''
    import base64
    try:
        raw = base64.b64decode(b64)
    except Exception as e:
        return {'ok': False, 'err': '图片数据解码失败: %s' % e}
    if len(raw) > 20 * 1024 * 1024:
        return {'ok': False, 'err': '图片超过 20MB 限制'}
    d = audit_dir(proj, date)
    ext = os.path.splitext(name)[1].lower() or '.png'
    if ext not in ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'):
        ext = '.png'
    ts = datetime.datetime.now().strftime('%H%M%S')
    fname = '%s_%s_%s%s' % (atype, ts, os.path.splitext(name)[0][:24].replace(' ', '_'), ext)
    fpath = os.path.join(d, fname)
    with open(fpath, 'wb') as f:
        f.write(raw)
    return {'ok': True, 'file': fname, 'path': os.path.relpath(fpath, audit_dir(proj)), 'msg': '图片已保存'}


def _titan_days(F):
    """泰坦 15 个施工日日期（优先读 30号每日计划 sheet，缺省推算）"""
    days = []
    if os.path.exists(F['p30']):
        wb = load_workbook(F['p30'])
        if '每日计划' in wb.sheetnames:
            ws = wb['每日计划']
            for r in range(2, ws.max_row + 1):
                if str(ws.cell(row=r, column=1).value or '').startswith('第'):
                    d = ws.cell(row=r, column=2).value
                    days.append(str(d) if d else '')
    if len(days) != 15:
        days = [(titan_defs.T0 + datetime.timedelta(days=i)).strftime('%Y-%m-%d') for i in range(1, 16)]
    return days


def _cur_titan_day(days):
    """当前施工日：今天<开工日→0(未开工)；否则=已进入的施工日序号(1..15)"""
    try:
        today = datetime.date.today()
        d0 = datetime.datetime.strptime(days[0], '%Y-%m-%d').date()
        if today < d0:
            return 0
        cur = 0
        for i, ds in enumerate(days):
            if today >= datetime.datetime.strptime(ds, '%Y-%m-%d').date():
                cur = i + 1
        return cur
    except Exception:
        return 0


def update_progress_check(proj):
    """泰坦：按 30号每日计量流水重算「子项级」每日/总体进度，对照表14-1 15天计划写 34 号表。
    判定按施工日（工期 D1~D15）而非日历日期：未开工施工日不判进度；已开工日按累计实际/累计计划判超前/达标/预警/落后。"""
    if proj != 'titan':
        return None
    F = _paths('titan')
    plan = {k:list(v) for k,v in titan_defs.PLAN.items()}
    wb_plan = load_workbook(F['p30'])
    if '每日计划' in wb_plan.sheetnames:
        wp = wb_plan['每日计划']
        rr = [r for r in range(2,wp.max_row+1) if str(wp.cell(r,1).value or '').startswith('第')]
        for key,col in zip(('fen','b7','b9','odf','duan','ces'),(3,5,7,9,11,13)):
            plan[key] = [float(wp.cell(r,col).value or 0) for r in rr]
    wb_plan.close()
    keys = titan_plan_keys()
    subs = SUBS_OF['titan']
    schedules={sub:sub_schedule(proj,sub) for sub in subs}
    days=sorted(set(d for v in schedules.values() for d in v))
    for sub,key in zip(subs,keys):plan[key]=[schedules[sub].get(d,0) for d in days]
    total_plan=sum(sum(plan[k]) for k in keys)
    cur_day = _cur_titan_day(days)
    # 实际量：30号每日计量 按 (日期, 子系统) 聚合
    daily = {}
    if os.path.exists(F['p30']):
        wb = load_workbook(F['p30'])
        if '每日计量' in wb.sheetnames:
            ws = wb['每日计量']
            for r in range(2, ws.max_row + 1):
                d = ws.cell(row=r, column=1).value
                sub = ws.cell(row=r, column=2).value
                v = ws.cell(row=r, column=4).value
                ds = d.strftime('%Y-%m-%d') if isinstance(d, (datetime.date, datetime.datetime)) else str(d)
                if is_date(ds) and sub:
                    try:
                        daily[(ds, str(sub))] = daily.get((ds, str(sub)), 0) + float(v or 0)
                    except Exception:
                        pass
    days=sorted(set(days)|{d for d,sub in daily if sub in subs})
    for sub,key in zip(subs,keys):plan[key]=[schedules[sub].get(d,0) for d in days]
    cur_day=_cur_titan_day(days)
    # 子项级 90 行（施工日 × 6 子系统）
    rows = []
    daily_sum = []  # 每日总量（图表/汇总用）
    for i in range(len(days)):
        d = days[i]
        day_plan_sum = 0
        day_act_sum = 0
        for j, sub in enumerate(subs):
            key = keys[j]
            p_daily = plan[key][i] if i < len(plan[key]) else 0
            act = daily.get((d, sub), 0)
            cum_p = sum(plan[key][:i + 1])
            cum_a = 0
            for k in range(i + 1):
                cum_a += daily.get((days[k], sub), 0)
            day_plan_sum += p_daily
            day_act_sum += act
            if (i + 1) <= cur_day:
                if cum_p > 0:
                    rate = round(cum_a / cum_p * 100, 1)
                    if cum_a >= cum_p:
                        st = '超前'
                    elif cum_a >= cum_p * 0.98:
                        st = '达标'
                    elif cum_a >= cum_p * 0.90:
                        st = '预警'
                    else:
                        st = '落后'
                else:
                    st = '达标' if cum_a >= 0 else '落后'
                    rate = '—'
            else:
                st = '未开工'
                rate = '—'
            rows.append([i + 1, d, sub, p_daily, act, act - p_daily,
                         cum_p, cum_a, cum_a - cum_p, rate, st])
        daily_sum.append([i + 1, d, day_plan_sum, day_act_sum, day_act_sum - day_plan_sum])
    # 总体进度检查（按当前施工日应达量）
    cum_plan_end = sum(sum(plan[k]) for k in keys)
    total_act = sum(daily.get((d, sub), 0) for d in days for sub in subs)
    total_rate = round(total_act / total_plan * 100, 2) if total_plan else 0.0
    if cur_day > 0:
        plan_to_now = 0
        for k in keys:
            plan_to_now += sum(plan[k][:cur_day])
        plan_to_now_pct = round(plan_to_now / total_plan * 100, 2) if total_plan else 0
        diff_pct = round(total_rate - plan_to_now_pct, 2)
        if total_act >= plan_to_now:
            over_st = '超前'
        elif total_act >= plan_to_now * 0.98:
            over_st = '达标'
        elif total_act >= plan_to_now * 0.90:
            over_st = '预警'
        else:
            over_st = '落后'
    else:
        plan_to_now_pct = 0.0
        diff_pct = 0.0
        over_st = '未开工'
    over_rows = [
        ['总工程量', total_plan, '单位'],
        ['累计实际完成', total_act, '单位'],
        ['总体完成率', str(total_rate) + '%', ''],
        ['当前应达（第%s施工日）' % (cur_day if cur_day > 0 else '—'), str(plan_to_now_pct) + '%', '按工期累计计划'],
        ['偏差', (('+' if diff_pct >= 0 else '') + str(diff_pct)) + '%', '负=落后'],
        ['判定', over_st, '超前>=100% / 达标>=98% / 预警>=90% / 落后<90%'],
        ['检查时间', datetime.datetime.now().strftime('%Y-%m-%d %H:%M'), ''],
    ]
    # 写 34 号表
    p = F['p34']
    if not os.path.exists(p):
        wb = openpyxl.Workbook()
        wb.active.title = '每日进度检查'
        wb.create_sheet('总体进度检查')
        wb.save(p)
    wb = load_workbook(p)
    ws = wb['每日进度检查']
    headers = ['施工日', '日期', '子系统', '计划当日量', '实际当日量', '当日差额', '累计计划量', '累计实际量', '累计差额', '完成率%', '判定状态']
    rebuild_sheet(ws, headers, rows, [7, 11, 16, 10, 10, 10, 10, 10, 10, 8, 9], center_cols=set(range(1, 12)))
    ws2 = wb['总体进度检查']
    rebuild_sheet(ws2, ['检查项', '数值', '说明'], over_rows, [24, 20, 46], center_cols={1, 2})
    wb.save(p)
    return {'cur_day': cur_day, 'days': rows, 'daily_sum': daily_sum, 'total': over_rows}


def load_workbook(path):
    # 先把工作簿读入内存，立即释放文件句柄；Excel/WPS 短暂占用时重试。
    # 这样项目切换等只读操作不会因毫秒级保存锁而随机失败。
    for attempt in range(12):
        try:
            try:
                with open(path, 'rb') as source:
                    content = source.read()
            except PermissionError:
                if os.name != 'nt':
                    raise
                # WPS 常以允许共享读取、但不兼容 Python CRT 默认共享模式的方式打开文件。
                # 直接使用 Win32 FILE_SHARE_READ/WRITE/DELETE 获取只读快照。
                import ctypes, msvcrt
                from ctypes import wintypes
                create_file = ctypes.windll.kernel32.CreateFileW
                create_file.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                         wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
                create_file.restype = wintypes.HANDLE
                handle = create_file(os.path.abspath(path), 0x80000000, 0x7, None, 3, 0x80, None)
                if handle == ctypes.c_void_p(-1).value:
                    raise
                try:
                    fd = msvcrt.open_osfhandle(handle, os.O_RDONLY | os.O_BINARY)
                except Exception:
                    ctypes.windll.kernel32.CloseHandle(handle)
                    raise
                with os.fdopen(fd, 'rb') as source:
                    content = source.read()
            return openpyxl.load_workbook(io.BytesIO(content))
        except PermissionError as exc:
            if attempt < 11:
                time.sleep(0.5)
                continue
            name = os.path.basename(path)
            raise RuntimeError('文件“'+name+'”正被 Excel/WPS 占用，请关闭该工作簿后重试') from exc


def parse_people_import(data):
    """解析人员管理/组织架构的 Excel、CSV 批量导入文件，返回规范化对象列表。"""
    import base64, csv
    kind = str(data.get('kind') or '')
    if kind not in ('people', 'org'):
        raise ValueError('导入类型无效')
    filename = os.path.basename(str(data.get('filename') or ''))
    encoded = str(data.get('content') or '')
    if ',' in encoded:
        encoded = encoded.split(',', 1)[1]
    try:
        raw = base64.b64decode(encoded, validate=True)
    except Exception as exc:
        raise ValueError('导入文件内容无效') from exc
    if not raw or len(raw) > 8 * 1024 * 1024:
        raise ValueError('导入文件为空或超过 8MB')
    if filename.lower().endswith(('.xlsx', '.xlsm')):
        book = openpyxl.load_workbook(io.BytesIO(raw), data_only=True, read_only=True)
        sheet = book.active
        values = [list(row) for row in sheet.iter_rows(values_only=True)]
        book.close()
    elif filename.lower().endswith(('.csv', '.tsv')):
        text = None
        for encoding in ('utf-8-sig', 'gb18030'):
            try:
                text = raw.decode(encoding); break
            except UnicodeDecodeError:
                pass
        if text is None:
            raise ValueError('CSV 编码无法识别，请另存为 UTF-8 CSV')
        delimiter = '\t' if filename.lower().endswith('.tsv') or '\t' in text.splitlines()[0] else ','
        values = list(csv.reader(io.StringIO(text), delimiter=delimiter))
    else:
        raise ValueError('仅支持 .xlsx、.xlsm、.csv、.tsv 文件')
    values = [row for row in values if any(str(v or '').strip() for v in row)]
    if len(values) < 2:
        raise ValueError('导入文件至少需要表头和一行人员数据')
    aliases = {
        'people': {
            'date':['日期','到场日期'], 'name':['姓名','人员姓名'], 'pos':['职位','岗位'],
            'src':['人员来源','来源'], 'team':['所属班组','班组','分组'], 'wt':['工作类型','工种'],
            'area':['工作区域','负责区域','区域'], 'content':['工作内容','工作内容摘要'],
            'hours':['工时(h)','工时','工时（h）'], 'note':['备注']},
        'org': {
            'name':['姓名','人员姓名'], 'job':['岗位','职位'], 'group':['所属分组','分组','所属班组','班组'],
            'role':['角色','项目角色'], 'phone':['联系电话','电话','手机号'], 'area':['负责区域','工作区域'],
            'sub':['负责子系统','子系统'], 'status':['状态'], 'sort':['排序','序号'], 'note':['备注'],
            'gender':['性别'], 'trade':['工种','作业工种'], 'id_no':['证件号','身份证号'],
            'entry_date':['入场日期','进场日期'], 'bio':['个人简介','人员简介']}
    }[kind]
    def norm(value):
        return ''.join(str(value or '').strip().lower().replace('（','(').replace('）',')').split())
    lookup = {norm(alias): field for field, names in aliases.items() for alias in names}
    columns = {i: lookup[norm(value)] for i, value in enumerate(values[0]) if norm(value) in lookup}
    if 'name' not in columns.values():
        raise ValueError('未找到“姓名”列；请使用页面导出表格的表头')
    rows = []
    for source in values[1:]:
        item = {field: (source[i] if i < len(source) and source[i] is not None else '') for i, field in columns.items()}
        item = {k: (v.isoformat() if isinstance(v, (datetime.date, datetime.datetime)) else v) for k, v in item.items()}
        item['name'] = str(item.get('name') or '').strip()
        if item['name']:
            rows.append(item)
    if not rows:
        raise ValueError('没有识别到有效人员记录')
    return {'ok': True, 'rows': rows, 'count': len(rows)}


def people_import_template(kind):
    """生成人员管理或组织架构批量导入 Excel 模板。"""
    if kind == 'people':
        title = '人员管理批量导入模板'
        headers = ['日期','姓名','职位','人员来源','所属班组','工作类型','工作区域','工作内容','工时(h)','备注']
        example = [datetime.date.today().isoformat(),'张三','施工人员','中方','综合布线组','施工','A区','桥架及线缆敷设',8,'示例行，可删除']
        widths = [14,14,16,14,18,14,18,34,12,24]
        notes = ['日期格式：YYYY-MM-DD；未填写日期时，系统使用页面当前日期。','姓名为必填字段；相同日期和姓名将更新原记录。','人员来源、班组和工作类型可按项目实际填写。']
    elif kind == 'org':
        title = '人员组织架构批量导入模板'
        headers = ['姓名','性别','联系电话','工种','岗位','证件号','入场日期','所属分组','角色','负责区域','负责子系统','个人简介','状态','排序','备注']
        example = ['李四','男','13800000000','光缆熔接工','技术员','ID-001',datetime.date.today().isoformat(),'施工管理组','技术负责人','B区','网络系统','持低压电工证，负责现场施工与复核','在职',1,'示例行，可删除']
        widths = [14,10,18,18,16,20,14,20,18,18,20,36,12,10,26]
        notes = ['姓名为必填字段；同名人员将更新原记录。','照片请在人员明细窗口单独上传；支持 JPG、PNG、WEBP，单张不超过5MB。','状态建议：在职、停用；缺少的岗位和分组会自动创建。']
    else:
        raise ValueError('模板类型无效')
    book = openpyxl.Workbook()
    sheet = book.active; sheet.title = '导入数据'
    sheet.append(headers); sheet.append(example)
    for cell in sheet[1]:
        cell.font = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
        cell.fill = HDR_FILL; cell.alignment = CENTER; cell.border = BORDER
    for cell in sheet[2]:
        cell.font = Font(name='微软雅黑', size=10); cell.alignment = WRAP; cell.border = BORDER
    for i, width in enumerate(widths, 1): sheet.column_dimensions[get_column_letter(i)].width = width
    sheet.freeze_panes = 'A2'; sheet.auto_filter.ref = 'A1:'+get_column_letter(len(headers))+'2'
    guide = book.create_sheet('填写说明')
    guide.append([title]); guide.merge_cells(start_row=1,start_column=1,end_row=1,end_column=2)
    guide['A1'].font = Font(name='微软雅黑', size=14, bold=True, color='FFFFFF'); guide['A1'].fill = HDR_FILL; guide['A1'].alignment = CENTER
    guide.append(['序号','说明'])
    for i, note in enumerate(notes, 1): guide.append([i,note])
    guide.column_dimensions['A'].width=10; guide.column_dimensions['B'].width=80
    for row in guide.iter_rows(min_row=2):
        for cell in row: cell.font=Font(name='微软雅黑',size=10);cell.border=BORDER;cell.alignment=WRAP
    buf=io.BytesIO();book.save(buf);book.close()
    return buf.getvalue(), title+'.xlsx'


def export_subsystem_actions(proj, sub):
    """将当前子系统的问题、待办、风险合并到一份三工作表 Excel。"""
    if proj not in SUBS_OF:
        raise ValueError('项目无效')
    sub = str(sub or '').strip()
    data = bootstrap(proj)
    issues = [r for r in data.get('issues', []) if str((r[7] if proj == 'titan' else r[4]) or '项目级') == sub or not sub]
    todos = [r for r in data.get('todos', []) if str((r[9] if len(r)>9 else '') or '项目级') == sub or not sub]
    risks = [r for r in data.get('risks', []) if str((r[1] if len(r)>1 else '') or '项目级') == sub or not sub]
    book = openpyxl.Workbook(); issue_sheet = book.active; issue_sheet.title = '问题管理'
    if proj == 'titan':
        issue_rows = [[r[0],r[1],r[3],r[7],r[8],r[11],r[12],r[14],r[10],r[18],r[19],r[20] if len(r)>20 else ''] for r in issues]
    else:
        issue_rows = [[r[0],r[1],r[5],r[4],r[7],r[13],r[14],r[15],r[12],r[18],r[19],r[20] if len(r)>20 else ''] for r in issues]
    rebuild_sheet(issue_sheet,['编号','问题描述','分类/优先级','所属子系统','影响区域/范围','责任人','计划完成','状态','解决措施','是否已处理','处理方式','问题处理结果'],issue_rows,[16,38,16,22,28,16,14,14,32,14,28,40],{1,3,4,6,7,8,10})
    todo_sheet=book.create_sheet('待办管理')
    todo_rows=[[r[i] if i<len(r) else '' for i in (0,1,2,9,10,3,4,5,6,7,8)] for r in todos]
    rebuild_sheet(todo_sheet,['编号','待办事项','优先级','所属子系统','影响区域/范围','责任人','截止日期','状态','备注','是否已处理','处理方式'],todo_rows,[16,38,12,22,28,16,14,14,28,14,28],{1,3,4,6,7,8,10})
    risk_sheet=book.create_sheet('风险识别')
    risk_rows=[[r[i] if i<len(r) else '' for i in (0,18,2,3,1,15,16,17,4,5,6,7,8,9,10,11,12,14)] for r in risks]
    rebuild_sheet(risk_sheet,['编号','风险识别日期','风险描述','类别','所属子系统','影响区域/范围','影响估算','估算单位/类型','发生概率','影响程度','风险等级','触发条件','应对策略','应对措施','责任人','复查日期','状态','经验总结'],risk_rows,[16,14,38,14,22,28,24,18,12,12,12,28,14,32,16,14,16,32],{1,2,4,5,8,9,10,11,13,15,16,17})
    for sheet in book.worksheets:
        sheet.freeze_panes='A2';sheet.auto_filter.ref=sheet.dimensions
        sheet.sheet_properties.pageSetUpPr.fitToPage=True;sheet.page_setup.orientation='landscape';sheet.page_setup.fitToWidth=1;sheet.page_setup.fitToHeight=0
    safe=''.join('_' if c in '\\/:*?"<>|' else c for c in (sub or '每日记录_全部子系统'))[:60]
    buf=io.BytesIO();book.save(buf);book.close()
    return buf.getvalue(), safe+'_风险问题待办汇总.xlsx'


def subsystem_action_item(data):
    """修改或删除子系统问题、待办、风险；编号保持历史唯一且不复用。"""
    proj=str(data.get('proj') or 'bg');kind=str(data.get('kind') or '');action=str(data.get('action') or '')
    record_id=str(data.get('id') or '').strip();fields=data.get('fields') or {}
    if proj not in SUBS_OF: raise ValueError('项目无效')
    if kind not in ('issue','todo','risk') or action not in ('update','delete'): raise ValueError('操作类型无效')
    if not record_id: raise ValueError('记录编号不能为空')
    path=_paths(proj)['p05'];book=load_workbook(path)
    sheet_name={'issue':'问题清单','todo':'待办事项','risk':'风险登记册'}[kind]
    if sheet_name not in book.sheetnames: raise ValueError('未找到对应数据表')
    sheet=book[sheet_name];row_no=next((r for r in range(2,sheet.max_row+1) if str(sheet.cell(r,1).value or '')==record_id),None)
    if row_no is None: raise ValueError('未找到记录：'+record_id)
    if action=='delete':
        sheet.delete_rows(row_no,1)
        rows=[]
        for r in range(2,sheet.max_row+1):
            rid=str(sheet.cell(r,1).value or '')
            match=''.join(ch for ch in rid if ch.isdigit())
            rows.append((int(match or 0),r,rid))
        rows.sort(key=lambda item:item[0])
        id_map={}
        for number,(_,r,old_id) in enumerate(rows,1):
            if kind=='issue': new_id=('TISSUE-' if proj=='titan' else 'ISSUE-')+str(number).zfill(3)
            elif kind=='risk': new_id=('TRISK-' if proj=='titan' else 'RISK-')+str(number).zfill(3)
            else: new_id=('TD-T'+str(number).zfill(2)) if proj=='titan' else ('TD-'+str(number).zfill(3))
            id_map[old_id]=new_id;sheet.cell(r,1).value=new_id
        if kind=='issue' and '风险登记册' in book.sheetnames:
            risk_sheet=book['风险登记册']
            for r in range(2,risk_sheet.max_row+1):
                old=str(risk_sheet.cell(r,14).value or '')
                if old in id_map:risk_sheet.cell(r,14).value=id_map[old]
        book.save(path);book.close()
        return {'ok':True,'msg':'已删除 '+record_id}
    max_column={'issue':21,'todo':11,'risk':19}[kind]
    columns=fields.get('columns') if isinstance(fields.get('columns'),dict) else None
    if columns is not None:
        for raw_col,value in columns.items():
            try: col=int(raw_col)
            except (TypeError,ValueError): continue
            if 2 <= col <= max_column: sheet.cell(row_no,col).value=value
    else:
        if kind=='issue':
            mapping={'result':21,'desc':2,'scope':8,'owner':12 if proj=='titan' else 14,'due':13 if proj=='titan' else 15,
                     'solution':11 if proj=='titan' else 13,'status':15 if proj=='titan' else 16}
        elif kind=='todo':
            mapping={'desc':2,'priority':3,'owner':4,'due':5,'status':6,'note':7,'scope':11}
        else:
            mapping={'desc':3,'category':4,'probability':5,'impact':6,'scope':16,'estimate':17,'estimate_unit':18,'trigger':8,'strategy':9,
                     'response':10,'owner':11,'review_date':12,'status':13,'lesson':15}
        for key,col in mapping.items():
            if key in fields: sheet.cell(row_no,col).value=fields[key]
    if kind=='risk' and (columns is not None or 'probability' in fields or 'impact' in fields):
        prob=str(sheet.cell(row_no,5).value or '中');impact=str(sheet.cell(row_no,6).value or '中')
        score={'低':1,'中':2,'高':3}.get(prob,2)*{'低':1,'中':2,'高':3}.get(impact,2)
        sheet.cell(row_no,7).value='高' if score>=6 else ('中' if score>=3 else '低')
    if kind=='issue': sheet.cell(1,21).value='问题处理结果'
    book.save(path);book.close()
    return {'ok':True,'msg':'已修改 '+record_id}


def update_progress_check_bg(proj='bg'):
    """BG：按 30号「每日计量」生成「日期×子系统」进度检查 + 总体进度检查，写 BG 34 号表。
    计划基准：30号工程量总览「计划总量」列（BG 无每日分解计划，计划当日/累计计划 标记【待确认】，不编造）。
    判定：累计实际 vs 计划总量 —— >=100% 超前 / >=98% 达标 / >=90% 预警 / <90% 落后 / =0 未开工 / 无计划总量 待确认。
    """
    F = _paths(proj)
    subs = SUBS_OF[proj]
    profile=pm_core.profile(DIR,proj,PROJ_NAMES[proj],subs)
    project_start=str(profile.get('plan_start') or '')
    project_finish=str(profile.get('plan_finish') or '')
    project_days=''
    if project_start and project_finish:
        try:project_days=(datetime.date.fromisoformat(project_finish)-datetime.date.fromisoformat(project_start)).days+1
        except ValueError:project_days=''
    period_text=(project_start+' ~ '+project_finish) if project_start and project_finish else '未设置'
    # 计划总量 / 累计 按子系统聚合（30号工程量总览 QTY 行）
    plan_total = {}
    if os.path.exists(F['p30']):
        wb = load_workbook(F['p30'])
        ws = wb['工程量总览']
        for r in range(2, ws.max_row + 1):
            no = ws.cell(row=r, column=1).value
            if not no:
                continue
            sub = canonical_sub(proj, ws.cell(row=r, column=2).value)
            if not sub:
                continue
            try:
                pt = float(ws.cell(row=r, column=4).value or 0)
            except Exception:
                pt = 0
            plan_total[sub] = plan_total.get(sub, 0) + pt
    # 每日实际：30号「每日计量」(日期, 子系统) 聚合
    daily = {}
    if os.path.exists(F['p30']):
        wb = load_workbook(F['p30'])
        if '每日计量' in wb.sheetnames:
            ws = wb['每日计量']
            for r in range(2, ws.max_row + 1):
                d = ws.cell(row=r, column=1).value
                sub = ws.cell(row=r, column=2).value
                v = ws.cell(row=r, column=4).value
                ds = d.strftime('%Y-%m-%d') if isinstance(d, (datetime.date, datetime.datetime)) else str(d)
                if is_date(ds) and sub:
                    try:
                        daily[(ds, str(sub))] = daily.get((ds, str(sub)), 0) + float(v or 0)
                    except Exception:
                        pass
    schedules={sub:sub_schedule(proj,sub) for sub in subs}
    dates = sorted({d for (d, _s) in daily.keys()}|{d for v in schedules.values() for d in v})
    rows = []
    daily_sum = []
    for i, d in enumerate(dates):
        day_act_sum = 0
        for sub in subs:
            act = daily.get((d, sub), 0)
            pt = plan_total.get(sub, 0) or 0
            cum_a = sum(daily.get((dates[k], sub), 0) for k in range(i + 1))
            if pt > 0:
                rate = round(cum_a / pt * 100, 1)
                if cum_a >= pt:
                    st = '超前'
                elif cum_a >= pt * 0.98:
                    st = '达标'
                elif cum_a >= pt * 0.90:
                    st = '预警'
                elif cum_a > 0:
                    st = '落后'
                else:
                    st = '未开工'
            else:
                rate = None
                st = '待确认'
            p=schedules[sub].get(d);cp=sum(v for date,v in schedules[sub].items() if date<=d) if schedules[sub] else None
            rows.append([i+1,d,sub,p,act,act-p if p is not None else None,cp,cum_a,cum_a-cp if cp is not None else None,rate,st])
            day_act_sum += act
        daily_sum.append([i + 1, d, None, day_act_sum, None])
    total_plan = sum(plan_total.values())
    total_act = sum(daily.get((d, sub), 0) for d in dates for sub in subs)
    total_rate = round(total_act / total_plan * 100, 2) if total_plan else 0.0
    if total_plan <= 0:
        over_st = '待确认'
    elif total_act >= total_plan:
        over_st = '超前'
    elif total_act >= total_plan * 0.98:
        over_st = '达标'
    elif total_act >= total_plan * 0.90:
        over_st = '预警'
    elif total_act > 0:
        over_st = '落后'
    else:
        over_st = '未开工'
    diff_plan = round(total_act - total_plan, 2)
    # —— 子系统独立进度检查（10 子系统各一行，按 90 天工期 D+X 标注）——
    # 计划工期/单位从 31 号表（子系统状态总览：ps/pe 为 D+X~D+Y，unit 为单位）
    sys_info = {}
    if os.path.exists(F['p31']):
        wb31 = load_workbook(F['p31'])
        ws31 = wb31['子系统状态']
        for r in range(2, ws31.max_row + 1):
            name = str(ws31.cell(row=r, column=1).value or '')
            if name in subs:
                sys_info[name] = {
                    'ps': ws31.cell(row=r, column=3).value or '',
                    'pe': ws31.cell(row=r, column=4).value or '',
                    'unit': ws31.cell(row=r, column=8).value or '',
                }
    sub_rows = []
    for sub in subs:
        pt = plan_total.get(sub, 0) or 0
        cum_a = sum(daily.get((d, sub), 0) for d in dates)
        if pt > 0:
            rate = round(cum_a / pt * 100, 1)
            if cum_a >= pt:
                st = '超前'
            elif cum_a >= pt * 0.98:
                st = '达标'
            elif cum_a >= pt * 0.90:
                st = '预警'
            elif cum_a > 0:
                st = '落后'
            else:
                st = '未开工'
            remain = round(pt - cum_a, 2)
            diff = round(cum_a - pt, 2)
        else:
            rate = None
            st = '待确认'
            remain = None
            diff = None
        info = sys_info.get(sub, {})
        period = (str(info.get('ps') or '') + ' ~ ' + str(info.get('pe') or '')) if (info.get('ps') or info.get('pe')) else '【待确认】'
        sub_rows.append([sub, pt, cum_a, rate, remain, diff, st, period, info.get('unit') or '', '30号表计划总量聚合'])
    over_rows = [
        ['项目计划工期', (str(project_days)+' 自然日') if project_days else '未设置', period_text],
        ['总工程量（计划总量）', total_plan, '单位（30号表 Σ计划总量）'],
        ['累计实际完成', total_act, '单位（30号每日计量聚合）'],
        ['总体完成率', str(total_rate) + '%', ''],
        ['与计划总量偏差', (('+' if diff_plan >= 0 else '') + str(diff_plan)), '负=未完成量'],
        ['判定', over_st, '超前>=100% / 达标>=98% / 预警>=90% / 落后<90% / 0=未开工'],
        ['检查时间', datetime.datetime.now().strftime('%Y-%m-%d %H:%M'), ''],
    ]
    # 写 34 号表（与泰坦同构；BG 另含 子系统进度检查 sheet）
    p = F['p34']
    if not os.path.exists(p):
        wb = openpyxl.Workbook()
        wb.active.title = '每日进度检查'
        wb.create_sheet('总体进度检查')
        wb.create_sheet('子系统进度检查')
        wb.save(p)
    wb = load_workbook(p)
    if '子系统进度检查' not in wb.sheetnames:
        wb.create_sheet('子系统进度检查')
    ws = wb['每日进度检查']
    headers = ['日期序号', '日期', '子系统', '计划当日量', '实际当日量', '当日差额', '累计计划量', '累计实际量', '累计差额', '完成率%', '判定状态']
    rebuild_sheet(ws, headers, rows, [8, 11, 16, 10, 10, 10, 10, 10, 10, 8, 9], center_cols=set(range(1, 12)))
    ws2 = wb['总体进度检查']
    rebuild_sheet(ws2, ['检查项', '数值', '说明'], over_rows, [24, 20, 46], center_cols={1, 2})
    ws3 = wb['子系统进度检查']
    hdr3 = ['子系统', '计划总量', '累计实际', '完成率%', '剩余量', '偏差(±)', '判定', '子系统计划工期', '单位', '数据来源']
    rebuild_sheet(ws3, hdr3, sub_rows, [14, 10, 10, 8, 10, 10, 9, 18, 8, 22], center_cols=set(range(1, 10)))
    wb.save(p)
    return {'dates': dates, 'days': rows, 'daily_sum': daily_sum, 'total': over_rows,
            'total_plan': total_plan, 'total_act': total_act, 'sub': sub_rows}


def is_date(s):
    try:
        datetime.datetime.strptime(str(s), '%Y-%m-%d')
        return True
    except Exception:
        return False


def read_rows(ws, ncols, date_col=0):
    """读 sheet 全部数据行(从第2行起), 返回 list[list]; date_col>=0 时按该列过滤非日期行"""
    out = []
    for r in range(2, ws.max_row + 1):
        row = [ws.cell(row=r, column=c).value for c in range(1, ncols + 1)]
        if all(v is None or str(v).strip() == '' for v in row):
            continue
        if date_col >= 0 and not is_date(row[date_col]):
            continue
        out.append(row)
    return out


def rebuild_sheet(ws, headers, rows, widths, center_cols=None):
    """重建 sheet(保留表头样式), 用于人员明细/每日计量等流水表"""
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.border = BORDER; cell.alignment = CENTER
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = BODY_FONT; cell.border = BORDER
            cell.alignment = CENTER if (center_cols and c in center_cols) else WRAP
    # 清理多余行
    for r in range(len(rows) + 2, ws.max_row + 1):
        ws.delete_rows(len(rows) + 2)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def bootstrap(proj='bg'):
    """读取指定项目 29/30/31/32/05 当前数据"""
    refresh_subsystems()
    out = {}
    F = _paths(proj)
    subs_ok = SUBS_OF.get(proj, SUBS)
    today = datetime.date.today().strftime('%Y-%m-%d')
    point_sync_error=''
    try:point_statistics.sync(sys.modules[__name__],proj)
    except Exception as e:point_sync_error=str(e)
    refresh_subsystems()
    subs_ok=SUBS_OF.get(proj,SUBS)
    # 项目常量（前端总览/图表按 meta 渲染，bg 与 titan 相互独立）
    if proj == 'titan':
        out['meta'] = {
            'proj': 'titan', 'name': titan_defs.PROJ_NAME, 'scope': titan_defs.PROJ_SCOPE,
            'plan_days': titan_defs.PLAN_DAYS, 'today': today, 't0': str(titan_defs.T0),
            'days': 15, 'subs': SUBS_OF['titan'],
            'miles': [[m[0], m[1], m[2]] for m in titan_defs.MILES],
            'status_list': titan_defs.STATUS_LIST,
            'plan': {k: v for k, v in titan_defs.PLAN.items()},
            'banzu': [[b[0], b[1], b[2], b[3], b[4]] for b in titan_defs.BANZU],
        }
    else:
        out['meta'] = {
            'proj': proj, 'name': PROJ_NAMES[proj],
            'plan_days': '',
            'today': today, 'warranty': '质保3年', 'trial': '试运行2自然月',
            'subs': SUBS_OF[proj], 'miles': ([[m[0], m[1], m[2]] for m in BG_MILES] if proj=='bg' else []),
            'status_list': STATUS_LIST,
        }
    # 29 每日到场汇总 + 人员明细
    wb = load_workbook(F['p29'])
    ws = wb['每日到场汇总']
    att = {}
    for r in range(2, ws.max_row + 1):
        d = ws.cell(row=r, column=1).value
        if not is_date(d):
            continue
        att[str(d)] = {
            'date': str(d), 'total': ws.cell(row=r, column=3).value or 0,
            'cn': ws.cell(row=r, column=4).value or 0, 'local': ws.cell(row=r, column=5).value or 0,
            'build': ws.cell(row=r, column=6).value or 0, 'install': ws.cell(row=r, column=7).value or 0,
            'debug': ws.cell(row=r, column=8).value or 0, 'accept': ws.cell(row=r, column=9).value or 0,
            'manage': ws.cell(row=r, column=10).value or 0, 'work': ws.cell(row=r, column=11).value or '',
            'status': ws.cell(row=r, column=12).value or '未开始',
        }
    out['attendance'] = att
    ws2 = wb['人员明细']
    raw_people=read_rows(ws2,ws2.max_column,date_col=0)
    canonical=str(ws2.cell(1,4).value or '')=='人员来源'
    out['people']=[dict(zip(('date','name','pos','src','team','wt','area','content','hours','note'),r[:10])) if canonical else dict(zip(('date','name','pos','team','wt','area','content','hours','note'),r[:9])) for r in raw_people]
    # 30 工程量总览 + 每日计量（泰坦含每日计划）
    wb = load_workbook(F['p30'])
    ws = wb['工程量总览']
    qty = []
    for r in range(2, ws.max_row + 1):
        no = ws.cell(row=r, column=1).value
        if not no:
            continue
        qty.append({
            'no': str(no), 'sub': canonical_sub(proj, ws.cell(row=r, column=2).value), 'item': ws.cell(row=r, column=3).value,
            'total': ws.cell(row=r, column=5 if proj=='titan' else 4).value, 'unit': ws.cell(row=r, column=4 if proj=='titan' else 5).value,
            'today': ws.cell(row=r, column=6).value or 0, 'cum': ws.cell(row=r, column=7).value or 0,
            'plan': ws.cell(row=r, column=11).value or '', 'note': ws.cell(row=r, column=10).value or '',
            'resp': ws.cell(row=r, column=12).value or '', 'status': ws.cell(row=r, column=13).value or '待确认',
        })
    out['qty'] = qty
    ws2 = wb['每日计量']
    out['metrics'] = read_rows(ws2, 8, date_col=0)
    out['daily_plan'] = read_rows(wb['子系统每日计划'],3,date_col=0) if '子系统每日计划' in wb.sheetnames else []
    # 泰坦每日计划（读30号每日计划sheet实际值，前端堆叠图/人员趋势用）
    if proj == 'titan' and '每日计划' in wb.sheetnames:
        wplan = wb['每日计划']
        plan = {'days': [], 'fen': [], 'b7': [], 'b9': [], 'odf': [], 'duan': [], 'ces': [], 'cum': [], 'people': []}
        for r in range(2, wplan.max_row + 1):
            if str(wplan.cell(row=r, column=1).value or '').startswith('第'):
                d = wplan.cell(row=r, column=2).value or ''
                plan['days'].append(str(d) if d else ('T+'+str(len(plan['days'])+1)))
                plan['fen'].append(wplan.cell(row=r, column=3).value or 0)
                plan['b7'].append(wplan.cell(row=r, column=5).value or 0)
                plan['b9'].append(wplan.cell(row=r, column=7).value or 0)
                plan['odf'].append(wplan.cell(row=r, column=9).value or 0)
                plan['duan'].append(wplan.cell(row=r, column=11).value or 0)
                plan['ces'].append(wplan.cell(row=r, column=13).value or 0)
                plan['cum'].append(wplan.cell(row=r, column=16).value or 0)
                plan['people'].append(wplan.cell(row=r, column=17).value or 0)
        # 累计列可能是公式字符串，解析失败回退定义值
        defs = titan_defs.PLAN
        for k in ('fen','b7','b9','odf','duan','ces','cum','people'):
            for i in range(len(plan[k])):
                v = plan[k][i]
                if isinstance(v, str) or not isinstance(v, (int, float)):
                    try:
                        plan[k][i] = float(v)
                    except Exception:
                        plan[k][i] = defs[k][i] if i < len(defs[k]) else 0
        if len(plan['fen']) == 15:
            out['meta']['plan'] = plan
        else:
            out['meta']['plan'] = {k: v for k, v in titan_defs.PLAN.items()}
            out['meta']['plan']['days'] = ['T+'+str(i+1) for i in range(15)]
    # 31 子系统状态
    wb = load_workbook(F['p31'])
    ws = wb['子系统状态']
    sys_status = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not name or str(name) not in subs_ok:
            continue
        sys_status[str(name)] = {
            'status': ws.cell(row=r, column=2).value or '未开始',
            'ps': ws.cell(row=r, column=3).value or '', 'pe': ws.cell(row=r, column=4).value or '',
            'today': ws.cell(row=r, column=5).value or 0, 'cum': ws.cell(row=r, column=6).value or 0,
            'total': ws.cell(row=r, column=7).value, 'unit': ws.cell(row=r, column=8).value or '',
            'resp': ws.cell(row=r, column=12).value or '', 'time': ws.cell(row=r, column=13).value or '',
            'note': ws.cell(row=r, column=14).value or '', 'metric': ws.cell(row=r, column=15).value or '',
            'src': ws.cell(row=r, column=16).value or '',
        }
    out['sys_status'] = sys_status
    # 32 施工步骤 / 详细工程量 / 验收记录
    wb = load_workbook(F['p32'])
    ws = wb['施工步骤']
    steps = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not name or str(name) not in subs_ok:
            continue
        steps.append({
            'sub': str(name), 'no': ws.cell(row=r, column=2).value, 'step': ws.cell(row=r, column=3).value,
            'ps': ws.cell(row=r, column=4).value or '', 'pe': ws.cell(row=r, column=5).value or '',
            'status': ws.cell(row=r, column=6).value or '未开始', 'finish': ws.cell(row=r, column=7).value or '',
            'resp': ws.cell(row=r, column=8).value or '', 'note': ws.cell(row=r, column=9).value or '',
        })
    step_quantities = {}
    step_notes = {}
    if '施工步骤每日完成量' in wb.sheetnames:
        for sub, step, date, qty, note in wb['施工步骤每日完成量'].iter_rows(min_row=2, max_col=5, values_only=True):
            if sub and step and date and qty is not None:
                step_quantities.setdefault((str(sub), str(step)), {})[str(date)] = qty
            if sub and step and date and note is not None:
                step_notes.setdefault((str(sub), str(step)), {})[str(date)] = str(note)
    for step in steps:
        step['quantities'] = step_quantities.get((step['sub'], str(step['step'])), {})
        step['daily_notes'] = step_notes.get((step['sub'], str(step['step'])), {})
    out['steps'] = steps
    ws = wb['详细工程量']
    details = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not name or str(name) not in subs_ok:
            continue
        details.append({
            'sub': str(name), 'item': ws.cell(row=r, column=2).value, 'unit': ws.cell(row=r, column=3).value,
            'total': ws.cell(row=r, column=4).value, 'today': ws.cell(row=r, column=5).value or 0,
            'cum': ws.cell(row=r, column=6).value or 0, 'src': ws.cell(row=r, column=9).value or '',
            'note': ws.cell(row=r, column=10).value or '',
            'ps':str(ws.cell(row=r,column=11).value or ''), 'pe':str(ws.cell(row=r,column=12).value or ''),
            'workers':json.loads(ws.cell(row=r,column=13).value or '[]'),
        })
    out['details'] = details
    ws = wb['验收记录']
    acceptance = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not name or str(name) not in subs_ok:
            continue
        acceptance.append({
            'sub': str(name), 'type': ws.cell(row=r, column=2).value,
            'status': ws.cell(row=r, column=3).value or '未开始', 'date': ws.cell(row=r, column=4).value or '',
            'who': ws.cell(row=r, column=5).value or '', 'concl': ws.cell(row=r, column=6).value or '',
            'reason': ws.cell(row=r, column=7).value or '', 'fixdate': ws.cell(row=r, column=8).value or '',
        })
    out['acceptance'] = acceptance
    # 05 问题清单（第一列为编号非日期，需禁用日期过滤）+ 待办事项
    wb = load_workbook(F['p05'])
    ws = wb['问题清单']
    out['issues'] = read_rows(ws, 21, date_col=-1)
    if '待办事项' in wb.sheetnames:
        out['todos'] = read_rows(wb['待办事项'], len(TODO_HEADERS), date_col=-1)
    else:
        out['todos'] = []
    if '风险登记册' in wb.sheetnames:
        out['risks'] = read_rows(wb['风险登记册'], len(RISK_HEADERS), date_col=-1)
    else:
        out['risks'] = []
    # 工程审计登记表（不存在则为空）
    if os.path.exists(F.get('p33', '')):
        wb = load_workbook(F['p33'])
        out['audit'] = read_rows(wb['审计登记'], len(AUDIT_HEADERS), date_col=-1)
    else:
        out['audit'] = []
    # 施工进度检查（34号表，两项目通用；泰坦按施工日，BG 按日期，保存时自动生成）
    if os.path.exists(F.get('p34', '')):
        wb = load_workbook(F['p34'])
        if '每日进度检查' in wb.sheetnames:
            out['progress_check'] = read_rows(wb['每日进度检查'], 11, date_col=-1)
        else:
            out['progress_check'] = []
        if '总体进度检查' in wb.sheetnames:
            out['progress_total'] = read_rows(wb['总体进度检查'], 3, date_col=-1)
        else:
            out['progress_total'] = []
        if '子系统进度检查' in wb.sheetnames:
            out['progress_sub'] = read_rows(wb['子系统进度检查'], 10, date_col=-1)
        else:
            out['progress_sub'] = []
        if proj == 'titan':
            out['meta']['cur_day'] = _cur_titan_day(_titan_days(F))
    else:
        out['progress_check'] = []
        out['progress_total'] = []
        out['progress_sub'] = []
        if proj == 'titan':
            out['meta']['cur_day'] = _cur_titan_day(_titan_days(F))
    # 26 人员组织架构（两项目独立）
    try:
        org = org_deliver.load_org(proj)
        out['org'] = org
        out['org_stats'] = org_deliver.org_stats(org)
    except Exception as e:
        out['org'] = {'people': [], 'groups': [], 'jobs': []}
        out['org_stats'] = {}
    # 甲方交付文件登记（两项目独立）
    try:
        out['deliver'] = org_deliver.load_manifest(proj)
    except Exception as e:
        out['deliver'] = []
    if proj != 'titan' and out.get('daily_plan'):
        plan_map={(str(r[0])[:10],str(r[1])):float(r[2] or 0) for r in out['daily_plan'] if len(r)>=3}
        rows=out.get('progress_check',[])
        existing={(str(r[1])[:10],str(r[2])) for r in rows}
        for date,sub in plan_map:
            if (date,sub) not in existing:rows.append([0,date,sub,None,0,None,None,0,None,None,'待确认'])
        for r in rows:
            date,sub=str(r[1])[:10],str(r[2])
            if (date,sub) in plan_map:
                p=plan_map[(date,sub)];cp=sum(v for (d,s),v in plan_map.items() if s==sub and d<=date)
                r[3]=p;r[5]=float(r[4] or 0)-p;r[6]=cp;r[8]=float(r[7] or 0)-cp
        out['progress_check']=sorted(rows,key=lambda r:(str(r[1]),str(r[2])))
    out['sub_settings']=sub_settings(proj)
    for sub, setting in out['sub_settings'].items():
        if sub in out['sys_status']:
            out['sys_status'][sub].update({k:setting.get(k,'') for k in ('actual_start','actual_finish')})
    out['sub_schedules']={sub:sub_schedule(proj,sub) for sub in SUBS_OF[proj]}
    if proj=='titan':
        schedules=out['sub_schedules'];days=sorted(set(d for v in schedules.values() for d in v)|{str(r[0])[:10] for r in out['metrics'] if r[1] in SUBS_OF[proj]});running=0;grand=sum(sum(v.values()) for v in schedules.values());pl={'days':days,'cum':[],'people':[0]*len(days)}
        for sub,key in zip(SUBS_OF[proj],titan_plan_keys()):pl[key]=[schedules[sub].get(d,0) for d in days]
        for d in days:
            running+=sum(v.get(d,0) for v in schedules.values());pl['cum'].append(round(running/grand*100,2) if grand else 0)
        out['meta']['plan']=pl
        out['meta']['plan_keys']=titan_plan_keys()
        out['meta']['cur_day']=_cur_titan_day(days)
    profile=pm_core.profile(DIR,proj,PROJ_NAMES[proj],SUBS_OF[proj])
    out['meta']['project_no']=profile.get('project_no') or ''
    out['meta']['project_type']=profile.get('project_type') or ''
    out['meta']['plan_start']=profile.get('plan_start') or ''
    out['meta']['plan_finish']=profile.get('plan_finish') or ''
    if os.path.exists(os.path.join(project_root(proj),'点位图基础数据.json')):
        points=point_map.snapshot(project_root(proj))
        out['device_statistics']=points['device_statistics']
        out['pointmap_sync']=points['global_sync']
        if point_sync_error:out['pointmap_sync'].update(status='同步失败',message=point_sync_error)
    return out


_build_bootstrap=bootstrap
def bootstrap(proj='bg'):
    if proj in point_statistics.ACTIVE:return _build_bootstrap(proj)
    signature=runtime_cache.project_signature(project_root(proj),(CUSTOM_PROJECTS_FILE,))
    key=(os.path.abspath(DIR),proj)
    cached=BOOTSTRAP_CACHE.get(key,signature)
    if cached is not None:return cached
    result=_build_bootstrap(proj)
    if result.get('pointmap_sync',{}).get('status')=='同步失败':return result
    after=runtime_cache.project_signature(project_root(proj),(CUSTOM_PROJECTS_FILE,))
    return BOOTSTRAP_CACHE.put(key,after,result) if after==signature else result


def save_step_records(data):
    proj = data.get('proj', 'bg')
    if proj not in SUBS_OF:
        raise ValueError('项目无效')
    # Validate daily step quantities before writing any workbook.
    for updates in data.get('steps_updates', {}).values():
        for update in updates.values():
            for date, note in update.get('daily_notes', {}).items():
                if datetime.date.fromisoformat(date).isoformat() != date or not isinstance(note, str):
                    raise ValueError('施工步骤备注日期或内容无效')
            for date, qty in update.get('quantities', {}).items():
                if datetime.date.fromisoformat(date).isoformat() != date:
                    raise ValueError('施工步骤填报日期无效')
                if qty is not None and (isinstance(qty, bool) or not isinstance(qty, (int, float)) or not 0 <= qty < float('inf')):
                    raise ValueError('当日完成数量必须为非负有限数字')
    path = _paths(proj)['p32']
    wb = load_workbook(path)
    ws = wb['施工步骤']
    steps_upd = data.get('steps_updates', {})  # {sub: {step: {status, finish, note}}}
    daily_steps = wb['施工步骤每日完成量'] if '施工步骤每日完成量' in wb.sheetnames else wb.create_sheet('施工步骤每日完成量')
    for col, title in enumerate(['子系统', '施工步骤', '填报日期', '当日完成数量', '备注'], 1):
        daily_steps.cell(1, col, title)
    daily_index = {(str(daily_steps.cell(r, 1).value), str(daily_steps.cell(r, 2).value), str(daily_steps.cell(r, 3).value)): r for r in range(2, daily_steps.max_row + 1)}
    for sub, mp in steps_upd.items():
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value) != str(sub):
                continue
            stepname = ws.cell(row=r, column=3).value
            if stepname not in mp:
                continue
            u = mp[stepname]
            for date in dict.fromkeys([*u.get('quantities', {}), *u.get('daily_notes', {})]):
                key = (str(sub), str(stepname), date)
                row = daily_index.get(key)
                if row is None:
                    row = daily_steps.max_row + 1
                    daily_index[key] = row
                    for col, value in enumerate(key, 1):
                        daily_steps.cell(row, col, value)
                if date in u.get('quantities', {}):
                    daily_steps.cell(row, 4).value = u['quantities'][date]
                if date in u.get('daily_notes', {}):
                    daily_steps.cell(row, 5).value = u['daily_notes'][date]
                    daily_steps.cell(row, 5).data_type = 's'
            if 'status' in u:
                ws.cell(row=r, column=6, value=u['status'])
            if 'finish' in u:
                ws.cell(row=r, column=7, value=u['finish'])
            if 'note' in u:
                ws.cell(row=r, column=9, value=u['note'])
    wb.save(path)
    wb.close()
    return {'ok': True, 'msg': '施工步骤数量及备注已按填报日期保存'}


def save_payload(data):
    """把前端全量状态写回指定项目 Excel"""
    msgs = []
    proj = data.get('proj', 'bg')
    F = _paths(proj)
    date_str = data.get('today') or datetime.date.today().strftime('%Y-%m-%d')
    # Validate daily step quantities before writing any workbook.
    for updates in data.get('steps_updates', {}).values():
        for update in updates.values():
            for date, note in update.get('daily_notes', {}).items():
                if datetime.date.fromisoformat(date).isoformat() != date or not isinstance(note, str):
                    raise ValueError('施工步骤备注日期或内容无效')
            for date, qty in update.get('quantities', {}).items():
                if datetime.date.fromisoformat(date).isoformat() != date:
                    raise ValueError('施工步骤填报日期无效')
                if qty is not None and (isinstance(qty, bool) or not isinstance(qty, (int, float)) or not 0 <= qty < float('inf')):
                    raise ValueError('当日完成数量必须为非负有限数字')
    # 预读 30 号每日计量现有行（用于同日修正与 upsert，BG/泰坦共用）
    flow_rows = []  # [行号, 日期, 子系统, 明细项, 今日量, 单位, 累计量]
    try:
        wb30_pre = load_workbook(F['p30'])
        ws30_pre = wb30_pre['每日计量']
        for rr in range(2, ws30_pre.max_row + 1):
            flow_rows.append([rr] + [ws30_pre.cell(row=rr, column=c).value for c in range(1, 8)])
        wb30_pre.close()
    except Exception:
        flow_rows = []
    # ---- 29 人员管理 ----
    wb = load_workbook(F['p29'])
    ws = wb['每日到场汇总']
    att = data.get('attendance', {})
    for date, a in att.items():
        matched=False
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value) == str(date):
                matched=True
                vals = [a.get('total', 0), a.get('cn', 0), a.get('local', 0), a.get('build', 0),
                        a.get('install', 0), a.get('debug', 0), a.get('accept', 0), a.get('manage', 0),
                        a.get('work', ''), a.get('status', '未开始')]
                for c, v in enumerate(vals, 3):
                    ws.cell(row=r, column=c, value=v)
                break
        if not matched:ws.append([date,'']+[a.get(k,0) for k in ('total','cn','local','build','install','debug','accept','manage')]+[a.get('work',''),a.get('status','未开始')])
    # 人员明细：按日期替换（前端传对象数组，转二维数组写入）
    ws2 = wb['人员明细']
    people = data.get('people', [])  # 全量（dict 列表）
    people_rows = [[(p.get(k, '') if isinstance(p, dict) else p[i] if i < len(p) else '')
                    for i, k in enumerate(('date', 'name', 'pos', 'src', 'team', 'wt', 'area', 'content', 'hours', 'note'))]
                   for p in people]
    rebuild_sheet(ws2,['日期','姓名','职位','人员来源','所属班组','工作类型','工作区域','工作内容','工时(h)','备注'],people_rows,[12,12,14,12,14,12,16,34,9,24],center_cols={1,2,3,4,5,6,7,9})
    wb.save(F['p29'])
    msgs.append('29 人员管理表')
    # 人员管理 → 人员组织架构联动（新姓名自动加入 26 号表，两项目独立）
    try:
        sync_n = org_deliver.sync_org_from_people(proj, people_rows)
        if sync_n:
            msgs.append('组织架构自动同步%d人' % sync_n)
    except Exception as e:
        msgs.append('组织架构联动失败(%s)' % e)
    # ---- 30 工程量 ----
    wb = load_workbook(F['p30'])
    ws = wb['工程量总览']
    qty_upd = data.get('qty_updates', {})
    for no, u in qty_upd.items():
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value) == str(no):
                if 'today' in u:
                    ws.cell(row=r, column=6, value=u['today'])
                if 'cum' in u:
                    ws.cell(row=r, column=7, value=u['cum'])
                if 'plan' in u and u['plan'] is not None and u['plan'] != '':
                    ws.cell(row=r, column=11, value=u['plan'])
                if 'status' in u:
                    ws.cell(row=r, column=13, value=u['status'])
                break
    # 每日计量：不再全量重建（BG/泰坦统一由 qty_flows+detail_flows 按(日期,子系统,明细) upsert 维护，
    # 避免同日修改重复累加、清空(0)后流水残留导致进度检查不一致）
    qty_flows = []
    if proj == 'bg':
        for no, u in qty_upd.items():
            t = 0
            try:
                t = float(u.get('today', 0) or 0)
            except Exception:
                t = 0
            if t and t > 0:
                for r in range(2, ws.max_row + 1):
                    if str(ws.cell(row=r, column=1).value) == str(no):
                        sub = ws.cell(row=r, column=2).value or ''
                        item = '今日完成：' + str(ws.cell(row=r, column=3).value or '')
                        unit = ws.cell(row=r, column=5).value or ''
                        cum = u.get('cum', t)
                        old_row = next((fr for fr in flow_rows
                                        if str(fr[1] or '') == date_str and str(fr[2] or '') == str(sub)
                                        and str(fr[3] or '') == item), None)
                        qty_flows.append({'row': old_row[0] if old_row else None, 'date': date_str,
                                          'sub': str(sub), 'item': item,
                                          'today': t, 'unit': str(unit) or '', 'cum': cum})
                        break
    wb.save(F['p30'])
    msgs.append('30 进度工程量')
    # ---- 31 子系统状态 ----
    wb = load_workbook(F['p31'])
    ws = wb['子系统状态']
    sys_upd = data.get('sys_updates', {})
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    for name, u in sys_upd.items():
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value) == str(name):
                if 'status' in u:
                    ws.cell(row=r, column=2, value=u['status'])
                if 'today' in u:
                    ws.cell(row=r, column=5, value=u['today'])
                if 'cum' in u:
                    ws.cell(row=r, column=6, value=u['cum'])
                if 'resp' in u:
                    ws.cell(row=r, column=12, value=u['resp'])
                ws.cell(row=r, column=13, value=now)
                if 'note' in u:
                    ws.cell(row=r, column=14, value=u['note'])
                break
    wb.save(F['p31'])
    msgs.append('31 子系统状态')
    # ---- 32 施工步骤 / 详细工程量 / 验收记录 ----
    wb = load_workbook(F['p32'])
    ws = wb['施工步骤']
    steps_upd = data.get('steps_updates', {})  # {sub: {step: {status, finish, note}}}
    daily_steps = wb['施工步骤每日完成量'] if '施工步骤每日完成量' in wb.sheetnames else wb.create_sheet('施工步骤每日完成量')
    for col, title in enumerate(['子系统', '施工步骤', '填报日期', '当日完成数量', '备注'], 1):
        daily_steps.cell(1, col, title)
    daily_index = {(str(daily_steps.cell(r, 1).value), str(daily_steps.cell(r, 2).value), str(daily_steps.cell(r, 3).value)): r for r in range(2, daily_steps.max_row + 1)}
    for sub, mp in steps_upd.items():
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value) != str(sub):
                continue
            stepname = ws.cell(row=r, column=3).value
            if stepname not in mp:
                continue
            u = mp[stepname]
            for date in dict.fromkeys([*u.get('quantities', {}), *u.get('daily_notes', {})]):
                key = (str(sub), str(stepname), date)
                row = daily_index.get(key)
                if row is None:
                    row = daily_steps.max_row + 1
                    daily_index[key] = row
                    for col, value in enumerate(key, 1):
                        daily_steps.cell(row, col, value)
                if date in u.get('quantities', {}):
                    daily_steps.cell(row, 4).value = u['quantities'][date]
                if date in u.get('daily_notes', {}):
                    daily_steps.cell(row, 5).value = u['daily_notes'][date]
                    daily_steps.cell(row, 5).data_type = 's'
            if 'status' in u:
                ws.cell(row=r, column=6, value=u['status'])
            if 'finish' in u:
                ws.cell(row=r, column=7, value=u['finish'])
            if 'note' in u:
                ws.cell(row=r, column=9, value=u['note'])
    ws = wb['详细工程量']
    details_upd = data.get('details_updates', {})  # {sub: {item: {today}}}
    detail_flows = []  # 同步到每日计量的流水（同日同明细 upsert，不重复累加）
    detail_sub_today = {}   # sub -> today合计
    detail_sub_cum = {}     # sub -> cum合计
    for sub, mp in details_upd.items():
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value) != str(sub):
                continue
            item = ws.cell(row=r, column=2).value
            if item not in mp:
                continue
            u = mp[item]
            today = 0
            try:
                today = float(u.get('today', 0) or 0)
            except Exception:
                today = 0
            if today < 0:
                today = 0  # 工程量不可为负，负数按 0 处理
            prev_cum = ws.cell(row=r, column=6).value or 0
            try:
                prev_cum = float(prev_cum)
            except Exception:
                prev_cum = 0
            # 当天已填报过该明细 → 修正（替换当日值）；当天未填报 → 追加
            old_row = next((fr for fr in flow_rows
                            if str(fr[1] or '') == date_str and str(fr[2] or '') == sub and str(fr[3] or '') == item), None)
            old_today = float(old_row[4] or 0) if old_row else 0
            cum = prev_cum - old_today + today
            ws.cell(row=r, column=5, value=today)
            ws.cell(row=r, column=6, value=cum)
            if old_row:
                # 当日已有该明细流水（含 today=0 清零场景）→ 更新今日量与累计量
                detail_flows.append({
                    'row': old_row[0], 'date': date_str, 'sub': sub,
                    'item': item, 'today': today, 'unit': ws.cell(row=r, column=3).value or '',
                    'cum': cum,
                })
            elif today > 0:
                detail_flows.append({
                    'row': None, 'date': date_str, 'sub': sub,
                    'item': item, 'today': today, 'unit': ws.cell(row=r, column=3).value or '',
                    'cum': cum,
                })
            detail_sub_today[sub] = detail_sub_today.get(sub, 0) + today
            detail_sub_cum[sub] = detail_sub_cum.get(sub, 0) + cum
    ws = wb['验收记录']
    acc_upd = data.get('acceptance_updates', {})  # {sub: {type: {...}}}
    for sub, mp in acc_upd.items():
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(row=r, column=1).value) != str(sub):
                continue
            t = ws.cell(row=r, column=2).value
            if t not in mp:
                continue
            u = mp[t]
            for col, key in ((3, 'status'), (4, 'date'), (5, 'who'), (6, 'concl'), (7, 'reason'), (8, 'fixdate')):
                if key in u:
                    ws.cell(row=r, column=col, value=u[key])
    wb.save(F['p32'])
    msgs.append('32 施工步骤与明细')
    # 31 号今日/累计由明细汇总联动（覆盖前端直接传入值，保持一致）
    if detail_sub_today:
        wb31 = load_workbook(F['p31'])
        ws31 = wb31['子系统状态']
        for r in range(2, ws31.max_row + 1):
            name = str(ws31.cell(row=r, column=1).value or '')
            if name in detail_sub_today:
                ws31.cell(row=r, column=5, value=detail_sub_today.get(name, 0))
                ws31.cell(row=r, column=6, value=detail_sub_cum.get(name, 0))
        wb31.save(F['p31'])
        msgs.append('31 明细联动')
    # 泰坦：明细联动 30 号「工程量总览」今日/累计（31 号同样联动，保持三表一致）
    if proj == 'titan' and detail_sub_today:
        try:
            wb30o = load_workbook(F['p30'])
            ws30o = wb30o['工程量总览']
            for r in range(2, ws30o.max_row + 1):
                name = str(ws30o.cell(row=r, column=2).value or '')
                if name in detail_sub_today:
                    ws30o.cell(row=r, column=6, value=detail_sub_today.get(name, 0))
                    ws30o.cell(row=r, column=7, value=detail_sub_cum.get(name, 0))
            wb30o.save(F['p30'])
            msgs.append('30 总览联动')
        except Exception as e:
            msgs.append('30 总览联动失败(%s)' % e)
    # 明细流水统一 upsert 到 30 号「每日计量」（qty_flows + detail_flows）
    all_flows = qty_flows + detail_flows
    if all_flows:
        wb30 = load_workbook(F['p30'])
        ws30 = wb30['每日计量']
        for f in all_flows:
            if f['row']:
                # 当日已存在该明细 → 更新今日量与累计量（保留原样式）
                rr = f['row']
                ws30.cell(row=rr, column=4, value=f['today'])
                ws30.cell(row=rr, column=6, value=f['cum'])
            else:
                r = ws30.max_row + 1
                row = [f['date'], f['sub'], f['item'], f['today'], f['unit'], f['cum'], '工具自动', '']
                for c, v in enumerate(row, 1):
                    cell = ws30.cell(row=r, column=c, value=v)
                    cell.font = BODY_FONT; cell.border = BORDER
                    cell.alignment = CENTER if c in (1, 2, 4, 5, 6, 7) else WRAP
        wb30.save(F['p30'])
        msgs.append('30 每日计量流水')
    # ---- 05 问题清单追加 ----
    new_issues = data.get('issues_new', [])
    if new_issues:
        wb = load_workbook(F['p05'])
        ws = wb['问题清单']
        pref = 'TISSUE-' if proj == 'titan' else 'ISSUE-'
        max_no = 0
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if v and str(v).startswith(pref):
                try:
                    max_no = max(max_no, int(str(v).split('-')[1]))
                except Exception:
                    pass
        r = ws.max_row + 1
        ws.cell(1,21).value='问题处理结果'
        for ni in new_issues:
            max_no += 1
            if proj=='titan':
                row=[f'{pref}{max_no:03d}',ni.get('desc',''),ni.get('category','施工问题'),ni.get('prio','P2'),date_str,ni.get('location',''),ni.get('basis','工具登记'),ni.get('sub') or '项目级',ni.get('impact','待评估'),ni.get('resp',''),ni.get('solution','待处理'),ni.get('owner',''),ni.get('due',''),'','未关闭','否','','','未处理','']
            else:
                row = [f'{pref}{max_no:03d}', ni.get('desc', ''), ni.get('stage','执行'), ni.get('reporter','工具登记'), ni.get('sub') or '项目级',
                       ni.get('category','施工问题'), ni.get('cause','待分析'), ni.get('impact','待评估'), ni.get('schedule_impact','待评估'), ni.get('cost_impact','待评估'), ni.get('quality_impact','待评估'),
                       ni.get('resp', ''), ni.get('solution', '待处理'), ni.get('owner', ''), ni.get('due', ''),
                       '未关闭', '否', '', '未处理', '']
            row.append(ni.get('result',''))
            for c, v in enumerate(row, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = BODY_FONT; cell.border = BORDER; cell.alignment = WRAP
            r += 1
        wb.save(F['p05'])
        msgs.append('05 问题清单')
    # ---- 05 问题处理状态更新（是否已处理 / 处理方式，联动状态列） ----
    issues_upd = data.get('issues_updates', {})
    if issues_upd:
        wb = load_workbook(F['p05'])
        ws = wb['问题清单']
        ws.cell(1,21).value='问题处理结果'
        st_col = 16 if proj != 'titan' else 15  # 通用项目状态列16 / 泰坦状态列15
        for r in range(2, ws.max_row + 1):
            no = ws.cell(row=r, column=1).value
            if no and str(no) in issues_upd:
                u = issues_upd[str(no)]
                if 'handled' in u:
                    handled = '已处理' if str(u['handled']) in ('已处理', '已解决') else '未处理'
                    ws.cell(row=r, column=19, value=handled)
                    ws.cell(row=r, column=st_col, value='已关闭' if handled == '已处理' else '未关闭')
                if 'solution' in u:
                    solution_col = 13 if proj != 'titan' else 11
                    ws.cell(row=r, column=solution_col, value=u['solution'])
                    ws.cell(row=r, column=20, value=u['solution'])
                if 'result' in u:
                    ws.cell(row=r, column=21).value = u['result']
                if 'method' in u:
                    ws.cell(row=r, column=20, value=u['method'])
        wb.save(F['p05'])
        msgs.append('05 问题处理状态')
    # ---- 05 待办事项更新（状态 / 是否已处理 / 处理方式） ----
    todos_upd = data.get('todos_updates', {})
    if todos_upd:
        wb = load_workbook(F['p05'])
        if '待办事项' in wb.sheetnames:
            ws = wb['待办事项']
            for r in range(2, ws.max_row + 1):
                no = ws.cell(row=r, column=1).value
                if no and str(no) in todos_upd:
                    u = todos_upd[str(no)]
                    if 'status' in u:
                        ws.cell(row=r, column=6, value=u['status'])
                    if 'handled' in u:
                        ws.cell(row=r, column=8, value='已处理' if str(u['handled']) == '已处理' else '未处理')
                    if 'method' in u:
                        ws.cell(row=r, column=9, value=u['method'])
            wb.save(F['p05'])
            msgs.append('05 待办状态')
    # ---- 05 待办事项新增（手动添加的待办，自动编号追加） ----
    todos_new = data.get('todos_new') or []
    if todos_new:
        wb = load_workbook(F['p05'])
        if '待办事项' not in wb.sheetnames:
            ws = wb.create_sheet('待办事项')
            for c, h in enumerate(TODO_HEADERS, 1):
                ws.cell(row=1, column=c, value=h)
        ws = wb['待办事项']
        # 已有编号取最大序号
        max_no = 0
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if v:
                try:
                    num = int(''.join(ch for ch in str(v) if ch.isdigit()))
                    max_no = max(max_no, num)
                except Exception:
                    pass
        next_r = ws.max_row + 1
        for item in todos_new:
            if not item.get('item') or not str(item['item']).strip():
                continue
            max_no += 1
            no = ('TD-T' + str(max_no).zfill(2)) if proj == 'titan' else ('TD-' + str(max_no).zfill(3))
            if ws.cell(1,10).value in (None,''):ws.cell(1,10).value='所属子系统'
            if ws.cell(1,11).value in (None,''):ws.cell(1,11).value='影响区域/范围'
            vals = [no, str(item['item']).strip(), item.get('pri') or 'P2', item.get('owner') or '',
                    item.get('due') or '', '未开始', item.get('note') or '', '未处理', '',item.get('sub') or '项目级',item.get('scope') or '']
            for c, v in enumerate(vals, 1):
                ws.cell(row=next_r, column=c, value=v)
            next_r += 1
        wb.save(F['p05'])
        msgs.append('05 新增待办')
    # ---- 05 风险登记册（按子系统识别、评估、应对和复查） ----
    risks_new=data.get('risks_new') or [];risks_upd=data.get('risks_updates') or {}
    if risks_new or risks_upd:
        wb=load_workbook(F['p05'])
        if '风险登记册' not in wb.sheetnames:
            ws=wb.create_sheet('风险登记册')
            for c,h in enumerate(RISK_HEADERS,1):ws.cell(1,c).value=h
        ws=wb['风险登记册'];pref='TRISK-' if proj=='titan' else 'RISK-';max_no=0
        for r in range(2,ws.max_row+1):
            v=str(ws.cell(r,1).value or '')
            if v.startswith(pref):
                try:max_no=max(max_no,int(v.split('-')[-1]))
                except ValueError:pass
            if v in risks_upd:
                u=risks_upd[v]
                for col,key in ((9,'strategy'),(10,'response'),(11,'owner'),(12,'review_date'),(13,'status'),(14,'issue_no'),(15,'lesson')):
                    if key in u:ws.cell(r,col).value=u[key]
        for item in risks_new:
            if not str(item.get('desc') or '').strip():continue
            max_no+=1;prob=str(item.get('probability') or '中');impact=str(item.get('impact') or '中')
            score={'低':1,'中':2,'高':3}.get(prob,2)*{'低':1,'中':2,'高':3}.get(impact,2)
            level='高' if score>=6 else ('中' if score>=3 else '低')
            vals=[f'{pref}{max_no:03d}',item.get('sub') or '项目级',str(item.get('desc')).strip(),item.get('category') or '施工',prob,impact,level,item.get('trigger') or '',item.get('strategy') or '减轻',item.get('response') or '',item.get('owner') or '',item.get('review_date') or '','已识别','','',item.get('scope') or '',item.get('estimate') or '',item.get('estimate_unit') or '',item.get('identified_date') or datetime.date.today().isoformat()]
            ws.append(vals)
        for c,h in enumerate(RISK_HEADERS,1):ws.cell(1,c).value=h
        wb.save(F['p05']);msgs.append('05 风险登记册')
    # 重算每日/总体施工进度检查并写 34 号表（泰坦：表14-1 施工日计划；BG：计划总量基准，每日计划【待确认】）
    try:
        if proj == 'titan':
            update_progress_check(proj)
        else:
            update_progress_check_bg(proj)
        msgs.append('34 进度检查')
    except Exception as e:
        msgs.append('34 进度检查失败(%s)' % e)
    return {'ok': True, 'msg': '已同步：' + '、'.join(msgs)}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype='application/json; charset=utf-8'):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False).encode('utf-8')
        elif isinstance(body, str):
            body = body.encode('utf-8')
        http_delivery.send_bytes(self,code,body,ctype)

    def do_GET(self):
        if not AUTH.gate(self):
            return
        path=urlparse(self.path).path
        # Static assets and the locked CAD underlay do not wait for workbook writes.
        if path=='/api/pointmap_asset' or path=='/' or path.endswith(('.js','.css')):
            self._do_GET()
            return
        with DATA_LOCK:
            refresh_subsystems()
            self._do_GET()

    def _do_GET(self):
        u = urlparse(self.path)
        if u.path == '/api/pointmap':
            try:
                q = parse_qs(u.query)
                proj = (q.get('proj') or [''])[0]
                if proj not in SUBS_OF: raise ValueError('未知项目')
                self._send(200, point_map.snapshot(project_root(proj)))
            except (ValueError, OSError) as e: self._send(400, {'ok':False,'err':str(e)})
        elif u.path in ('/api/pointmap_asset', '/api/pointmap_dxf'):
            try:
                proj = (parse_qs(u.query).get('proj') or [''])[0]
                if proj not in SUBS_OF: raise ValueError('未知项目')
                root = project_root(proj)
                if u.path == '/api/pointmap_asset':
                    http_delivery.send_file(self,point_import.asset_path(root),'image/svg+xml; charset=utf-8')
                    return
                else:
                    content = point_map.export_dxf(root)
                    content_type = 'application/dxf'
                self.send_response(200)
                self.send_header('Content-Type', content_type)
                self.send_header('Content-Length', str(len(content)))
                self.send_header('Cache-Control', 'no-cache')
                if u.path == '/api/pointmap_dxf': self.send_header('Content-Disposition', 'attachment; filename="LV-design.dxf"')
                self.end_headers()
                self.wfile.write(content)
            except (ValueError, OSError) as e: self._send(400, {'ok':False,'err':str(e)})
        elif u.path == '/api/pointmap_import':
            try:
                q = parse_qs(u.query)
                proj = (q.get('proj') or [''])[0]
                if proj not in SUBS_OF: raise ValueError('未知项目')
                self._send(200, point_import.status(project_root(proj), (q.get('id') or [''])[0]))
            except (ValueError, OSError) as e: self._send(400, {'ok':False,'err':str(e)})
        elif u.path in ('/app_runtime.js', '/point_canvas.js'):
            self._send_file(u.path[1:], 'application/javascript; charset=utf-8')
        elif u.path == '/point_map.js':
            self._send_file('point_map.js', 'application/javascript; charset=utf-8')
        elif u.path in ('/api/materials', '/api/materials/file'):
            try:
                proj = (parse_qs(u.query).get('proj') or ['bg'])[0]
                if proj not in SUBS_OF: raise ValueError('未知项目')
                materials.handle_get(self, project_root(proj))
            except (ValueError, OSError) as e: self._send(400, {'ok':False, 'err':str(e)})
        elif u.path == '/materials.js':
            self._send_file('materials.js', 'application/javascript; charset=utf-8')
        elif u.path == '/api/management':
            try:
                proj = (parse_qs(u.query).get('proj') or ['bg'])[0]
                if proj not in SUBS_OF: raise ValueError('未知项目')
                self._send(200, management.snapshot(project_root(proj)))
            except Exception as e: self._send(400, {'ok':False, 'err':str(e)})
        elif u.path == '/management.js':
            self._send_file('management.js', 'application/javascript; charset=utf-8')
        elif u.path == '/api/bootstrap':
            try:
                q = parse_qs(u.query)
                proj = (q.get('proj') or ['bg'])[0]
                if proj not in SUBS_OF:
                    proj = 'bg'
                self._send(200, bootstrap(proj))
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/':
            self._send_file('index.html', 'text/html; charset=utf-8')
        elif u.path == '/echarts.min.js':
            self._send_file('echarts.min.js', 'application/javascript; charset=utf-8')
        elif u.path == '/daily.js':
            self._send_file('daily.js', 'application/javascript; charset=utf-8')
        elif u.path == '/pm.js':
            self._send_file('pm.js', 'application/javascript; charset=utf-8')
        elif u.path == '/construction_flow.js':
            self._send_file('construction_flow.js', 'application/javascript; charset=utf-8')
        elif u.path == '/daily_plan_import.js':
            self._send_file('daily_plan_import.js', 'application/javascript; charset=utf-8')
        elif u.path == '/api/daily_plan_template':
            try:
                path=os.path.join(BASE,'outputs','daily-plan-import','每日计划数量导入模板.xlsx')
                with open(path,'rb') as f:body=f.read()
                self.send_response(200)
                self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+quote('每日计划数量导入模板.xlsx'))
                self.send_header('Content-Length',str(len(body)));self.end_headers();self.wfile.write(body)
            except Exception as e:self._send(500,{'ok':False,'err':str(e)})
        elif u.path in ('/api/construction_flow', '/api/construction_report'):
            try:
                q = parse_qs(u.query); proj = (q.get('proj') or ['bg'])[0]
                if proj not in SUBS_OF: raise ValueError('项目不存在')
                model = construction_flow.calculate(pm_snapshot(proj), construction_flow.load(project_root(proj)))
                report = construction_flow.report(model, PROJ_NAMES[proj])
                if u.path == '/api/construction_flow':
                    self._send(200, {'ok':True, 'model':model, 'report':report, 'exports':['html','docx']})
                else:
                    fmt = (q.get('fmt') or ['html'])[0]
                    if fmt == 'docx':
                        import construction_word
                        body = construction_word.make_report(model, PROJ_NAMES[proj])
                        ctype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    elif fmt == 'html':
                        body = report.encode('utf-8')
                        ctype = 'text/html; charset=utf-8'
                    else: raise ValueError('不支持的施工汇报格式')
                    self.send_response(200)
                    self.send_header('Content-Type', ctype)
                    self.send_header('Content-Disposition', "attachment; filename*=UTF-8''"+quote(PROJ_NAMES[proj]+'_施工进度汇报.'+fmt))
                    self.send_header('Content-Length', str(len(body))); self.end_headers(); self.wfile.write(body)
            except Exception as e: self._send(400, {'ok':False, 'err':str(e)})
        elif u.path == '/subsystem_export.js':
            self._send_file('subsystem_export.js', 'application/javascript; charset=utf-8')
        elif u.path == '/subsystem_create.js':
            self._send_file('subsystem_create.js', 'application/javascript; charset=utf-8')
        elif u.path == '/api/daily':
            try:
                q = parse_qs(u.query); proj = (q.get('proj') or ['bg'])[0]
                self._send(200, load_daily(proj))
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/version':
            self._send(200, {'ok': True, 'version': '2026.10.18-performance', 'app_dir': BASE})
        elif u.path == '/api/pm':
            try:
                q = parse_qs(u.query); proj = (q.get('proj') or ['bg'])[0]
                self._send(200, pm_snapshot(proj))
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/projects':
            self._send(200, {'ok':True,'projects':project_list()})
        elif u.path == '/api/people_import_template':
            try:
                q=parse_qs(u.query);kind=(q.get('kind') or [''])[0]
                body,name=people_import_template(kind)
                self.send_response(200)
                self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+quote(name))
                self.send_header('Content-Length',str(len(body)));self.end_headers();self.wfile.write(body)
            except (ValueError,TypeError) as e:self._send(400,{'ok':False,'err':str(e)})
            except Exception as e:self._send(500,{'ok':False,'err':'模板生成失败：'+str(e)})
        elif u.path == '/api/subsystem_actions_export':
            try:
                q=parse_qs(u.query);proj=(q.get('proj') or ['bg'])[0];sub=(q.get('sub') or [''])[0]
                body,name=export_subsystem_actions(proj,sub)
                self.send_response(200)
                self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+quote(name))
                self.send_header('Content-Length',str(len(body)));self.end_headers();self.wfile.write(body)
            except (ValueError,TypeError) as e:self._send(400,{'ok':False,'err':str(e)})
            except Exception as e:self._send(500,{'ok':False,'err':'汇总导出失败：'+str(e)})
        elif u.path == '/api/wbs_template':
            try:
                body=pm_core.wbs_template();name='标准WBS导入模板.xlsx'
                self.send_response(200)
                self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+quote(name))
                self.send_header('Content-Length',str(len(body)));self.end_headers();self.wfile.write(body)
            except Exception as e:self._send(500,{'ok':False,'err':str(e)})
        elif u.path == '/api/work_package_export':
            try:
                q=parse_qs(u.query);proj=(q.get('proj') or ['bg'])[0];task_id=(q.get('task_id') or [''])[0]
                body,name=export_work_package(proj,task_id)
                self.send_response(200)
                self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+quote(name))
                self.send_header('Content-Length',str(len(body)));self.end_headers();self.wfile.write(body)
            except ValueError as e:self._send(400,{'ok':False,'err':str(e)})
            except Exception as e:self._send(500,{'ok':False,'err':str(e)})
        elif u.path == '/api/work_package_summary':
            try:
                q=parse_qs(u.query);proj=(q.get('proj') or ['bg'])[0]
                body,name=export_work_package_summary(proj)
                self.send_response(200)
                self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+quote(name))
                self.send_header('Content-Length',str(len(body)));self.end_headers();self.wfile.write(body)
            except ValueError as e:self._send(400,{'ok':False,'err':str(e)})
            except Exception as e:self._send(500,{'ok':False,'err':str(e)})
        elif u.path == '/README_使用说明.txt':
            self._send_file('README_使用说明.txt', 'text/plain; charset=utf-8')
        elif u.path == '/api/audit_img':
            # 审计图片访问：?proj=bg|titan&date=YYYY-MM-DD&file=xxx
            try:
                q = parse_qs(u.query)
                proj = (q.get('proj') or ['bg'])[0]
                date = (q.get('date') or [''])[0]
                fname = (q.get('file') or [''])[0]
                if proj not in SUBS_OF or not date or not fname:
                    self._send(400, {'ok': False, 'err': '缺少参数'})
                    return
                d = audit_dir(proj, date)
                fpath = os.path.realpath(os.path.join(d, fname))
                if not fpath.startswith(os.path.realpath(audit_dir(proj, date)) + os.sep) or not os.path.exists(fpath):
                    self._send(404, {'ok': False, 'err': '图片不存在'})
                    return
                ctype = 'image/jpeg' if fname.lower().endswith(('.jpg', '.jpeg')) else 'image/png'
                with open(fpath, 'rb') as f:
                    body = f.read()
                self.send_response(200)
                self.send_header('Content-Type', ctype)
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/org':
            try:
                q = parse_qs(u.query)
                proj = (q.get('proj') or ['bg'])[0]
                if proj not in SUBS_OF:
                    proj = 'bg'
                org = org_deliver.load_org(proj)
                self._send(200, {'ok': True, 'org': org, 'org_stats': org_deliver.org_stats(org)})
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/org_photo':
            try:
                q = parse_qs(u.query); proj = (q.get('proj') or [''])[0]; fname = (q.get('file') or [''])[0]
                if proj not in SUBS_OF or not fname or os.path.basename(fname) != fname: raise ValueError('人员照片参数无效')
                photo_dir = os.path.realpath(org_deliver._paths(proj)['photos'])
                fpath = os.path.realpath(os.path.join(photo_dir, fname))
                if not fpath.startswith(photo_dir + os.sep) or not os.path.isfile(fpath):
                    self._send(404, {'ok':False,'err':'人员照片不存在'}); return
                ext = os.path.splitext(fname)[1].lower(); ctype = {'.jpg':'image/jpeg','.jpeg':'image/jpeg','.png':'image/png','.webp':'image/webp'}.get(ext)
                if not ctype: raise ValueError('人员照片格式无效')
                with open(fpath,'rb') as photo_file: body=photo_file.read()
                self.send_response(200);self.send_header('Content-Type',ctype);self.send_header('Content-Length',str(len(body)));self.send_header('Cache-Control','no-cache');self.end_headers();self.wfile.write(body)
            except ValueError as e: self._send(400, {'ok':False,'err':str(e)})
            except Exception as e: self._send(500, {'ok':False,'err':str(e)})
        elif u.path == '/api/deliver':
            try:
                q = parse_qs(u.query)
                proj = (q.get('proj') or ['bg'])[0]
                if proj not in SUBS_OF:
                    proj = 'bg'
                self._send(200, {'ok': True, 'manifest': org_deliver.load_manifest(proj)})
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/sub_progress_export':
            try:
                q=parse_qs(u.query);body,name=export_sub_progress((q.get('proj') or ['bg'])[0],(q.get('sub') or [''])[0])
                self.send_response(200);self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+quote(name));self.send_header('Content-Length',str(len(body)));self.end_headers();self.wfile.write(body)
            except Exception as e:self._send(500,{'ok':False,'err':str(e)})
        elif u.path == '/api/export':
            try:
                q = parse_qs(u.query)
                proj = (q.get('proj') or ['bg'])[0]
                if proj not in SUBS_OF:
                    proj = 'bg'
                table = (q.get('table') or [''])[0]
                fmt = (q.get('format') or ['xlsx'])[0]
                if fmt not in ('xlsx', 'csv', 'pdf', 'docx'):
                    self._send(400, {'ok': False, 'err': 'format 仅支持 xlsx/csv/pdf/docx'})
                    return
                inc_del = (q.get('include_deleted') or [''])[0] == '1'
                fpath = org_deliver.export_to(proj, table, fmt, include_deleted=inc_del,
                                              editor=(q.get('editor') or [''])[0],
                                              version=(q.get('version') or ['V1.0'])[0])
                self._send(200, {'ok': True, 'path': fpath, 'name': os.path.basename(fpath)})
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/download':
            # 导出/交付文件下载: ?proj=bg&rel=导出文件/xxx.xlsx
            try:
                q = parse_qs(u.query)
                proj = (q.get('proj') or ['bg'])[0]
                rel = (q.get('rel') or [''])[0]
                if proj not in SUBS_OF or not rel:
                    self._send(400, {'ok': False, 'err': '缺少参数'})
                    return
                b1 = org_deliver._paths(proj)['export']
                b2 = org_deliver._paths(proj)['deliver']
                c1 = os.path.realpath(os.path.join(b1, rel))
                c2 = os.path.realpath(os.path.join(b2, rel))
                fpath = None
                if c1.startswith(os.path.realpath(b1) + os.sep) and os.path.exists(c1):
                    fpath = c1
                elif c2.startswith(os.path.realpath(b2) + os.sep) and os.path.exists(c2):
                    fpath = c2
                if not fpath:
                    self._send(404, {'ok': False, 'err': '文件不存在'})
                    return
                ctype = 'application/octet-stream'
                if fpath.lower().endswith('.pdf'):
                    ctype = 'application/pdf'
                elif fpath.lower().endswith('.docx'):
                    ctype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                elif fpath.lower().endswith('.xlsx'):
                    ctype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                with open(fpath, 'rb') as f:
                    body = f.read()
                self.send_response(200)
                self.send_header('Content-Type', ctype)
                self.send_header('Content-Length', str(len(body)))
                fname = os.path.basename(fpath)
                fname_ascii = fname.encode('ascii', 'replace').decode('ascii')
                self.send_header('Content-Disposition',
                                 "attachment; filename=\"%s\"; filename*=UTF-8''%s" %
                                 (fname_ascii, quote(fname)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        else:
            self._send(404, {'ok': False, 'err': 'not found'})

    def _send_file(self, name, ctype):
        p = os.path.join(BASE, name)
        if not os.path.exists(p):
            self._send(404, {'ok': False, 'err': name + ' missing'})
            return
        http_delivery.send_file(self,p,ctype)

    def do_POST(self):
        if not AUTH.gate(self):
            return
        if urlparse(self.path).path in ('/api/pointmap_import','/api/materials/upload'):
            # Upload modules use their own SQLite/job locks and immutable file paths.
            with DATA_LOCK:refresh_subsystems()
            try:self._do_POST()
            finally:BOOTSTRAP_CACHE.clear();PM_CACHE.clear()
            return
        with DATA_LOCK:
            refresh_subsystems()
            try:self._do_POST()
            finally:BOOTSTRAP_CACHE.clear();PM_CACHE.clear()

    def _do_POST(self):
        u = urlparse(self.path)
        if u.path == '/api/pointmap_import':
            try:
                materials.check_local_request(self)
                q = parse_qs(u.query)
                proj = (q.get('proj') or [''])[0]
                if proj not in SUBS_OF: raise ValueError('未知项目')
                options = {key:(q.get(key) or [''])[0] for key in ('width','x','y','page')}
                result = point_import.upload(project_root(proj), (q.get('name') or [''])[0], self.rfile, int(self.headers.get('Content-Length',0)), options)
                self._send(200, result)
            except (ValueError, OSError) as e: self._send(400, {'ok':False,'err':str(e)})
            return
        if u.path == '/api/materials/upload':
            try:
                materials.check_local_request(self)
                q = parse_qs(u.query)
                proj = (q.get('proj') or [''])[0]
                if proj not in SUBS_OF: raise ValueError('未知项目')
                result = materials.upload(project_root(proj), (q.get('name') or [''])[0], (q.get('category') or [''])[0], self.rfile, int(self.headers.get('Content-Length', 0)), (q.get('note') or [''])[0])
                self._send(200, result)
            except (ValueError, OSError) as e: self._send(400, {'ok':False, 'err':str(e)})
            return
        try:
            ln = int(self.headers.get('Content-Length', 0))
            if u.path == '/api/pointmap' and not 0 < ln <= 256 * 1024:
                self._send(413, {'ok':False,'err':'点位图请求超过256KB限制'});return
            if u.path == '/api/daily_plan_import' and not 0 < ln <= 12 * 1024 * 1024:
                self._send(413, {'ok':False,'err':'导入请求过大，文件不得超过8MB'});return
            if u.path == '/api/subsystem_export' and not 0 < ln <= 40 * 1024 * 1024:
                self._send(413, {'ok':False,'err':'导出数据超过40MB限制'});return
            if u.path == '/api/org_save' and not 0 < ln <= 35 * 1024 * 1024:
                self._send(413, {'ok':False,'err':'人员档案数据超过35MB限制'});return
            data = json.loads(self.rfile.read(ln).decode('utf-8'))
        except Exception:
            self._send(400,{'ok':False,'err':'请求必须是有效 JSON'});return
        if not isinstance(data,dict):self._send(400,{'ok':False,'err':'请求数据格式无效'});return
        if u.path == '/api/pointmap':
            try:
                materials.check_local_request(self)
                proj = data.get('proj')
                if proj not in SUBS_OF: raise ValueError('未知项目')
                result=point_map.mutate(project_root(proj), data)
                try:result['global_sync_result']=point_statistics.sync(sys.modules[__name__],proj)
                except Exception as e:result['sync_error']='点位已保存，全局同步失败：'+str(e)
                self._send(200,result)
            except (ValueError, OSError) as e: self._send(400, {'ok':False,'err':str(e)})
        elif u.path == '/api/materials/action':
            try:
                materials.check_local_request(self)
                proj = data.get('proj')
                if proj not in SUBS_OF: raise ValueError('未知项目')
                self._send(200, materials.action(project_root(proj), data))
            except (ValueError, OSError) as e: self._send(400, {'ok':False, 'err':str(e)})
        elif u.path == '/api/management':
            try:
                proj = data.get('proj')
                if proj not in SUBS_OF: raise ValueError('未知项目')
                self._send(200, management.mutate(project_root(proj), data))
            except Exception as e: self._send(400, {'ok':False, 'err':str(e)})
        elif u.path == '/api/projects':
            try:self._send(200,create_project(data))
            except (ValueError,TypeError) as e:self._send(400,{'ok':False,'err':str(e)})
            except Exception as e:self._send(500,{'ok':False,'err':'创建项目失败：'+str(e)})
        elif u.path == '/api/people_import_preview':
            try:self._send(200,parse_people_import(data))
            except (ValueError,TypeError) as e:self._send(400,{'ok':False,'err':str(e)})
            except Exception as e:self._send(500,{'ok':False,'err':'批量导入失败：'+str(e)})
        elif u.path == '/api/subsystem_action_item':
            try:self._send(200,subsystem_action_item(data))
            except (ValueError,TypeError) as e:self._send(400,{'ok':False,'err':str(e)})
            except Exception as e:self._send(500,{'ok':False,'err':'操作失败：'+str(e)})
        elif u.path == '/api/subsystem_create':
            try:
                from subsystem_create import create_subsystem
                self._send(200, create_subsystem(sys.modules[__name__], data))
            except (ValueError, TypeError) as e:
                self._send(400, {'ok':False,'err':str(e)})
            except Exception as e:
                self._send(500, {'ok':False,'err':'添加失败，已回滚：'+str(e)})
        elif u.path == '/api/subsystem_delete':
            try:
                from subsystem_create import delete_subsystem
                self._send(200,delete_subsystem(sys.modules[__name__],data))
            except (ValueError,TypeError) as e:self._send(400,{'ok':False,'err':str(e)})
            except Exception as e:self._send(500,{'ok':False,'err':'删除失败，已回滚：'+str(e)})
        elif u.path == '/api/subsystem_export':
            try:
                if data.get('proj') not in SUBS_OF or data.get('sub') not in SUBS_OF[data['proj']]:
                    raise ValueError('子系统无效')
                from subsystem_export import build_bundle
                body, name = build_bundle(data)
                self.send_response(200)
                self.send_header('Content-Type', 'application/zip')
                self.send_header('Content-Length', str(len(body)))
                self.send_header('Content-Disposition', "attachment; filename*=UTF-8''" + quote(name))
                self.end_headers()
                self.wfile.write(body)
            except (ValueError, TypeError, KeyError) as e:
                self._send(400, {'ok':False,'err':str(e)})
            except Exception as e:
                self._send(500, {'ok':False,'err':'生成导出包失败：'+str(e)})
        elif u.path == '/api/save':
            try:
                result=save_payload(data)
                point_statistics.sync(sys.modules[__name__],data.get('proj','bg'))
                self._send(200,result)
            except Exception as e:
                import traceback
                traceback.print_exc()
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/audit':
            try:
                self._send(200, save_audit(data))
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/daily':
            try:
                self._send(200, save_daily(data))
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/sub_item':
            try:self._send(200,save_sub_item(data))
            except Exception as e:self._send(500,{'ok':False,'err':str(e)})
        elif u.path == '/api/sub_settings':
            try:self._send(200,save_sub_settings(data))
            except Exception as e:self._send(500,{'ok':False,'err':str(e)})
        elif u.path == '/api/daily_plan':
            try:
                self._send(200,save_daily_plan(data))
            except Exception as e:
                self._send(500,{'ok':False,'err':str(e)})
        elif u.path == '/api/daily_plan_import':
            try:
                import daily_plan_import
                self._send(200,daily_plan_import.handle(sys.modules[__name__],data))
            except Exception as e:self._send(400,{'ok':False,'err':str(e)})
        elif u.path == '/api/step_save':
            try:
                self._send(200, save_step_records(data))
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/sub_actual':
            try:
                self._send(200, sub_actual(data.get('proj','bg'), data.get('date',''), data.get('updates')))
            except Exception as e:
                self._send(500, {'ok':False,'err':str(e)})
        elif u.path == '/api/sub_plan':
            try:
                self._send(200, save_sub_plan(data))
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/pm':
            try:
                proj = data.get('proj', 'bg')
                if proj not in SUBS_OF: proj = 'bg'
                action = data.get('action')
                if action == 'save_profile':
                    result = pm_core.save_profile(DIR, proj, PROJ_NAMES[proj], SUBS_OF[proj], data.get('profile') or {}, data.get('operator',''))
                    if result.get('ok'):
                        if proj=='titan':update_progress_check(proj)
                        else:update_progress_check_bg(proj)
                elif action == 'generate_wbs':
                    result = pm_core.generate_wbs(DIR, proj, PROJ_NAMES[proj], SUBS_OF[proj], data.get('quantity_map') or {}, data.get('operator',''))
                elif action == 'update_task':
                    result = pm_core.update_task(DIR, proj, PROJ_NAMES[proj], SUBS_OF[proj], data.get('task_id',''), data.get('task') or {}, data.get('operator',''))
                elif action == 'delete_task':
                    result = pm_core.delete_task(DIR, proj, PROJ_NAMES[proj], SUBS_OF[proj], data.get('task_id',''), data.get('operator',''))
                elif action == 'record_actual':
                    result = pm_core.record_actual(DIR, proj, PROJ_NAMES[proj], SUBS_OF[proj], data.get('task_id',''), data.get('work_date') or datetime.date.today().isoformat(), data.get('qty',0), data.get('note',''), data.get('operator',''))
                elif action == 'import_wbs':
                    result = pm_core.import_wbs(DIR, proj, PROJ_NAMES[proj], SUBS_OF[proj], data.get('filename',''), data.get('content',''), bool(data.get('dry_run',True)), data.get('operator',''))
                else:
                    result = {'ok': False, 'err': '未知项目管理操作'}
                self._send(200, result)
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/construction_flow':
            try:
                proj = data.get('proj', 'bg')
                if proj not in SUBS_OF: raise ValueError('项目不存在')
                config = data.get('config') or {}
                model = construction_flow.calculate(pm_snapshot(proj), config)
                construction_flow.save(project_root(proj), model['config'])
                self._send(200, {'ok':True, 'msg':'施工关系已保存并重新排程', 'model':model, 'report':construction_flow.report(model,PROJ_NAMES[proj])})
            except Exception as e: self._send(400, {'ok':False, 'err':str(e)})
        elif u.path == '/api/upload':
            try:
                self._send(200, save_audit_img(data))
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/org_save':
            try:
                proj = data.get('proj', 'bg')
                if proj not in SUBS_OF:
                    proj = 'bg'
                self._send(200, org_deliver.save_org(proj, data))
            except Exception as e:
                import traceback
                traceback.print_exc()
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/deliver_gen':
            try:
                proj = data.get('proj', 'bg')
                if proj not in SUBS_OF:
                    proj = 'bg'
                self._send(200, org_deliver.gen_deliver(
                    proj, data.get('type', 'quantity'), data.get('format', 'pdf'),
                    editor=data.get('editor', ''), note=data.get('note', '')))
            except Exception as e:
                import traceback
                traceback.print_exc()
                self._send(500, {'ok': False, 'err': str(e)})
        elif u.path == '/api/deliver_update':
            try:
                proj = data.get('proj', 'bg')
                if proj not in SUBS_OF:
                    proj = 'bg'
                self._send(200, org_deliver.update_manifest(proj, data.get('no', ''), data.get('fields', {})))
            except Exception as e:
                self._send(500, {'ok': False, 'err': str(e)})
        else:
            self._send(404, {'ok': False, 'err': 'not found'})


def save_sub_item(data):
    import shutil, math, copy
    proj,sub,kind,action=data.get('proj'),data.get('sub'),data.get('kind'),data.get('action')
    if proj not in SUBS_OF or sub not in SUBS_OF[proj] or kind not in ('detail','step') or action not in ('add','edit','delete'):raise ValueError('操作参数无效')
    F=_paths(proj);books={k:load_workbook(F[k]) for k in ('p30','p31','p32')}
    ws=books['p32']['详细工程量' if kind=='detail' else '施工步骤'];namecol=2 if kind=='detail' else 3
    old=str(data.get('old_name') or '');name=str(data.get('name') or '').strip()
    matches=[r for r in range(2,ws.max_row+1) if ws.cell(r,1).value==sub]
    target=next((r for r in matches if str(ws.cell(r,namecol).value)==old),None)
    if action!='add' and target is None:raise ValueError('子项不存在，请刷新页面')
    if action!='delete':
        if not name:raise ValueError('子项名称不能为空')
        if any(r!=target and str(ws.cell(r,namecol).value)==name for r in matches):raise ValueError('同一子系统内子项名称不能重复')
        start=str(data.get('start') or '');finish=str(data.get('finish') or '')
        for date in (start,finish):
            if date:datetime.date.fromisoformat(date)
        if start and finish and finish<start:raise ValueError('计划完成不能早于计划开始')
        qty=None if data.get('total') in (None,'') else float(data['total'])
        if qty is not None and (not math.isfinite(qty) or qty<0):raise ValueError('计划量必须为非负数字')
    backup=os.path.join(os.path.dirname(F['p32']),'子项修改备份',datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    os.makedirs(backup,exist_ok=True)
    for k in books:shutil.copy2(F[k],os.path.join(backup,os.path.basename(F[k])))
    if action=='delete':
        archive=books['p32']['已删除子项'] if '已删除子项' in books['p32'].sheetnames else books['p32'].create_sheet('已删除子项')
        archive.append([datetime.datetime.now().isoformat(),kind]+[ws.cell(target,c).value for c in range(1,ws.max_column+1)])
        ws.delete_rows(target)
    else:
        r=target if action=='edit' else ws.max_row+1
        if action=='add':
            if kind=='detail':values=[sub,name,data.get('unit',''),qty,0,0,f'=IF(ISNUMBER(D{r}),D{r}-F{r},"")',f'=IFERROR(F{r}/D{r},"")','手工新增',data.get('note',''),start,finish]
            else:values=[sub,len(matches)+1,name,start,finish,'未开始','','',data.get('note','')]
            for c,v in enumerate(values,1):ws.cell(r,c).value=v
            for cell in ws[r]:cell.font=BODY_FONT;cell.border=BORDER;cell.alignment=CENTER
        else:
            ws.cell(r,namecol).value=name
            if kind=='detail':
                ws.cell(r,3).value=data.get('unit','');ws.cell(r,4).value=qty;ws.cell(r,10).value=data.get('note','');ws.cell(r,11).value=start;ws.cell(r,12).value=finish
            else:ws.cell(r,4).value=start;ws.cell(r,5).value=finish;ws.cell(r,9).value=data.get('note','')
        if kind=='detail':
            ws.cell(1,11).value='计划开始';ws.cell(1,12).value='计划完成'
            if 'workers' in data:
                people=org_deliver.load_org(proj)['people'];known={str(p['pid']):p for p in people}
                ids=list(dict.fromkeys(str(v) for v in data['workers']))
                if any(v not in known for v in ids):raise ValueError('施工人员不存在，请重新选择')
                ws.cell(1,13).value='施工人员ID';ws.cell(1,14).value='施工人数'
                ws.cell(r,13).value=json.dumps(ids,ensure_ascii=False);ws.cell(r,14).value=len(ids)
    if kind=='detail' and action!='add':
        flow=books['p30']['每日计量']
        for r in range(flow.max_row,1,-1):
            if flow.cell(r,2).value==sub and str(flow.cell(r,3).value)==old:
                if action=='edit':flow.cell(r,3).value=name;flow.cell(r,5).value=data.get('unit','')
                else:
                    archive=books['p30']['已删除计量'] if '已删除计量' in books['p30'].sheetnames else books['p30'].create_sheet('已删除计量')
                    archive.append([flow.cell(r,c).value for c in range(1,flow.max_column+1)]);flow.delete_rows(r)
        overview=books['p30']['工程量总览']
        for r in range(2,overview.max_row+1):
            if canonical_sub(proj, overview.cell(r,2).value)==sub and str(overview.cell(r,3).value)==old:
                if action=='edit':overview.cell(r,3).value=name;overview.cell(r,5 if proj=='titan' else 4).value=qty;overview.cell(r,4 if proj=='titan' else 5).value=data.get('unit','')
                else:overview.cell(r,6).value=0;overview.cell(r,7).value=0
        state=books['p31']['子系统状态']
        dd=books['p32']['详细工程量'];rr=[r for r in range(2,dd.max_row+1) if dd.cell(r,1).value==sub]
        for r in range(2,state.max_row+1):
            if state.cell(r,1).value==sub:
                state.cell(r,5).value=sum(float(dd.cell(dr,5).value or 0) for dr in rr);state.cell(r,6).value=sum(float(dd.cell(dr,6).value or 0) for dr in rr)
                if state.cell(r,15).value==old:state.cell(r,15).value=name if action=='edit' else ''
    try:
        for k,book in books.items():book.save(F[k]);book.close()
        if proj=='titan':update_progress_check(proj)
        else:update_progress_check_bg(proj)
    except Exception:
        for k in books:shutil.copy2(os.path.join(backup,os.path.basename(F[k])),F[k])
        raise
    return {'ok':True,'msg':'子项已保存，统计与图表已同步；删除内容保留在归档及备份中'}


def sub_settings(proj):
    path=os.path.join(os.path.dirname(_paths(proj)['p30']),'子系统计划设置.json')
    if not os.path.exists(path):return {}
    with open(path,encoding='utf-8') as f:return json.load(f)


def sub_schedule(proj,sub):
    cfg=sub_settings(proj).get(sub,{})
    if 'daily' in cfg:return cfg['daily']
    w=load_workbook(_paths(proj)['p30']);ret={}
    if proj=='titan' and sub in BASE_SUBS['titan']:
        ws=w['每日计划'];c=[3,5,7,9,11,13][SUBS_OF[proj].index(sub)]
        for r in range(2,ws.max_row+1):
            if str(ws.cell(r,1).value or '').startswith('第'):ret[str(ws.cell(r,2).value)[:10]]=float(ws.cell(r,c).value or 0)
    elif '子系统每日计划' in w.sheetnames:
        for r in list(w['子系统每日计划'].values)[1:]:
            if r[1]==sub:ret[str(r[0])[:10]]=float(r[2] or 0)
    w.close();return ret


def save_sub_settings(data):
    import math
    proj,sub=data.get('proj'),data.get('sub')
    if proj not in SUBS_OF or sub not in SUBS_OF[proj]:raise ValueError('子系统无效')
    cfg=sub_settings(proj);old=cfg.get(sub,{})
    name=str(data.get('name') or '').strip()
    if not name:raise ValueError('名称不能为空')
    if any(k!=sub and v.get('name',k)==name for k,v in cfg.items()) or any(k!=sub and k==name for k in SUBS_OF[proj]):raise ValueError('名称不能重复')
    actual_start=str(data.get('actual_start',old.get('actual_start','')) or '')
    actual_finish=str(data.get('actual_finish',old.get('actual_finish','')) or '')
    for value in (actual_start,actual_finish):
        if value and datetime.date.fromisoformat(value).isoformat()!=value:raise ValueError('实际日期无效')
    if actual_finish and (not actual_start or actual_finish<actual_start):raise ValueError('实际结束需填写实际开始，且不能早于实际开始')
    start=datetime.date.fromisoformat(data['start']);end=datetime.date.fromisoformat(data['finish'])
    if end<start or (end-start).days>730:raise ValueError('工期需为 1 至 731 个自然日')
    count=(end-start).days+1;dates=[(start+datetime.timedelta(days=i)).isoformat() for i in range(count)]
    baseline=sub_schedule(proj,sub);active=[v for d,v in sorted(baseline.items()) if v>0]
    total=sum(active)
    if proj=='bg' or not baseline:
        w=load_workbook(_paths(proj)['p30']);ws=w['工程量总览']
        total=sum(float(ws.cell(r,5 if proj=='titan' else 4).value or 0) for r in range(2,ws.max_row+1) if canonical_sub(proj,ws.cell(r,2).value)==sub)
        w.close()
    daily={d:0.0 for d in dates}
    if active:
        for i,v in enumerate(active):daily[dates[min(count-1,int(i*count/len(active)))]]+=v
    if not active:daily=dict(zip(dates,allocate_plan(total,[1]*count)))
    elif proj=='bg':daily=dict(zip(dates,allocate_plan(total,list(daily.values()))))
    if proj=='titan':daily=dict(zip(dates,allocate_plan(round(sum(daily.values())),list(daily.values()))))
    seq=int(data.get('start_no') or 1)
    if seq<1:raise ValueError('起始条号必须大于 0')
    cfg[sub]={**old,'name':name,'start':start.isoformat(),'finish':end.isoformat(),'start_no':seq,'daily':daily,'actual_start':actual_start,'actual_finish':actual_finish,'offset_start':int(str(SUBS_MAP[sub]['plan_start']).split('+')[-1]) if proj=='bg' and sub in SUBS_MAP else 1}
    if old.get('start')==start.isoformat() and old.get('finish')==end.isoformat():cfg[sub]['daily']=baseline
    path=os.path.join(os.path.dirname(_paths(proj)['p30']),'子系统计划设置.json')
    import shutil
    if os.path.exists(path):shutil.copy2(path,path+'.bak')
    with open(path+'.tmp','w',encoding='utf-8') as f:json.dump(cfg,f,ensure_ascii=False,indent=2)
    os.replace(path+'.tmp',path)
    w=load_workbook(_paths(proj)['p31']);ws=w['子系统状态']
    for r in range(2,ws.max_row+1):
        if ws.cell(r,1).value==sub:ws.cell(r,3).value=start.isoformat();ws.cell(r,4).value=end.isoformat()
    ws.cell(1,15).value='实际开始';ws.cell(1,16).value='实际结束'
    for r in range(2,ws.max_row+1):
        if ws.cell(r,1).value==sub:
            ws.cell(r,15).value=actual_start;ws.cell(r,16).value=actual_finish
    w.save(_paths(proj)['p31']);w.close()
    if proj=='titan':update_progress_check(proj)
    else:update_progress_check_bg(proj)
    return {'ok':True,'msg':'名称及工期已保存，计划总量保留，日期已同步'}


def export_sub_progress(proj,sub):
    if proj not in SUBS_OF or sub not in SUBS_OF[proj]:raise ValueError('子系统无效')
    cfg=sub_settings(proj).get(sub,{});plan=sub_schedule(proj,sub)
    w=load_workbook(_paths(proj)['p30']);actual={};notes={}
    for r in list(w['每日计量'].values)[1:]:
        if r[1]==sub:
            date=r[0].strftime('%Y-%m-%d') if isinstance(r[0],(datetime.date,datetime.datetime)) else str(r[0])[:10]
            actual[date]=actual.get(date,0)+float(r[3] or 0)
            if len(r)>7 and r[7]:notes[date]=str(r[7])
    w.close()
    start=str(cfg.get('start') or '');finish=str(cfg.get('finish') or '')
    if start and finish and start<=finish:
        dates=[]
        cursor=datetime.date.fromisoformat(start);last=datetime.date.fromisoformat(finish)
        while cursor<=last:
            dates.append(cursor.isoformat());cursor+=datetime.timedelta(days=1)
    else:
        dates=sorted(set(plan)|set(actual))
    dates=sorted(set(dates)|set(actual)|{v for v in (cfg.get('actual_start'),cfg.get('actual_finish')) if v})
    total=sum(plan.get(d,0) for d in dates);seq=cfg.get('start_no',1)
    out=openpyxl.Workbook();ws=out.active;ws.title='施工进度表'
    headers=['工作日','日期','起始条号','结束条号','计划条数','累计计划','计划进度','实际完成（填）','累计实际','实际进度','累计偏差（条）','当日达成率','备注']
    ws.append([cfg.get('name',sub)+'施工进度表']);ws.merge_cells('A1:M1');ws.append(headers)
    cp=0
    for i,d in enumerate(dates,1):
        q=plan.get(d,0);cp+=q;r=i+2
        valid=q>0 and float(q).is_integer() and float(cp).is_integer()
        a=seq+int(cp-q) if valid else '';b=seq+int(cp)-1 if valid else ''
        ws.append([i,d,a,b,q,f'=SUM(E$3:E{r})',f'=IF(SUM(E$3:E${len(dates)+2})=0,"",F{r}/SUM(E$3:E${len(dates)+2}))',actual.get(d,0),f'=SUM(H$3:H{r})',f'=IF(SUM(E$3:E${len(dates)+2})=0,"",I{r}/SUM(E$3:E${len(dates)+2}))',f'=I{r}-F{r}',f'=IF(E{r}=0,"",H{r}/E{r})',notes.get(d,'') if valid or q==0 else '计划量非整数，未生成条号'])
        for c in (7,10,12):ws.cell(r,c).number_format='0.00%'
    for row in ws.iter_rows():
        for cell in row:cell.border=BORDER;cell.alignment=CENTER;cell.font=BODY_FONT
    for cell in ws[2]:cell.fill=HDR_FILL;cell.font=HDR_FONT
    ws.row_dimensions[1].height=30;ws.freeze_panes='C3';ws.auto_filter.ref=f'A2:M{ws.max_row}'
    for c in range(1,14):ws.column_dimensions[get_column_letter(c)].width=16 if c!=13 else 38
    ws.sheet_properties.pageSetUpPr.fitToPage=True;ws.page_setup.orientation='landscape';ws.page_setup.paperSize=ws.PAPERSIZE_A3;ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=0;ws.print_title_rows='1:2'
    info=out.create_sheet('工期信息');info.append(['子系统','计划开始','计划结束','实际开始','实际结束']);info.append([cfg.get('name',sub),start,finish,cfg.get('actual_start',''),cfg.get('actual_finish','')])
    buf=io.BytesIO();out.save(buf);return buf.getvalue(),cfg.get('name',sub)+'施工进度表.xlsx'


def export_work_package(proj, task_id):
    """导出可独立下发、执行和验收的工作包文件。"""
    if proj not in SUBS_OF: raise ValueError('项目无效')
    snap=pm_snapshot(proj)
    task=next((x for x in snap.get('tasks',[]) if x.get('id')==task_id),None)
    if not task: raise ValueError('未找到工作包')
    if task.get('is_summary') and task.get('source_type')!='subsystem': raise ValueError('汇总节点不能作为工作包下载')
    source=bootstrap(proj);out=openpyxl.Workbook();ws=out.active;ws.title='工作包说明'
    display_plan=task.get('display_plan_qty',task.get('plan_qty',0));display_actual=task.get('display_actual_qty',task.get('actual_qty',0))
    rate=round(float(display_actual or 0)/float(display_plan or 0)*100,1) if float(display_plan or 0) else 0
    fields=[
        ('项目名称',snap.get('profile',{}).get('project_name') or PROJ_NAMES[proj]),('WBS编码',task.get('wbs_code','')),
        ('工作包名称',task.get('task_name','')),('所属子系统',task.get('system_name','')),('工作包类型',task.get('task_type') or '工作包'),
        ('责任人/班组',task.get('owner','')),('计划开始',task.get('plan_start','')),('计划完成',task.get('plan_finish','')),
        ('实际开工时间',task.get('actual_start','')),('实际完工时间',task.get('actual_finish','')),
        ('计划工程量',display_plan),('累计实际',display_actual),('计量单位',task.get('display_unit',task.get('unit',''))),
        ('完成率',str(rate)+'%'),('状态',task.get('status','')),('前置工作包',task.get('predecessor','')),
        ('工作说明',task.get('description','')),('交付物',task.get('deliverable','')),('验收标准',task.get('acceptance_criteria','')),
        ('数据来源','子系统自动同步' if task.get('source_type')=='subsystem' else 'WBS工作包')]
    ws.append([task.get('task_name','')+'工作包']);ws.merge_cells('A1:B1');ws.append(['字段','内容'])
    for row in fields:ws.append(row)
    ws['A1'].font=Font(name='微软雅黑',size=14,bold=True,color='FFFFFF');ws['A1'].fill=HDR_FILL;ws['A1'].alignment=CENTER
    for c in ws[2]:c.font=HDR_FONT;c.fill=HDR_FILL;c.border=BORDER;c.alignment=CENTER
    for row in ws.iter_rows(min_row=3):
        for cell in row:cell.font=BODY_FONT;cell.border=BORDER;cell.alignment=WRAP
    ws.column_dimensions['A'].width=20;ws.column_dimensions['B'].width=58;ws.freeze_panes='A3'
    sub=task.get('source_key')
    if task.get('source_type')=='subsystem' and sub:
        details=[x for x in source.get('details',[]) if x.get('sub')==sub]
        sh=out.create_sheet('工程量明细');rebuild_sheet(sh,['工程量名称','单位','计划量','今日完成','累计完成','计划开始','计划完成','来源','备注'],[[x.get(k,'') for k in ('item','unit','total','today','cum','ps','pe','src','note')] for x in details],[28,10,13,13,13,13,13,18,36],{2,3,4,5,6,7})
        steps=[x for x in source.get('steps',[]) if x.get('sub')==sub]
        step_rows=[]
        for step in steps:
            quantities=step.get('quantities') or {};daily_notes=step.get('daily_notes') or {}
            # Export every saved reporting date, including zero quantities and note-only days.
            for date in sorted(set(quantities)|set(daily_notes)) or ['']:
                step_rows.append([step.get(k,'') for k in ('no','step','ps','pe','status','finish','resp','note')]+[date,quantities.get(date,''),daily_notes.get(date,'')])
        sh=out.create_sheet('施工步骤');rebuild_sheet(sh,['序号','施工步骤','计划开始','计划完成','状态','完成日期','责任人','步骤备注','填报日期','当日完成数量','当日备注'],step_rows,[10,34,14,14,12,14,18,36,14,16,36],{1,3,4,5,6,9,10})
        acceptance=[x for x in source.get('acceptance',[]) if x.get('sub')==sub]
        sh=out.create_sheet('验收记录');rebuild_sheet(sh,['验收类型','状态','日期','验收人','结论','不通过原因','整改完成日期'],[[x.get(k,'') for k in ('type','status','date','who','concl','reason','fixdate')] for x in acceptance],[18,12,14,18,36,36,16],{2,3,7})
        plan=sub_schedule(proj,sub);actual={};notes={}
        for r in source.get('metrics',[]):
            if len(r)>3 and r[1]==sub:
                d=str(r[0])[:10];actual[d]=actual.get(d,0)+float(r[3] or 0)
                if len(r)>7 and r[7]:notes[d]=str(r[7])
        cfg=(source.get('sub_settings') or {}).get(sub,{})
        start=str(cfg.get('start') or task.get('plan_start') or '');finish=str(cfg.get('finish') or task.get('plan_finish') or '')
        dates=[]
        if start and finish and start<=finish:
            cursor=datetime.date.fromisoformat(start);last=datetime.date.fromisoformat(finish)
            while cursor<=last:dates.append(cursor.isoformat());cursor+=datetime.timedelta(days=1)
        dates=sorted(set(dates)|set(plan)|set(actual)|{v for v in (cfg.get('actual_start'),cfg.get('actual_finish')) if v})
        cp=ca=0;rows=[]
        for d in dates:
            cp+=float(plan.get(d,0));ca+=float(actual.get(d,0));rows.append([d,plan.get(d,0),actual.get(d,0),cp,ca,ca-cp,notes.get(d,'')])
        sh=out.create_sheet('每日计划与实际');rebuild_sheet(sh,['日期','计划量','实际量','累计计划','累计实际','累计偏差','备注'],rows,[15,13,13,14,14,14,40],{1,2,3,4,5,6})
    else:
        logs=pm_core.task_actual_logs(DIR,proj,PROJ_NAMES[proj],SUBS_OF[proj],task_id)
        sh=out.create_sheet('实际记录');rebuild_sheet(sh,['日期','完成量','工作说明','录入人','录入时间'],[[x.get(k,'') for k in ('work_date','qty','note','operator','created_at')] for x in logs],[15,14,42,18,22],{1,2})
    for sh in out.worksheets:
        sh.sheet_properties.pageSetUpPr.fitToPage=True;sh.page_setup.orientation='landscape';sh.page_setup.fitToWidth=1;sh.page_setup.fitToHeight=0
    safe=''.join('_' if c in '\\/:*?"<>|' else c for c in str(task.get('task_name') or '工作包'))[:80]
    buf=io.BytesIO();out.save(buf);return buf.getvalue(),str(task.get('wbs_code') or '')+'_'+safe+'_工作包.xlsx'


def export_work_package_summary(proj):
    """导出项目全部可执行工作包及其每日实际记录。"""
    if proj not in SUBS_OF: raise ValueError('项目无效')
    snap=pm_snapshot(proj);source=bootstrap(proj)
    tasks=[x for x in snap.get('tasks',[]) if x.get('source_type')=='subsystem' or not x.get('is_summary')]
    out=openpyxl.Workbook();ws=out.active;ws.title='工作包汇总'
    headers=['WBS编码','工作包名称','所属子系统','类型','责任人/班组','计划开始','计划完成','计划量','累计实际','单位','完成率','状态','前置工作包','交付物','验收标准','数据来源']
    rows=[]
    for t in tasks:
        plan=float(t.get('display_plan_qty',t.get('plan_qty',0)) or 0);actual=float(t.get('display_actual_qty',t.get('actual_qty',0)) or 0)
        rows.append([t.get('wbs_code',''),t.get('task_name',''),t.get('system_name',''),t.get('task_type') or '工作包',t.get('owner',''),t.get('plan_start',''),t.get('plan_finish',''),plan,actual,t.get('display_unit',t.get('unit','')),round(actual/plan*100,1) if plan else 0,t.get('status',''),t.get('predecessor',''),t.get('deliverable',''),t.get('acceptance_criteria',''),'子系统自动同步' if t.get('source_type')=='subsystem' else 'WBS工作包'])
    rebuild_sheet(ws,headers,rows,[13,32,24,12,20,14,14,13,13,10,12,12,16,28,32,16],{1,4,6,7,8,9,10,11,12})
    for r in range(2,ws.max_row+1):ws.cell(r,11).number_format='0.0%';ws.cell(r,11).value=float(ws.cell(r,11).value or 0)/100
    ws.freeze_panes='A2';ws.auto_filter.ref=f'A1:P{ws.max_row}'
    detail_rows=[]
    task_by_sub={x.get('source_key'):x for x in tasks if x.get('source_type')=='subsystem'}
    for x in source.get('details',[]):
        t=task_by_sub.get(x.get('sub'))
        if not t:continue
        plan=float(x.get('total') or 0);actual=float(x.get('cum') or 0);rate=actual/plan if plan else 0
        status='已完成' if plan and actual>=plan else ('施工中' if actual else '未开始')
        detail_rows.append([t.get('wbs_code',''),t.get('task_name',''),t.get('system_name',''),'工程量明细',x.get('item',''),x.get('ps',''),x.get('pe',''),plan,actual,x.get('unit',''),rate,status,t.get('owner',''),x.get('src',''),x.get('note','')])
    for x in source.get('steps',[]):
        t=task_by_sub.get(x.get('sub'))
        if not t:continue
        detail_rows.append([t.get('wbs_code',''),t.get('task_name',''),t.get('system_name',''),'施工步骤',str(x.get('no') or '')+' '+str(x.get('step') or ''),x.get('ps',''),x.get('pe',''),'','','',0,x.get('status',''),x.get('resp',''),'施工步骤',x.get('note','')])
    for x in source.get('acceptance',[]):
        t=task_by_sub.get(x.get('sub'))
        if not t:continue
        note='；'.join(v for v in [str(x.get('concl') or ''),str(x.get('reason') or ''),('整改完成：'+str(x.get('fixdate'))) if x.get('fixdate') else ''] if v)
        detail_rows.append([t.get('wbs_code',''),t.get('task_name',''),t.get('system_name',''),'验收节点',x.get('type',''),x.get('date',''),x.get('fixdate',''),'','','',0,x.get('status',''),x.get('who',''),'验收记录',note])
    for t in tasks:
        if t.get('source_type')=='subsystem':continue
        plan=float(t.get('plan_qty') or 0);actual=float(t.get('actual_qty') or 0)
        detail_rows.append([t.get('wbs_code',''),t.get('task_name',''),t.get('system_name',''),'WBS工作包',t.get('description') or t.get('task_name',''),t.get('plan_start',''),t.get('plan_finish',''),plan,actual,t.get('unit',''),actual/plan if plan else 0,t.get('status',''),t.get('owner',''),t.get('deliverable',''),t.get('acceptance_criteria','')])
    detail_rows.sort(key=lambda x:(x[0],{'工程量明细':0,'施工步骤':1,'验收节点':2,'WBS工作包':3}.get(x[3],9),str(x[4])))
    sh=out.create_sheet('工作包明细');rebuild_sheet(sh,['WBS编码','工作包名称','所属子系统','明细类别','明细内容','计划开始','计划完成','计划量','累计实际','单位','完成率','状态','责任人','来源/交付物','备注/验收标准'],detail_rows,[13,30,24,14,36,14,14,13,13,10,12,12,18,24,42],{1,4,6,7,8,9,10,11,12})
    for r in range(2,sh.max_row+1):sh.cell(r,11).number_format='0.0%'
    sh.freeze_panes='A2';sh.auto_filter.ref=f'A1:O{sh.max_row}'
    daily=[]
    for r in source.get('metrics',[]):
        if len(r)>3 and r[1] in task_by_sub:
            t=task_by_sub[r[1]];daily.append([str(r[0])[:10],t.get('wbs_code',''),t.get('task_name',''),t.get('system_name',''),r[2],r[3],r[4] if len(r)>4 else '',r[7] if len(r)>7 else '','子系统同步'])
    for t in tasks:
        if t.get('source_type')=='subsystem':continue
        for x in pm_core.task_actual_logs(DIR,proj,PROJ_NAMES[proj],SUBS_OF[proj],t['id']):
            daily.append([x.get('work_date',''),t.get('wbs_code',''),t.get('task_name',''),t.get('system_name',''),'工作包实际',x.get('qty',0),t.get('unit',''),x.get('note',''),'WBS录入'])
    daily.sort(key=lambda x:(x[0],x[1],x[4]))
    sh=out.create_sheet('每日实际明细');rebuild_sheet(sh,['日期','WBS编码','工作包名称','所属子系统','工程量/活动','完成量','单位','备注','数据来源'],daily,[14,13,30,24,28,13,10,40,14],{1,2,6,7,9});sh.freeze_panes='A2';sh.auto_filter.ref=f'A1:I{sh.max_row}'
    for sh in out.worksheets:
        sh.sheet_properties.pageSetUpPr.fitToPage=True;sh.page_setup.orientation='landscape';sh.page_setup.fitToWidth=1;sh.page_setup.fitToHeight=0
    stamp=datetime.date.today().isoformat();buf=io.BytesIO();out.save(buf)
    return buf.getvalue(),PROJ_NAMES[proj]+'_工作包汇总_'+stamp+'.xlsx'


def save_daily_plan(data):
    import math, shutil
    proj, sub, date = data.get('proj'), data.get('sub'), data.get('date')
    if proj not in SUBS_OF or sub not in SUBS_OF[proj]: raise ValueError('子系统无效')
    datetime.date.fromisoformat(date)
    qty=float(data.get('qty') or 0)
    if not math.isfinite(qty) or qty<0: raise ValueError('计划量必须为非负数字')
    if proj=='titan' and not qty.is_integer():raise ValueError('每日计划数量必须为整数')
    if proj=='titan':qty=int(qty)
    cfg=sub_settings(proj)
    if sub in cfg and 'daily' in cfg[sub]:
        if date not in cfg[sub]['daily']:raise ValueError('日期不在子系统计划工期内')
        cfg[sub]['daily'][date]=qty
        if proj=='titan':
            cfg[sub]['daily']={d:(int(round(v)) if d!=date else qty) for d,v in cfg[sub]['daily'].items()}
            cfg[sub]['daily']={d:max(0,v) for d,v in cfg[sub]['daily'].items()}
        path=os.path.join(os.path.dirname(_paths(proj)['p30']),'子系统计划设置.json')
        if os.path.exists(path):shutil.copy2(path,path+'.'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'.bak')
        with open(path,'w',encoding='utf-8') as f:json.dump(cfg,f,ensure_ascii=False,indent=2)
        save_sub_plan({'proj':proj,'sub':sub,'total':sum(cfg[sub]['daily'].values()),'preserve_daily':True})
        return {'ok':True,'msg':'每日计划已更新'}
    F=_paths(proj);w=load_workbook(F['p30'])
    backup=F['p30']+'.daily_plan_backup'
    shutil.copy2(F['p30'],backup)
    if proj=='titan':
        ws=w['每日计划'];col=[3,5,7,9,11,13][SUBS_OF[proj].index(sub)]
        rows=[r for r in range(2,ws.max_row+1) if str(ws.cell(r,1).value or '').startswith('第')]
        match=next((r for r in rows if str(ws.cell(r,2).value or '')[:10]==date),None)
        if match is None: w.close();raise ValueError('请选择施工计划中的日期')
        ws.cell(match,col).value=qty
        for r in rows:ws.cell(r,col).value=max(0,int(round(float(ws.cell(r,col).value or 0))))
        total=sum(float(ws.cell(r,col).value or 0) for r in rows)
    else:
        ws=w['子系统每日计划'] if '子系统每日计划' in w.sheetnames else w.create_sheet('子系统每日计划')
        if ws.max_row==1 and ws.cell(1,1).value is None: ws.append(['日期','子系统','计划量'])
        match=next((r for r in range(2,ws.max_row+1) if str(ws.cell(r,1).value or '')[:10]==date and ws.cell(r,2).value==sub),None)
        if match:ws.cell(match,3).value=qty
        else:ws.append([date,sub,qty])
        total=sum(float(ws.cell(r,3).value or 0) for r in range(2,ws.max_row+1) if ws.cell(r,2).value==sub)
    w.save(F['p30']);w.close()
    try:
        save_sub_plan({'proj':proj,'sub':sub,'total':total,'preserve_daily':True})
    except Exception:
        shutil.copy2(backup,F['p30']);raise
    return {'ok':True,'msg':'每日计划已保存，统计与图表已更新'}


def sub_actual(proj, work_date, updates=None):
    import math, shutil
    if proj not in SUBS_OF: raise ValueError('项目无效')
    datetime.date.fromisoformat(work_date)
    F = _paths(proj)
    w30 = load_workbook(F['p30']); flows = w30['每日计量']
    w32 = load_workbook(F['p32']); details = w32['详细工程量']
    def ds(v):
        return v.strftime('%Y-%m-%d') if isinstance(v,(datetime.date,datetime.datetime)) else str(v or '')[:10]
    daily = {}
    for row in range(2, flows.max_row+1):
        if ds(flows.cell(row,1).value)==work_date:
            key=(str(flows.cell(row,2).value or ''),str(flows.cell(row,3).value or ''))
            daily[key]=daily.get(key,0)+float(flows.cell(row,4).value or 0)
    if updates is None:
        ret=[{'sub':str(details.cell(r,1).value or ''),'item':str(details.cell(r,2).value or ''),'today':daily.get((str(details.cell(r,1).value or ''),str(details.cell(r,2).value or '')),0)} for r in range(2,details.max_row+1)]
        w30.close();w32.close();return {'ok':True,'date':work_date,'details':ret}
    known={(str(details.cell(r,1).value or ''),str(details.cell(r,2).value or '')):r for r in range(2,details.max_row+1)}
    clean={}
    for sub,items in updates.items():
        for item,u in items.items():
            if (sub,item) not in known: raise ValueError('子系统明细不存在')
            qty=float(u.get('today',0) or 0)
            if not math.isfinite(qty) or qty<0: raise ValueError('完成量必须为非负数字')
            clean[(sub,item)]=qty
    for key,qty in clean.items():
        r=known[key];previous=daily.get(key,0)
        cum=float(details.cell(r,6).value or 0)-previous+qty
        details.cell(r,5).value=qty;details.cell(r,6).value=cum
        rows=[fr for fr in range(2,flows.max_row+1) if ds(flows.cell(fr,1).value)==work_date and str(flows.cell(fr,2).value or '')==key[0] and str(flows.cell(fr,3).value or '')==key[1]]
        if rows:
            flows.cell(rows[0],4).value=qty
            for fr in rows[1:]:flows.cell(fr,4).value=0
        elif qty:
            flows.append([work_date,key[0],key[1],qty,details.cell(r,3).value,cum,'工具按日期录入',''])
        item_rows=[fr for fr in range(2,flows.max_row+1) if str(flows.cell(fr,2).value or '')==key[0] and str(flows.cell(fr,3).value or '')==key[1]]
        base=cum-sum(float(flows.cell(fr,4).value or 0) for fr in item_rows)
        running=base
        for fr in sorted(item_rows,key=lambda fr:(ds(flows.cell(fr,1).value),fr)):
            running+=float(flows.cell(fr,4).value or 0);flows.cell(fr,6).value=running
    w31=load_workbook(F['p31']);status=w31['子系统状态'];overview=w30['工程量总览']
    affected={k[0] for k in clean}
    for sub in affected:
        rr=[r for k,r in known.items() if k[0]==sub]
        today=sum(float(details.cell(r,5).value or 0) for r in rr)
        cumulative=sum(float(details.cell(r,6).value or 0) for r in rr)
        for r in range(2,status.max_row+1):
            if str(status.cell(r,1).value or '')==sub:
                status.cell(r,5).value=today;status.cell(r,6).value=cumulative
        for r in range(2,overview.max_row+1):
            if canonical_sub(proj, overview.cell(r,2).value)!=sub:continue
            item=str(overview.cell(r,3).value or '')
            dr=known.get((sub,item))
            if dr:
                overview.cell(r,6).value=float(details.cell(dr,5).value or 0)
                overview.cell(r,7).value=float(details.cell(dr,6).value or 0)
    backup=os.path.join(os.path.dirname(F['p30']),'完成量修改备份',datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    os.makedirs(backup,exist_ok=True)
    books={'p30':w30,'p31':w31,'p32':w32}
    for k in books:shutil.copy2(F[k],os.path.join(backup,os.path.basename(F[k])))
    try:
        for k,book in books.items():book.save(F[k]);book.close()
        if proj=='titan':update_progress_check(proj)
        else:update_progress_check_bg(proj)
    except Exception:
        for k in books:shutil.copy2(os.path.join(backup,os.path.basename(F[k])),F[k])
        raise
    return {'ok':True,'msg':work_date+' 完成量已保存，累计与全局统计已同步'}


def allocate_plan(total, weights):
    """Largest remainder allocation: exact sum, no negative tail, integer counts."""
    from decimal import Decimal, ROUND_FLOOR
    amount=Decimal(str(total or 0));quantum=Decimal(1) if amount==amount.to_integral_value() else Decimal('0.000001')
    units=int((amount/quantum).to_integral_value())
    ws=[max(Decimal(0),Decimal(str(v or 0))) for v in weights]
    if not ws:return []
    if not sum(ws):ws=[Decimal(1)]+[Decimal(0)]*(len(ws)-1)
    denom=sum(ws);raw=[Decimal(units)*v/denom for v in ws];parts=[int(v.to_integral_value(rounding=ROUND_FLOOR)) for v in raw]
    remainder=units-sum(parts)
    order=sorted(range(len(ws)),key=lambda i:(-(raw[i]-parts[i]),i))
    for i in order[:remainder]:parts[i]+=1
    return [int(v) if quantum==1 else float(Decimal(v)*quantum) for v in parts]


def save_sub_plan(data):
    import math, shutil
    proj = data.get('proj', 'bg')
    sub = data.get('sub', '')
    if proj not in SUBS_OF or sub not in SUBS_OF[proj]:
        return {'ok': False, 'err': '项目或子系统无效'}
    raw = data.get('total')
    total = None if raw is None or raw == '' else float(raw)
    if proj=='titan' and total is not None and not total.is_integer():return {'ok':False,'err':'按条、端、个计量的计划总量必须为整数'}
    if total is not None and (not math.isfinite(total) or total < 0):
        return {'ok': False, 'err': '计划总量必须为非负数字'}
    cfg=sub_settings(proj)
    if sub in cfg and 'daily' in cfg[sub]:
        existing=sum(cfg[sub]['daily'].values());days=list(cfg[sub]['daily'])
        vals=list(cfg[sub]['daily'].values()) if data.get('preserve_daily') else allocate_plan(total,[cfg[sub]['daily'][d] for d in days])
        cfg[sub]['daily']=dict(zip(days,vals))
        path=os.path.join(os.path.dirname(_paths(proj)['p30']),'子系统计划设置.json')
        if os.path.exists(path):shutil.copy2(path,path+'.'+datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'.bak')
        with open(path,'w',encoding='utf-8') as f:json.dump(cfg,f,ensure_ascii=False,indent=2)
    F = _paths(proj)
    books = {k: load_workbook(F[k]) for k in ('p30','p31','p32')}
    ws = books['p30']['工程量总览']
    col = 5 if proj == 'titan' else 4
    matches = [r for r in range(2,ws.max_row+1) if canonical_sub(proj, ws.cell(r,2).value) == sub]
    old = sum(float(ws.cell(r,col).value or 0) for r in matches)
    if not matches: return {'ok': False, 'err': '未找到工程量计划'}
    values = [None]*len(matches) if total is None else allocate_plan(total,[ws.cell(r,col).value for r in matches])
    item_values = {}
    for r,v in zip(matches,values):
        item_values[str(ws.cell(r,3).value or '')] = v
        ws.cell(r,col).value = v
    ws31 = books['p31']['子系统状态']
    for r in range(2,ws31.max_row+1):
        if str(ws31.cell(r,1).value or '') == sub:
            representative = str(ws31.cell(r,15).value or '')
            ws31.cell(r,7).value = total if proj == 'titan' else item_values.get(representative,total)
    ws32 = books['p32']['详细工程量']
    first = True
    for r in range(2,ws32.max_row+1):
        if str(ws32.cell(r,1).value or '') != sub: continue
        item = str(ws32.cell(r,2).value or '')
        if item in item_values:
            ws32.cell(r,4).value = item_values[item]
        elif proj == 'titan' and first:
            ws32.cell(r,4).value = total
        first = False
    if proj != 'titan' and '子系统每日计划' in books['p30'].sheetnames and not data.get('preserve_daily'):
        daily_bg = books['p30']['子系统每日计划']
        plan_rows = [r for r in range(2,daily_bg.max_row+1) if canonical_sub(proj,daily_bg.cell(r,2).value)==sub]
        for r,v in zip(plan_rows,allocate_plan(total,[daily_bg.cell(r,3).value for r in plan_rows])):
            daily_bg.cell(r,3).value=v
    if proj == 'titan' and sub in BASE_SUBS['titan'] and '每日计划' in books['p30'].sheetnames:
        j = SUBS_OF[proj].index(sub)
        daily = books['p30']['每日计划']
        rows = [r for r in range(2,daily.max_row+1) if str(daily.cell(r,1).value or '').startswith('第')]
        pc = [3,5,7,9,11,13][j]
        baseline = [float(daily.cell(r,pc).value or 0) for r in rows]
        denom = sum(baseline)
        if not denom: baseline = list(titan_defs.PLAN[['fen','b7','b9','odf','duan','ces'][j]]); denom=sum(baseline)
        vals = baseline if data.get('preserve_daily') else allocate_plan(total,baseline)
        for r,v in zip(rows,vals): daily.cell(r,pc).value=v
        grand = sum(float(ws.cell(r,5).value or 0) for r in range(2,ws.max_row+1) if str(ws.cell(r,2).value or '') in SUBS_OF[proj])
        running = 0
        for r in rows:
            running += sum(float(daily.cell(r,c).value or 0) for c in (3,5,7,9,11,13))
            daily.cell(r,16).value = round(running/grand*100,2) if grand else 0
    backup = os.path.join(os.path.dirname(F['p30']), '计划修改备份', datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    os.makedirs(backup,exist_ok=True)
    for k in books: shutil.copy2(F[k],os.path.join(backup,os.path.basename(F[k])))
    try:
        for k,book in books.items(): book.save(F[k]); book.close()
        if proj == 'titan': update_progress_check(proj)
        else: update_progress_check_bg(proj)
    except Exception:
        for k in books: shutil.copy2(os.path.join(backup,os.path.basename(F[k])),F[k])
        raise
    return {'ok': True, 'msg': '计划总量已保存，工程量与进度统计已同步'}


def main():
    # 延迟 2 秒自动打开浏览器（兼容 noconsole 打包与直接双击 exe）
    def _open_browser():
        try:
            import webbrowser
            webbrowser.open('http://localhost:%d' % PORT)
        except Exception:
            pass
    import threading
    if os.environ.get('BG_TOOL_NO_BROWSER') != '1':
        threading.Timer(2.0, _open_browser).start()
    print('=' * 60)
    print('BG 项目管理系统集成小工具 已启动')
    print('请在浏览器打开: http://localhost:%d' % PORT)
    print('按 Ctrl+C 停止服务')
    print('数据同步目标: 29/30/31/32/05 号 Excel 文件(保留格式)')
    print('=' * 60)
    srv = ThreadingHTTPServer(('127.0.0.1', PORT), Handler)
    srv.serve_forever()


if __name__ == '__main__':
    main()
