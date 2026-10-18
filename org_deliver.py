# -*- coding: utf-8 -*-
"""人员组织架构 + 全系统导出 + 甲方交付文件中心（BG/泰坦两项目独立）
数据层：26_项目人员组织架构.xlsx（人员信息/分组定义/岗位定义 三个 sheet）
导出层：xlsx/csv/pdf/docx 统一生成，正式交付格式（页眉页脚页码/签署栏）
交付层：甲方交付资料目录 + 版本登记表（版本管理）
"""
import os, io, csv, datetime, json, sys, base64, binascii, re
from html import escape
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from docx import Document as DocxDocument
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

# 兼容 PyInstaller onefile：frozen 时 __file__ 指向 _MEIPASS 临时解包目录，必须用 exe 所在目录
if getattr(sys, 'frozen', False):
    DIR = os.path.dirname(os.path.abspath(sys.executable))  # exe 所在目录 = 项目管理小工具
else:
    DIR = os.path.dirname(os.path.abspath(__file__))       # 项目管理小工具 目录
ROOT = os.path.dirname(DIR)                             # 文件体系根

NAVY = '1F3A5F'; ORANGE = 'E07B39'; LIGHT = 'F2F5F9'
HDR_FILL = PatternFill('solid', fgColor=NAVY)
HDR_FONT = Font(name='微软雅黑', bold=True, color='FFFFFF', size=10)
BODY_FONT = Font(name='微软雅黑', size=10)
BORDER = Border(*[Side(style='thin', color='C9D2DC')] * 4)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

ORG_HEADERS = ['人员ID', '姓名', '岗位', '所属分组', '角色', '联系电话', '负责区域',
               '负责子系统', '人员状态', '排序号', '备注', '创建时间', '更新时间', '所属项目',
               '性别', '工种', '证件号', '入场日期', '个人简介', '照片文件']
ORG_KEYS = ['pid', 'name', 'job', 'group', 'role', 'phone', 'area', 'sub', 'status', 'sort',
            'note', 'created', 'updated', 'proj', 'gender', 'trade', 'id_no', 'entry_date', 'bio', 'photo']
ORG_WIDTHS = [10, 10, 12, 14, 14, 13, 16, 16, 9, 7, 20, 11, 11, 12, 8, 12, 20, 12, 36, 24]
GROUP_HEADERS = ['分组名', '排序号', '说明']
JOB_HEADERS = ['岗位名', '排序号', '说明']
ROLE_HEADERS = ['角色名', '排序号', '说明']

DEFAULT_GROUPS = [
    ('项目管理组', 1, '项目负责人/技术/安全/质量 管理层'),
    ('施工管理组', 2, '施工负责人与现场协调'),
    ('网络系统施工组', 3, '网络系统施工班组'),
    ('综合布线施工组', 4, '综合布线施工班组'),
    ('光缆施工组', 5, '光缆施工班组（泰坦主线）'),
    ('监控系统施工组', 6, '视频监控施工班组'),
    ('门禁系统施工组', 7, '门禁/一卡通施工班组'),
    ('其他施工组', 8, '其他劳务/辅助班组'),
]
DEFAULT_JOBS = ['项目经理', '技术负责人', '施工负责人', '安全负责人', '质量负责人',
                '现场协调', '施工员', '技术员', '安全员', '质检员', '测试员', '仓库管理员']
ROLES = ['项目负责人', '技术负责人', '施工负责人', '安全负责人', '质量负责人', '其他']
STATUS_LIST = ['在职', '停用', '已删除']


def _paths(proj):
    if proj == 'titan':
        d = os.path.join(ROOT, '泰坦IDC')
        return {
            'p26': os.path.join(d, '26_项目人员组织架构.xlsx'),
            'export': os.path.join(d, '导出文件'),
            'deliver': os.path.join(d, '甲方交付资料'),
            'photos': os.path.join(d, '人员照片'),
        }
    root=ROOT if proj=='bg' else os.path.join(ROOT,'自定义项目',proj)
    return {
        'p26': os.path.join(root, '26_项目人员组织架构.xlsx'),
        'export': os.path.join(root, '导出文件'),
        'deliver': os.path.join(root, '甲方交付资料'),
        'photos': os.path.join(root, '人员照片'),
    }


def rebuild_sheet(ws, headers, rows, widths, center_cols=None):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.border = BORDER; cell.alignment = CENTER
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = BODY_FONT; cell.border = BORDER
            cell.alignment = CENTER if (center_cols and c in center_cols) else LEFT
    for r in range(len(rows) + 2, ws.max_row + 1):
        ws.delete_rows(len(rows) + 2)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def ensure_org_file(proj):
    p = _paths(proj)['p26']
    if os.path.exists(p):
        return p
    os.makedirs(os.path.dirname(p), exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = '人员信息'
    rebuild_sheet(ws, ORG_HEADERS, [], ORG_WIDTHS, center_cols={1, 4, 5, 9, 10, 12, 13, 14})
    ws2 = wb.create_sheet('分组定义')
    rebuild_sheet(ws2, GROUP_HEADERS, [list(g) for g in DEFAULT_GROUPS], [24, 7, 46], center_cols={2})
    ws3 = wb.create_sheet('岗位定义')
    rebuild_sheet(ws3, JOB_HEADERS, [[j, i + 1, ''] for i, j in enumerate(DEFAULT_JOBS)], [16, 7, 46], center_cols={2})
    ws4 = wb.create_sheet('角色定义')
    rebuild_sheet(ws4, ROLE_HEADERS, [[r, i + 1, ''] for i, r in enumerate(ROLES)], [16, 7, 46], center_cols={2})
    wb.save(p)
    return p


def load_org(proj):
    """读 26 号表：{people:[{...}], groups:[{name,sort,note}], jobs:[{name,sort,note}]}"""
    p = ensure_org_file(proj)
    wb = load_workbook(p)
    people = []
    ws = wb['人员信息']
    for r in range(2, ws.max_row + 1):
        row = [ws.cell(row=r, column=c).value for c in range(1, len(ORG_HEADERS) + 1)]
        if all(v is None or str(v).strip() == '' for v in row):
            continue
        people.append(dict(zip(ORG_KEYS, row)))
    groups, jobs, roles = [], [], []
    ws2 = wb['分组定义']
    for r in range(2, ws2.max_row + 1):
        name = ws2.cell(row=r, column=1).value
        if not name:
            continue
        groups.append({'name': str(name), 'sort': ws2.cell(row=r, column=2).value or 99, 'note': ws2.cell(row=r, column=3).value or ''})
    ws3 = wb['岗位定义']
    for r in range(2, ws3.max_row + 1):
        name = ws3.cell(row=r, column=1).value
        if not name:
            continue
        jobs.append({'name': str(name), 'sort': ws3.cell(row=r, column=2).value or 99, 'note': ws3.cell(row=r, column=3).value or ''})
    if '角色定义' in wb.sheetnames:
        ws4 = wb['角色定义']
        for r in range(2, ws4.max_row + 1):
            name = ws4.cell(row=r, column=1).value
            if not name:
                continue
            roles.append({'name': str(name), 'sort': ws4.cell(row=r, column=2).value or 99, 'note': ws4.cell(row=r, column=3).value or ''})
    if not roles:
        roles = [{'name': name, 'sort': i + 1, 'note': ''} for i, name in enumerate(ROLES)]
    # 兼容历史数据中的自定义角色：即使旧文件尚无“角色定义”sheet，也自动带入管理列表。
    known_roles = {x['name'] for x in roles}
    for person in people:
        role_name = str(person.get('role') or '').strip()
        if role_name and role_name not in known_roles:
            roles.append({'name': role_name, 'sort': 99, 'note': '由现有人员信息自动加入'})
            known_roles.add(role_name)
    groups.sort(key=lambda x: (x['sort'] if isinstance(x['sort'], (int, float)) else 99, x['name']))
    jobs.sort(key=lambda x: (x['sort'] if isinstance(x['sort'], (int, float)) else 99, x['name']))
    roles.sort(key=lambda x: (x['sort'] if isinstance(x['sort'], (int, float)) else 99, x['name']))
    return {'people': people, 'groups': groups, 'jobs': jobs, 'roles': roles}


def next_pid(people, proj):
    pre = 'BG-P' if proj == 'bg' else 'T-P'
    mx = 0
    for p in people:
        v = str(p['pid'] or '')
        if v.startswith(pre):
            try:
                mx = max(mx, int(v[len(pre):]))
            except Exception:
                pass
    return pre + '%02d' % (mx + 1)


def save_org(proj, data):
    """全量写回 26 号表。data={people:[...], groups:[...], jobs:[...]}（与 load_org 同构）
    规则：人员ID 自动补全；创建时间保留原值；更新时间=现在；排序默认 999；状态默认 在职；
    分组/岗位列表为空时保留原表定义（防止前端误传空清空配置）"""
    p = ensure_org_file(proj)
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    g = data.get('groups') or []
    j = data.get('jobs') or []
    rdefs = data.get('roles') or []
    if not g or not j or not rdefs:
        _old = load_org(proj)
        if not g:
            g = _old['groups']
        if not j:
            j = _old['jobs']
        if not rdefs:
            rdefs = _old['roles']
    people = []
    prefix = 'BG-P' if proj == 'bg' else 'T-P'
    for pd in data.get('people', []):
        if not (pd.get('name') or '').strip():
            continue
        rec = dict(pd)
        if not rec.get('status'):
            rec['status'] = '在职'
        if not rec.get('created'):
            rec['created'] = now
        rec['updated'] = now
        rec['proj'] = 'titan' if proj == 'titan' else 'bg'
        rec['sort'] = rec.get('sort')
        if rec['sort'] is None:
            rec['sort'] = 999
        rec['_old_pid'] = str(rec.get('pid') or '')
        people.append(rec)
    # 在职/停用人员先连续编号，已删除历史排在其后；排序列与新 ID 同步连续。
    people.sort(key=lambda x: (1 if x.get('status') == '已删除' else 0,
                               x['sort'] if isinstance(x['sort'], (int, float)) else 999, str(x['name'])))
    pid_map = {}
    for index, rec in enumerate(people, 1):
        new_pid = prefix + '%02d' % index
        if rec['_old_pid']:
            pid_map[rec['_old_pid']] = new_pid
        rec['pid'] = new_pid
        rec['sort'] = index if rec.get('status') != '已删除' else rec.get('sort', index)
        rec.pop('_old_pid', None)
        photo_data = rec.pop('photo_data', '') or ''
        if photo_data:
            match = re.fullmatch(r'data:image/(jpeg|png|webp);base64,(.+)', str(photo_data), re.DOTALL)
            if not match: raise ValueError('人员照片仅支持 JPG、PNG 或 WEBP')
            try: raw = base64.b64decode(match.group(2), validate=True)
            except (ValueError, binascii.Error): raise ValueError('人员照片数据无效')
            if len(raw) > 5 * 1024 * 1024: raise ValueError('单张人员照片不能超过 5MB')
            signatures = {'jpeg': b'\xff\xd8\xff', 'png': b'\x89PNG\r\n\x1a\n', 'webp': b'RIFF'}
            if not raw.startswith(signatures[match.group(1)]): raise ValueError('人员照片格式与文件内容不一致')
            photo_dir = _paths(proj)['photos']; os.makedirs(photo_dir, exist_ok=True)
            ext = 'jpg' if match.group(1) == 'jpeg' else match.group(1)
            filename = '%s_%s.%s' % (new_pid, datetime.datetime.now().strftime('%Y%m%d%H%M%S%f'), ext)
            with open(os.path.join(photo_dir, filename), 'wb') as photo_file: photo_file.write(raw)
            rec['photo'] = filename
    wb = load_workbook(p)
    ws = wb['人员信息']
    rows = [[rec.get(k) for k in ORG_KEYS] for rec in people]
    rebuild_sheet(ws, ORG_HEADERS, rows, ORG_WIDTHS, center_cols={1, 4, 5, 9, 10, 12, 13, 14, 15, 18})
    ws2 = wb['分组定义']
    rebuild_sheet(ws2, GROUP_HEADERS, [[x.get('name'), x.get('sort', 99), x.get('note', '')] for x in g], [24, 7, 46], center_cols={2})
    ws3 = wb['岗位定义']
    rebuild_sheet(ws3, JOB_HEADERS, [[x.get('name'), x.get('sort', 99), x.get('note', '')] for x in j], [16, 7, 46], center_cols={2})
    ws4 = wb['角色定义'] if '角色定义' in wb.sheetnames else wb.create_sheet('角色定义')
    rebuild_sheet(ws4, ROLE_HEADERS, [[x.get('name'), x.get('sort', 99), x.get('note', '')] for x in rdefs], [16, 7, 46], center_cols={2})
    wb.save(p)
    # 32号表“详细工程量”的施工人员保存的是人员 ID，重编号时同步替换引用。
    detail_path = os.path.join(ROOT if proj == 'bg' else (os.path.join(ROOT,'泰坦IDC') if proj == 'titan' else os.path.join(ROOT,'自定义项目',proj)), '32_子系统施工步骤与明细表.xlsx')
    if pid_map and os.path.exists(detail_path):
        detail_book = load_workbook(detail_path)
        if '详细工程量' in detail_book.sheetnames:
            detail_sheet = detail_book['详细工程量']; changed = False
            for row_no in range(2, detail_sheet.max_row + 1):
                raw = detail_sheet.cell(row_no, 13).value
                if not raw: continue
                try: ids = json.loads(raw)
                except Exception: continue
                mapped = [pid_map.get(str(pid), str(pid)) for pid in ids]
                if mapped != ids:
                    detail_sheet.cell(row_no, 13).value = json.dumps(mapped, ensure_ascii=False); changed = True
            if changed: detail_book.save(detail_path)
        detail_book.close()
    return {'ok': True, 'people': people}


def sync_org_from_people(proj, people_rows):
    """人员管理（29 号表每日到场）保存后联动：不在组织架构中的姓名自动加入 26 号表。
    入参 people_rows 为二维数组（与 save_payload 中 people_rows 同构：
    列序 date,name,pos,src,team,wt,area,content,hours）。
    - 岗位 job = pos；分组 group = team（不存在则自动创建分组）
    - 同名已存在（含已删除历史）不重复新增，保留历史可追溯
    - 返回本次新增人数
    """
    names = {}
    def person_key(value):
        import unicodedata
        return unicodedata.normalize('NFKC', str(value or '').strip()).casefold()
    for row in people_rows:
        if not row:
            continue
        nm = str(row[1] or '').strip() if len(row) > 1 else ''
        if not nm:
            continue
        key = person_key(nm)
        if key not in names:
            names[key] = (nm, row)
    if not names:
        return 0
    org = load_org(proj)
    existing = {}
    for person in org['people']:
        key = person_key(person.get('name'))
        if key and (key not in existing or (existing[key].get('status') == '已删除' and person.get('status') != '已删除')):
            existing[key] = person
    add = [key for key in names if key not in existing]
    restored = 0
    for key, (nm, row) in names.items():
        rec = existing.get(key)
        if rec and rec.get('status') == '已删除':
            rec['name'] = nm; rec['status'] = '在职'
            rec['job'] = str(row[2] or '').strip() or rec.get('job') or '施工人员'
            rec['group'] = str(row[4] or '').strip() or rec.get('group') or '未分组'
            rec['area'] = str(row[6] or '').strip() or rec.get('area')
            restored += 1
    if not add and not restored:
        return 0
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    groups = list(org['groups'])
    gnames = {str(g['name']) for g in groups}
    jobs = list(org['jobs'])
    jnames = {str(j['name']) for j in jobs}
    for key in add:
        nm, row = names[key]
        pos = str(row[2] or '').strip() if len(row) > 2 else ''
        pos = pos or '施工人员'
        team = str(row[4] or '').strip() if len(row) > 4 else ''
        area = str(row[6] or '').strip() if len(row) > 6 else ''
        date = str(row[0] or '').strip() if row else ''
        if team and team not in gnames:
            groups.append({'name': team, 'sort': 99, 'note': '由人员管理自动创建'})
            gnames.add(team)
        if pos and pos not in jnames:
            jobs.append({'name': pos, 'sort': 99, 'note': '由人员管理自动创建'})
            jnames.add(pos)
        org['people'].append({
            'pid': next_pid(org['people'], proj),
            'name': nm, 'job': pos, 'group': team or '未分组',
            'role': '其他', 'phone': None,
            'area': area or None,
            'sub': None, 'status': '在职', 'sort': 999,
            'note': '由人员管理自动同步（' + date + '）',
            'created': now, 'updated': now, 'proj': 'titan' if proj == 'titan' else 'bg'})
    save_org(proj, {'people': org['people'], 'groups': groups, 'jobs': jobs})
    return len(add) + restored


def org_stats(org):
    """人员统计（动态计算，禁止写死）"""
    ps = [p for p in org['people'] if str(p.get('status') or '') != '已删除']
    total = len(ps)
    active = sum(1 for p in ps if p.get('status') in (None, '', '在职'))
    stopped = sum(1 for p in ps if p.get('status') == '停用')
    by_role = {}
    for p in ps:
        r = str(p.get('role') or '其他').strip() or '其他'
        by_role[r] = by_role.get(r, 0) + 1
    by_group = {}
    for p in ps:
        g = str(p.get('group') or '未分组').strip() or '未分组'
        by_group[g] = by_group.get(g, 0) + 1
    by_job = {}
    for p in ps:
        j = str(p.get('job') or '未设置').strip() or '未设置'
        by_job[j] = by_job.get(j, 0) + 1
    by_sub = {}
    for p in ps:
        s = str(p.get('sub') or '').strip()
        if s:
            for ss in s.replace('，', ',').split(','):
                ss = ss.strip()
                if ss:
                    by_sub[ss] = by_sub.get(ss, 0) + 1
    subs_resp = len([p for p in ps if str(p.get('sub') or '').strip()])
    return {
        'total': total, 'active': active, 'stopped': stopped,
        'by_role': by_role, 'by_group': by_group, 'by_job': by_job, 'by_sub': by_sub,
        'subs_resp': subs_resp,
    }


# ============ 统一导出 ============
PROJ_NAMES = {'bg': 'BG弱电改造项目', 'titan': '泰坦IDC数据中心光缆施工工程'}


def _bootstrap(proj):
    return _server().bootstrap(proj)


def _server():
    # Reuse the running service, including its custom projects and data root.
    main = sys.modules.get('__main__')
    if main is not None and hasattr(main, 'bootstrap') and hasattr(main, 'SUBS_OF'):
        return main
    import server
    return server


def _fmt(v):
    if v is None:
        return ''
    if isinstance(v, float):
        if v == int(v):
            return str(int(v))
        return ('%.2f' % v).rstrip('0').rstrip('.')
    return str(v)


def collect_table(proj, table, include_deleted=False):
    d = _bootstrap(proj)
    title, headers, rows = _collect_table(proj, table, include_deleted, d)
    settings = d.get('sub_settings') or {}
    sub_columns = [i for i, h in enumerate(headers) if h in ('子系统', '所属系统', '负责子系统')]
    for row in rows:
        for col in sub_columns:
            if col < len(row):
                row[col] = (settings.get(str(row[col])) or {}).get('name') or row[col]
    return title, headers, rows


def _collect_table(proj, table, include_deleted, d):
    """返回 (标题, 表头列表, 数据行[list])，数据源与 API bootstrap 同一逻辑"""
    if table in ('daily_report', 'test_record'):
        daily = _server().load_daily(proj)
        rows = [list(r) for r in daily.get('records', [])
                if table == 'daily_report' or (len(r)>2 and r[2]=='测试')]
        return ('施工日报' if table=='daily_report' else '测试记录'), daily['headers'], rows
    if table == 'construction_plan':
        headers = ['日期', '子系统', '计划量', '单位', '计划开始', '计划结束', '实际开工时间', '实际完工时间']
        rows = []
        for sub, schedule in d.get('sub_schedules', {}).items():
            state = d.get('sys_status', {}).get(sub, {})
            cfg = d.get('sub_settings', {}).get(sub, {})
            for date, qty in sorted(schedule.items()):
                rows.append([date, sub, qty, state.get('unit',''), cfg.get('start') or state.get('ps',''),
                             cfg.get('finish') or state.get('pe',''), cfg.get('actual_start',''), cfg.get('actual_finish','')])
        return '施工进度计划', headers, sorted(rows, key=lambda r:(r[0],r[1]))
    if table == 'org':
        org = load_org(proj)
        ps = org['people']
        if not include_deleted:
            ps = [p for p in ps if str(p.get('status') or '') != '已删除']
        hdr = ORG_HEADERS
        rows = [[p.get(k) for k in ORG_KEYS] for p in ps]
        return '人员组织架构', hdr, rows
    if table == 'people':
        hdr = ['日期', '姓名', '岗位', '人员来源', '班组', '工作类型', '负责区域', '工作内容', '工时']
        hdr.append('备注')
        rows=[[r.get(k,'') for k in ('date','name','pos','src','team','wt','area','content','hours','note')] for r in d.get('people',[])]
        # 日报按日期集中排列，保留当天人员录入顺序。
        rows.sort(key=lambda row: _fmt(row[0]))
        counts = {}
        for row in rows:
            date = _fmt(row[0])
            counts[date] = counts.get(date, 0) + bool(_fmt(row[1]).strip())
        hdr.append('每日到场人数统计')
        rows = [row + [counts[_fmt(row[0])]] for row in rows]
        return '项目人员管理-日报明细', hdr, rows
    if table == 'attendance':
        hdr = ['日期', '合计人数', '中方', '当地', '施工', '安装', '调试', '验收', '管理', '工作内容', '状态']
        rows = []
        for date, a in sorted(d.get('attendance', {}).items()):
            rows.append([date, a.get('total'), a.get('cn'), a.get('local'), a.get('build'),
                         a.get('install'), a.get('debug'), a.get('accept'), a.get('manage'), a.get('work'), a.get('status')])
        return '每日到场汇总', hdr, rows
    if table == 'qty':
        hdr = ['编号', '子系统', '工程量项', '计划总量', '单位', '今日完成', '累计完成', '计划说明', '备注', '责任人', '状态']
        rows = [[q.get('no'), q.get('sub'), q.get('item'), q.get('total'), q.get('unit'),
                 q.get('today'), q.get('cum'), q.get('plan'), q.get('note'), q.get('resp'), q.get('status')] for q in d.get('qty', [])]
        return '工程量总览', hdr, rows
    if table == 'metrics':
        hdr = ['日期', '子系统', '明细项', '今日量', '单位', '累计量']
        rows = [list(r[:6]) for r in d.get('metrics', [])]
        return '每日计量流水', hdr, rows
    if table == 'sys_status':
        hdr = ['子系统', '当前状态', '计划开始', '计划结束', '今日完成', '累计完成', '计划总量', '单位', '负责人', '更新时间', '备注', '计量口径', '数据来源', '实际开工时间', '实际完工时间', '完成率%']
        rows = [[k, v.get('status'), v.get('ps'), v.get('pe'), v.get('today'), v.get('cum'),
                 v.get('total'), v.get('unit'), v.get('resp'), v.get('time'), v.get('note'), v.get('metric'), v.get('src'),
                 v.get('actual_start',''),v.get('actual_finish',''),round(float(v.get('cum') or 0)/float(v['total'])*100,1) if v.get('total') else 0]
                for k, v in d.get('sys_status', {}).items()]
        return '子系统状态总览', hdr, rows
    if table == 'steps':
        hdr = ['子系统', '序号', '施工步骤', '计划开始', '计划完成', '当前状态', '完成日期', '责任人', '备注', '填报日期', '当日完成数量', '当日备注']
        rows = []
        for s in d.get('steps', []):
            quantities=s.get('quantities') or {};notes=s.get('daily_notes') or {}
            for date in sorted(set(quantities)|set(notes)) or ['']:
                rows.append([s.get(k,'') for k in ('sub','no','step','ps','pe','status','finish','resp','note')]+[date,quantities.get(date,''),notes.get(date,'')])
        return '施工步骤', hdr, rows
    if table == 'details':
        hdr = ['子系统','明细项','单位','计划总量','今日完成','累计完成','数据来源','备注','计划开始','计划完成']
        rows = [[x.get('sub'), x.get('item'), x.get('unit'), x.get('total'),
                 x.get('today'), x.get('cum'), x.get('src'), x.get('note'),x.get('ps'),x.get('pe')] for x in d.get('details', [])]
        return '详细工程量', hdr, rows
    if table == 'acceptance':
        hdr = ['子系统', '验收类型', '状态', '验收日期', '验收人', '验收结论', '未通过原因', '整改完成日期']
        rows = [[a.get('sub'), a.get('type'), a.get('status'), a.get('date'), a.get('who'),
                 a.get('concl'), a.get('reason'), a.get('fixdate')] for a in d.get('acceptance', [])]
        return '验收记录', hdr, rows
    if table == 'issues':
        hdr = ['编号', '问题描述', '发现阶段', '发现人', '所属系统', '问题分类', '问题原因', '影响范围',
               '影响工期', '影响成本', '影响质量', '责任方', '解决方案', '责任人', '计划完成日期',
               '状态', '是否超期', '关闭依据', '是否已处理', '处理方式']
        if proj == 'titan':
            hdr = ['编号','问题描述','问题类型','优先级','提出日期','位置','登记依据','子系统','影响区域/范围','责任部门','解决方案','负责人','计划解决日期','实际解决日期','状态','是否关闭','涉及合同','备注','是否已处理','处理方式']
        hdr.append('问题处理结果')
        rows = [(list(r)+['']*21)[:21] for r in d.get('issues', [])]
        return '问题清单', hdr, rows
    if table == 'todos':
        hdr = ['编号', '事项', '优先级', '责任人', '截止', '状态', '备注', '是否已处理', '处理方式', '所属子系统', '影响区域/范围']
        rows = [list(r[:11]) for r in d.get('todos', [])]
        return '待办事项', hdr, rows
    if table == 'risks':
        hdr = ['风险编号', '所属子系统', '风险描述', '风险类别', '发生概率', '影响程度', '风险等级', '触发条件',
               '应对策略', '应对措施', '责任人', '计划复查日期', '状态', '转化问题编号', '经验总结', '影响区域/范围', '影响估算', '估算单位/类型']
        rows = [list(r[:18]) for r in d.get('risks', [])]
        return '风险登记册', hdr, rows
    if table == 'audit':
        hdr = ['日期', '检查类型', '问题编号', '问题描述', '图片文件', '是否已处理', '处理时间', '处理说明', '登记人', '登记时间']
        rows = [list(r[:10]) for r in d.get('audit', [])]
        return '工程审计登记', hdr, rows
    if table == 'progress_check':
        hdr = ['日期序号', '日期', '子系统', '计划当日量', '实际当日量', '当日差额', '累计计划量', '累计实际量', '累计差额', '完成率%', '判定状态']
        rows = [list(r[:11]) for r in d.get('progress_check', [])]
        return '施工进度检查（每日）', hdr, rows
    if table == 'progress_sub':
        hdr = ['子系统', '计划总量', '累计实际', '完成率%', '剩余量', '偏差(±)', '判定', '计划工期(90天D+X)', '单位', '数据来源']
        rows = [list(r[:10]) for r in d.get('progress_sub', [])]
        return '施工进度检查（子系统）', hdr, rows
    if table == 'progress_total':
        hdr = ['检查项', '数值', '说明']
        rows = [list(r[:3]) for r in d.get('progress_total', [])]
        return '施工进度检查（总体）', hdr, rows
    if table == 'titan_plan' and proj == 'titan':
        meta = d.get('meta', {})
        plan = meta.get('plan', {})
        days = plan.get('days', [])
        subs = meta.get('subs', [])
        keys = meta.get('plan_keys', ['fen','b7','b9','odf','duan','ces'])
        hdr = ['施工日', '日期'] + [d.get('sub_settings', {}).get(sub, {}).get('name', sub) for sub in subs] + ['累计计划', '计划人数']
        rows = []
        for i in range(len(days)):
            rows.append(['D' + str(i + 1), days[i]] + [plan.get(key, [0]*len(days))[i] for key in keys] + [plan.get('cum', [0]*len(days))[i], plan.get('people', [0]*len(days))[i]])
        return '施工进度计划', hdr, rows
    raise ValueError('未知表格: ' + table)


def _title_info(proj, title):
    server = _server()
    profile = server.pm_core.profile(server.DIR, proj, server.PROJ_NAMES.get(proj,proj), server.SUBS_OF.get(proj,[]))
    return {
        'proj_name': profile.get('project_name') or server.PROJ_NAMES.get(proj, proj),
        'title': title,
        'org': '安伟达-国轩项目部',
        'editor': '',
        'version': 'V1.0',
        'date': datetime.date.today().strftime('%Y-%m-%d'),
        'no': 'PRJ-%s' % ('TITAN' if proj == 'titan' else 'BG'),
    }


def export_to(proj, table, fmt, include_deleted=False, out_dir=None, editor='', version='V1.0'):
    """统一导出入口。fmt: xlsx/csv/pdf/docx。返回生成文件绝对路径"""
    title, hdr, rows = collect_table(proj, table, include_deleted)
    info = _title_info(proj, title)
    info['editor'] = editor
    info['version'] = version
    if out_dir is None:
        out_dir = _paths(proj)['export']
    os.makedirs(out_dir, exist_ok=True)
    fname = '%s_%s_%s.%s' % (PROJ_NAMES[proj].replace('/', '_'), title, info['date'].replace('-', ''), fmt)
    fpath = os.path.join(out_dir, fname)
    if fmt == 'xlsx':
        _to_xlsx(fpath, title, hdr, rows, info)
    elif fmt == 'csv':
        _to_csv(fpath, hdr, rows)
    elif fmt == 'pdf':
        _to_pdf(fpath, title, hdr, rows, info)
    elif fmt == 'docx':
        _to_docx(fpath, title, hdr, rows, info)
    else:
        raise ValueError('不支持的格式: ' + fmt)
    return fpath


def _auto_width(hdr, rows, max_w=38):
    widths = []
    for c in range(len(hdr)):
        w = max([len(str(hdr[c]))] + [len(str(r[c])) for r in rows] if rows else [len(str(hdr[c]))])
        widths.append(min(max(8, w * 1.8), max_w))
    return widths


def _to_xlsx(fpath, title, hdr, rows, info):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title[:28]
    # 标题块
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(hdr))
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(name='微软雅黑', bold=True, size=15, color=NAVY); c.alignment = CENTER
    ws.row_dimensions[1].height = 26
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(hdr))
    c2 = ws.cell(row=2, column=1, value='项目：%s    文档编号：%s    版本：%s    日期：%s    编制单位：%s    编制人：%s' %
                 (info['proj_name'], info['no'], info['version'], info['date'], info['org'], info['editor'] or '—'))
    c2.font = Font(name='微软雅黑', size=9, color='666666'); c2.alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[2].height = 18
    # 表头
    r0 = 3
    for cidx, h in enumerate(hdr, 1):
        cell = ws.cell(row=r0, column=cidx, value=h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.border = BORDER; cell.alignment = CENTER
    # 数据
    for i, row in enumerate(rows):
        r = r0 + 1 + i
        for cidx, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=cidx, value=v)
            cell.font = BODY_FONT; cell.border = BORDER
            cell.alignment = CENTER if cidx <= 6 else LEFT
    # 列宽
    widths = _auto_width(hdr, rows)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    if title == '项目人员管理-日报明细':
        # 与人员到场日报模板一致：日期和到场人数按天合并显示。
        for i, width in enumerate([18, 38, 11, 9, 13, 10, 13, 28, 8, 10, 24], 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        black_border = Border(*[Side(style='thin', color='000000')] * 4)
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=11):
            for cell in row:
                cell.border = black_border
                if cell.row > 3:
                    cell.font = Font(name='宋体', size=11)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
        start = 0
        while start < len(rows):
            end = start + 1
            while end < len(rows) and _fmt(rows[end][0]) == _fmt(rows[start][0]):
                end += 1
            first, last = start + 4, end + 3
            if _fmt(rows[start][0]).strip():
                if last > first:
                    for col in (1, 11):
                        ws.merge_cells(start_row=first, start_column=col, end_row=last, end_column=col)
                ws.cell(first, 11, '=COUNTA(B%d:B%d)' % (first, last))
                ws.cell(first, 11).number_format = '0"人"'
            start = end
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.print_title_rows = '1:3'
        ws.print_area = 'A1:K%d' % ws.max_row
    ws.freeze_panes = 'A%d' % (r0 + 1)
    wb.save(fpath)


def _to_csv(fpath, hdr, rows):
    with io.open(fpath, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(hdr)
        for r in rows:
            w.writerow([_fmt(v) for v in r])


def _to_pdf(fpath, title, hdr, rows, info):
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,
                                    PageBreak, KeepTogether)
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.platypus.frames import Frame
    from reportlab.platypus.doctemplate import PageTemplate, BaseDocTemplate
    try:
        pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
        FONT = 'STSong-Light'
    except Exception:
        FONT = 'Helvetica'
    w, h = landscape(A4)
    def _ph(canvas, doc):
        canvas.saveState()
        canvas.setFont(FONT, 8); canvas.setFillColor(colors.HexColor('#1F3A5F'))
        canvas.drawString(12 * mm, h - 9 * mm, '%s   |   %s   |   %s   |   版本 %s' % (info['proj_name'], info['title'], info['no'], info['version']))
        canvas.setStrokeColor(colors.HexColor('#1F3A5F')); canvas.setLineWidth(0.6)
        canvas.line(12 * mm, h - 11 * mm, w - 12 * mm, h - 11 * mm)
        canvas.setFont(FONT, 8); canvas.setFillColor(colors.HexColor('#666666'))
        canvas.drawCentredString(w / 2, 8 * mm, '第 %d 页' % doc.page)
        canvas.drawRightString(w - 12 * mm, 8 * mm, info['date'])
        canvas.setStrokeColor(colors.HexColor('#C9D2DC')); canvas.setLineWidth(0.4)
        canvas.line(12 * mm, 11 * mm, w - 12 * mm, 11 * mm)
        canvas.restoreState()
    doc = BaseDocTemplate(fpath, pagesize=landscape(A4),
                          leftMargin=12 * mm, rightMargin=12 * mm, topMargin=16 * mm, bottomMargin=14 * mm,
                          title='%s-%s' % (info['proj_name'], info['title']), author=info['org'])
    doc.addPageTemplates([PageTemplate(id='p', frames=[Frame(12 * mm, 14 * mm, w - 24 * mm, h - 32 * mm, id='f')], onPage=_ph)])
    styles = getSampleStyleSheet()
    ts = ParagraphStyle('title', parent=styles['Title'], fontName=FONT, fontSize=15, textColor=colors.HexColor('#1F3A5F'), spaceAfter=4, alignment=1)
    ns = ParagraphStyle('note', parent=styles['Normal'], fontName=FONT, fontSize=8, textColor=colors.HexColor('#666666'), spaceAfter=8)
    story = [Paragraph(title, ts),
             Paragraph('项目：%s    文档编号：%s    版本：%s    日期：%s    编制单位：%s    编制人：%s' %
                       (info['proj_name'], info['no'], info['version'], info['date'], info['org'], info['editor'] or '—'), ns)]
    data = [[Paragraph('<b>%s</b>' % escape(str(x)), ParagraphStyle('h', fontName=FONT, fontSize=8, wordWrap='CJK', textColor=colors.white)) for x in hdr]]
    for row in rows:
        data.append([Paragraph(escape(_fmt(v)).replace('\n','<br/>'), ParagraphStyle('c', fontName=FONT, fontSize=8, wordWrap='CJK')) for v in row])
    col_w = [(w - 24 * mm - 12) / max(1, len(hdr))] * len(hdr)
    tb = Table(data, colWidths=col_w, repeatRows=1, splitInRow=1)
    tb.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F3A5F')),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#C9D2DC')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F2F5F9')]),
        ('LEFTPADDING', (0, 0), (-1, -1), 3), ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 2), ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(tb)
    # 签署栏
    story.append(Spacer(1, 10))
    sign = Table([['编制：%s' % (info['editor'] or '——————'), '审核：——————', '批准：——————'],
                  ['日期：%s' % info['date'], '日期：——————', '日期：——————']], colWidths=[(w - 24 * mm) / 3] * 3)
    sign.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#999999')),
                              ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                              ('FONTNAME', (0, 0), (-1, -1), FONT), ('FONTSIZE', (0, 0), (-1, -1), 8)]))
    story.append(KeepTogether(sign))
    doc.build(story)


def _to_docx(fpath, title, hdr, rows, info):
    doc = DocxDocument()
    st = doc.styles['Normal']
    st.font.name = '微软雅黑'; st.font.size = Pt(9)
    st._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title); r.bold = True; r.font.size = Pt(16); r.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
    p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run('项目：%s    文档编号：%s    版本：%s    日期：%s    编制单位：%s    编制人：%s' %
                    (info['proj_name'], info['no'], info['version'], info['date'], info['org'], info['editor'] or '—'))
    r2.font.size = Pt(8); r2.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    tb = doc.add_table(rows=1, cols=len(hdr))
    tb.style = 'Table Grid'; tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, hh in enumerate(hdr):
        cell = tb.rows[0].cells[i]
        cell.text = ''
        rr = cell.paragraphs[0].add_run(hh); rr.bold = True; rr.font.size = Pt(8)
        rr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        shd = cell._tc.get_or_add_tcPr()
        from docx.oxml import OxmlElement
        el = OxmlElement('w:shd'); el.set(qn('w:val'), 'clear'); el.set(qn('w:fill'), '1F3A5F')
        shd.append(el)
    for row in rows:
        cells = tb.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ''
            rr = cells[i].paragraphs[0].add_run(_fmt(v)); rr.font.size = Pt(8)
    # 签署栏
    doc.add_paragraph()
    stb = doc.add_table(rows=2, cols=3)
    stb.style = 'Table Grid'
    for i, txt in enumerate(['编制：%s' % (info['editor'] or '——————'), '审核：——————', '批准：——————']):
        stb.rows[0].cells[i].text = txt
        stb.rows[0].cells[i].paragraphs[0].runs[0].font.size = Pt(9)
    for i, txt in enumerate(['日期：%s' % info['date'], '日期：——————', '日期：——————']):
        stb.rows[1].cells[i].text = txt
        stb.rows[1].cells[i].paragraphs[0].runs[0].font.size = Pt(9)
    doc.save(fpath)


# ============ 甲方交付文件中心 ============
DELIVER_TYPES = [
    ('construction_plan', '施工进度计划', ['pdf', 'docx', 'xlsx']),
    ('progress', '施工进度', ['pdf', 'xlsx']),
    ('quantity', '工程量清单', ['pdf', 'xlsx', 'csv']),
    ('org', '人员组织架构', ['pdf', 'docx', 'xlsx']),
    ('daily_report', '施工日报', ['pdf', 'xlsx']),
    ('test_record', '测试记录', ['pdf', 'xlsx']),
    ('acceptance', '验收资料', ['pdf', 'docx']),
    ('issue_fix', '问题整改记录', ['pdf', 'xlsx']),
    ('audit_record', '质量安全检查记录', ['pdf', 'xlsx']),
]
TABLE_OF_DELIVER = {
    'progress': 'sys_status',
    'quantity': 'details',
    'org': 'org',
    'daily_report': 'daily_report',
    'test_record': 'test_record',
    'acceptance': 'acceptance',
    'issue_fix': 'issues',
    'audit_record': 'audit',
}
MANIFEST_HEADERS = ['文档编号', '文档名称', '类型', '版本', '生成时间', '编制人', '审核状态', '修改说明', '文件路径']


def _manifest_path(proj):
    return os.path.join(_paths(proj)['deliver'], '交付文件登记表.xlsx')


def ensure_deliver(proj):
    d = _paths(proj)['deliver']
    os.makedirs(d, exist_ok=True)
    p = _manifest_path(proj)
    if not os.path.exists(p):
        wb = openpyxl.Workbook()
        ws = wb.active; ws.title = '交付文件'
        rebuild_sheet(ws, MANIFEST_HEADERS, [], [14, 24, 12, 8, 16, 10, 9, 24, 40], center_cols={1, 3, 4, 5, 6, 7})
        wb.save(p)
    return p


def load_manifest(proj):
    p = ensure_deliver(proj)
    wb = load_workbook(p)
    ws = wb['交付文件']
    rows = []
    for r in range(2, ws.max_row + 1):
        row = [ws.cell(row=r, column=c).value for c in range(1, len(MANIFEST_HEADERS) + 1)]
        if all(v is None or str(v).strip() == '' for v in row):
            continue
        rows.append(row)
    return rows


def gen_deliver(proj, doc_type, fmt, editor='', note=''):
    """生成甲方交付文件并登记版本。返回登记行与文件路径"""
    definition = next((t for t in DELIVER_TYPES if doc_type in (t[0],t[1])), None)
    if definition is None: raise ValueError('未知交付文件类型: '+str(doc_type))
    doc_type,tname,formats = definition
    if fmt not in formats: raise ValueError('该交付文件类型不支持格式: '+str(fmt))
    deliver_dir = _paths(proj)['deliver']
    os.makedirs(deliver_dir, exist_ok=True)
    # 文档名称
    tname = next((t[1] for t in DELIVER_TYPES if t[0] == doc_type), doc_type)
    # All projects use the same live schedule and source-table mappings.
    if doc_type == 'construction_plan':
        table = 'construction_plan'
    else:
        table = dict(TABLE_OF_DELIVER).get(doc_type, 'details')
    # 版本计算：同类型同格式已存在则 minor +0.1（V1.0→V1.1→V2.0）；不同格式视为独立文件
    manifest = load_manifest(proj)
    same = [m for m in manifest if str(m[2]) == tname and str(m[8]).endswith('.' + fmt)]
    ver = 'V1.0'
    if same:
        vs = [1.0]
        for m in same:
            v = str(m[3] or 'V1.0')
            try:
                a, b = v.replace('V', '').split('.')
                vs.append(int(a) + int(b) / 10.0)
            except Exception:
                pass
        ver = 'V%.1f' % (max(vs) + 0.1)
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    fname = '%s_%s_%s.%s' % (tname, ver, datetime.date.today().strftime('%Y%m%d'), fmt)
    fpath = os.path.join(deliver_dir, fname)
    title, hdr, rows = collect_table(proj, table)
    info = _title_info(proj, tname)
    info['editor'] = editor; info['version'] = ver
    if fmt == 'xlsx':
        _to_xlsx(fpath, tname, hdr, rows, info)
    elif fmt == 'pdf':
        _to_pdf(fpath, tname, hdr, rows, info)
    elif fmt == 'docx':
        _to_docx(fpath, tname, hdr, rows, info)
    elif fmt == 'csv':
        _to_csv(fpath, hdr, rows)
    else:
        raise ValueError('交付文件格式不支持: ' + fmt)
    # 登记
    no = 'DLV-%s-%02d' % (proj.upper(), len(manifest) + 1)
    rec = [no, tname, tname, ver, now, editor or '—', '待审核', note or '', fpath]
    manifest.insert(0, rec)
    p = _manifest_path(proj)
    wb = load_workbook(p)
    ws = wb['交付文件']
    rebuild_sheet(ws, MANIFEST_HEADERS, manifest, [14, 24, 12, 8, 16, 10, 9, 24, 40], center_cols={1, 3, 4, 5, 6, 7})
    wb.save(p)
    return {'ok': True, 'record': rec, 'path': fpath, 'version': ver}


def update_manifest(proj, no, fields):
    """更新交付登记（如审核状态）"""
    manifest = load_manifest(proj)
    for m in manifest:
        if str(m[0]) == no:
            for k, v in fields.items():
                idx = {'审核状态': 6, '修改说明': 7}.get(k)
                if idx is not None:
                    m[idx] = v
            break
    p = _manifest_path(proj)
    wb = load_workbook(p)
    ws = wb['交付文件']
    rebuild_sheet(ws, MANIFEST_HEADERS, manifest, [14, 24, 12, 8, 16, 10, 9, 24, 40], center_cols={1, 3, 4, 5, 6, 7})
    wb.save(p)
    return {'ok': True}
