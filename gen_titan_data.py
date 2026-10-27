# -*- coding: utf-8 -*-
"""生成泰坦IDC项目 5 个独立数据 Excel（29/30/31/32/05），存于泰坦IDC目录，与BG项目完全独立"""
import os
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import titan_defs as T

ROOT = r'D:\项目资料\安伟达\国轩\BG项目管理系统集成文件体系'
OUT = os.path.join(ROOT, '泰坦IDC')
os.makedirs(OUT, exist_ok=True)

HDR_FILL = PatternFill('solid', fgColor='1F3A5F')
WARN_FILL = PatternFill('solid', fgColor='FFF3D9')
GREEN_FILL = PatternFill('solid', fgColor='E2F0E5')
BODY_FONT = Font(name='微软雅黑', size=9)
HDR_FONT = Font(name='微软雅黑', size=9, bold=True, color='FFFFFF')
BOLD = Font(name='微软雅黑', size=9, bold=True)
BORDER = Border(*[Side(style='thin', color='B8C4D0')] * 4)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
WRAP = Alignment(vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

def put(ws, r, c, v, font=None, fill=None, align=None, border=True):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font or BODY_FONT
    if fill: cell.fill = fill
    cell.alignment = align or (CENTER if isinstance(v, (int, float)) else WRAP)
    if border: cell.border = BORDER
    return cell

def header(ws, headers, widths=None):
    for c, h in enumerate(headers, 1):
        put(ws, 1, c, h, font=HDR_FONT, fill=HDR_FILL)
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def day_list(n):
    return [(T.T0 + timedelta(days=i)).strftime('%Y-%m-%d') for i in range(1, n + 1)]

# ============ 29 项目人员管理表 ============
wb = Workbook()
ws = wb.active; ws.title = '每日到场汇总'
header(ws, ['日期', '星期', '总人数', '中方人数', '当地人数', '施工人数', '安装人数', '调试人数', '验收人数', '管理人员', '工作内容', '状态'],
       [12, 6, 8, 8, 8, 8, 8, 8, 8, 8, 46, 12])
days = day_list(15)
for i, d in enumerate(days):
    r = i + 2
    wd = ['一', '二', '三', '四', '五', '六', '日'][(T.T0 + timedelta(days=i + 1)).weekday()]
    put(ws, r, 1, d); put(ws, r, 2, '周' + wd)
    put(ws, r, 3, 0); put(ws, r, 4, 0); put(ws, r, 5, 0)
    put(ws, r, 6, 0); put(ws, r, 7, 0); put(ws, r, 8, 0); put(ws, r, 9, 0); put(ws, r, 10, 0)
    put(ws, r, 11, '（T+' + str(i + 1) + '）计划人数' + str(T.PLAN['people'][i]) + '人·' + T.DAILY_TASKS[i], align=LEFT)
    put(ws, r, 12, '未开始')
r = len(days) + 2
put(ws, r, 1, '合计', font=BOLD); put(ws, r, 3, f'=SUM(C2:C{r-1})', font=BOLD)
put(ws, r, 11, '15天人员投入峰值46人/日（第4~11天高峰期）', align=LEFT, font=BOLD)
for c in range(1, 13): ws.cell(row=r, column=c).border = BORDER
dv = DataValidation(type='list', formula1='"未开始,待施工,施工中,测试中,验收中,已完成,暂停"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f'L2:L{r-1}')

ws = wb.create_sheet('人员明细')
header(ws, ['日期', '姓名', '职位', '所属班组', '工作类型', '工作区域', '工作内容', '工时(h)', '备注'], [12, 12, 14, 16, 12, 16, 34, 9, 24])

ws = wb.create_sheet('班组人力计划')
header(ws, ['班组', '人数', '进场时段', '主要职责', '关键要求'], [22, 8, 14, 52, 46])
for i, b in enumerate(T.BANZU):
    r = i + 2
    for c, v in enumerate(b, 1): put(ws, r, c, v, align=LEFT if c in (4, 5) else CENTER)
wb.save(os.path.join(OUT, '29_项目人员管理表.xlsx'))
print('29 OK')

# ============ 30 项目进度工程量跟踪表 ============
wb = Workbook()
ws = wb.active; ws.title = '工程量总览'
header(ws, ['序号', '子系统', '项目', '单位', '计划总量', '今日完成', '累计完成', '剩余量', '完成率%', '备注', '计划完成率%', '负责人', '状态', '数据来源', '计量口径', '说明'],
       [6, 20, 18, 8, 10, 10, 10, 10, 10, 30, 12, 16, 10, 24, 20, 26])
subs = T.SUBSYSTEMS
for i, s in enumerate(subs):
    r = i + 2
    detail_total = [d for d in s['details'] if d[2]]
    tot = detail_total[0][2] if detail_total else None
    unit = detail_total[0][1] if detail_total else ''
    src = detail_total[0][3] if detail_total else ''
    put(ws, r, 1, i + 1); put(ws, r, 2, s['name']); put(ws, r, 3, detail_total[0][0] if detail_total else s['name'])
    put(ws, r, 4, unit); put(ws, r, 5, tot); put(ws, r, 6, 0); put(ws, r, 7, 0)
    put(ws, r, 8, f'=IF(ISNUMBER(E{r}),E{r}-G{r},"")')
    put(ws, r, 9, f'=IF(ISNUMBER(E{r}),IF(E{r}=0,"",ROUND(G{r}/E{r}*100,1)),"")')
    put(ws, r, 10, s['details'][0][3] if s['details'] else '')
    put(ws, r, 11, '')
    put(ws, r, 12, s['owner']); put(ws, r, 13, '未开始')
    put(ws, r, 14, src); put(ws, r, 15, '15天累计计划=总工程量（与施工计划表14-1一致）')
    put(ws, r, 16, f"{s['name']}：{s['plan_start']}~{s['plan_end']}，占15天工期")
ws.freeze_panes = 'A2'

ws = wb.create_sheet('每日计划')
header(ws, ['施工日', '日期', '仓库分拣计划', '分拣累计', '7F布线计划', '7F累计', '9F布线计划', '9F累计', 'ODF安装计划', 'ODF累计', '端接计划', '端接累计', '测试计划', '测试累计', '当日总计划', '累计完成率%', '预计人员', '当天主要任务', '进度状态'],
       [10, 12, 10, 9, 10, 9, 10, 9, 10, 9, 9, 9, 9, 9, 10, 10, 8, 46, 10])
P = T.PLAN
for i in range(15):
    r = i + 2
    put(ws, r, 1, f'第{i+1}天' + ('·M' if i in (4, 9, 12, 14) else ''))
    put(ws, r, 2, days[i])
    put(ws, r, 3, P['fen'][i]); put(ws, r, 4, f'=D{r-1}+C{r}' if i else f'=C{r}')
    put(ws, r, 5, P['b7'][i]);  put(ws, r, 6, f'=F{r-1}+E{r}' if i else f'=E{r}')
    put(ws, r, 7, P['b9'][i]);  put(ws, r, 8, f'=H{r-1}+G{r}' if i else f'=G{r}')
    put(ws, r, 9, P['odf'][i]); put(ws, r, 10, f'=J{r-1}+I{r}' if i else f'=I{r}')
    put(ws, r, 11, P['duan'][i]); put(ws, r, 12, f'=L{r-1}+K{r}' if i else f'=K{r}')
    put(ws, r, 13, P['ces'][i]); put(ws, r, 14, f'=N{r-1}+M{r}' if i else f'=M{r}')
    put(ws, r, 15, f'=C{r}+E{r}+G{r}+I{r}+K{r}+M{r}')
    put(ws, r, 16, f'=ROUND((D{r}+F{r}+H{r}+J{r}+L{r}+N{r})/7216*100,2)')
    put(ws, r, 17, P['people'][i])
    put(ws, r, 18, T.DAILY_TASKS[i], align=LEFT)
    put(ws, r, 19, '绿·按计划')
r = 17
put(ws, r, 1, '合计', font=BOLD); put(ws, r, 3, f'=SUM(C2:C{r-1})', font=BOLD); put(ws, r, 5, f'=SUM(E2:E{r-1})', font=BOLD)
put(ws, r, 7, f'=SUM(G2:G{r-1})', font=BOLD); put(ws, r, 9, f'=SUM(I2:I{r-1})', font=BOLD)
put(ws, r, 11, f'=SUM(K2:K{r-1})', font=BOLD); put(ws, r, 13, f'=SUM(M2:M{r-1})', font=BOLD)
put(ws, r, 15, f'=SUM(O2:O{r-1})', font=BOLD); put(ws, r, 16, '100', font=BOLD)
put(ws, r, 17, '峰值46', font=BOLD); put(ws, r, 18, '15天累计计划=总工程量7216，100%一致', align=LEFT, font=BOLD)
ws.freeze_panes = 'A2'

ws = wb.create_sheet('每日计量')
header(ws, ['日期', '子系统', '项目', '今日完成', '单位', '累计完成', '来源', '备注'], [12, 18, 16, 10, 8, 10, 12, 30])

ws = wb.create_sheet('判定规则')
header(ws, ['进度判定', '触发条件', '响应措施'], [14, 52, 60])
rules = [
    ('正常·绿色', '实际进度≥计划进度（偏差≥0）', '按计划推进，维持班组配置'),
    ('轻微滞后·黄色', '实际进度较计划滞后≤1天 或 累计完成率低于计划≤5个百分点', '增加人员或调整班组，利用非关键工序支援关键工序，次日追回进度'),
    ('明显滞后·红色', '实际进度较计划滞后＞1天 或 累计完成率低于计划＞5个百分点', '制定赶工措施：延长作业时段（经甲方批准）、增派班组与仪器、调整流水节拍，并向甲方书面报备'),
]
for i, row in enumerate(rules):
    for c, v in enumerate(row, 1): put(ws, i + 2, c, v, align=LEFT if c > 1 else CENTER)
set_widths(ws, [14, 52, 60])

ws = wb.create_sheet('量纲表')
header(ws, ['子系统', '计量单位', '说明'], [24, 10, 60])
for i, s in enumerate(subs):
    d = s['details'][0]
    put(ws, i + 2, 1, s['name']); put(ws, i + 2, 2, d[1]); put(ws, i + 2, 3, d[3], align=LEFT)
set_widths(ws, [24, 10, 60])
wb.save(os.path.join(OUT, '30_项目进度工程量跟踪表.xlsx'))
print('30 OK')

# ============ 31 子系统状态总览表 ============
wb = Workbook()
ws = wb.active; ws.title = '子系统状态'
header(ws, ['子系统', '当前状态', '计划开始', '计划结束', '今日完成', '累计完成', '计划总量', '单位', '剩余量', '完成率%', '进度判定', '负责人', '更新时间', '今日工作内容', '代表计量项', '数据来源'],
       [20, 10, 10, 10, 10, 10, 10, 8, 10, 10, 14, 16, 12, 34, 14, 22])
for i, s in enumerate(subs):
    r = i + 2
    det = [d for d in s['details'] if d[2]]
    tot = det[0][2] if det else None; unit = det[0][1] if det else ''
    put(ws, r, 1, s['name']); put(ws, r, 2, '未开始')
    put(ws, r, 3, s['plan_start']); put(ws, r, 4, s['plan_end'])
    put(ws, r, 5, 0); put(ws, r, 6, 0); put(ws, r, 7, tot); put(ws, r, 8, unit)
    put(ws, r, 9, f'=IF(ISNUMBER(G{r}),G{r}-F{r},"")')
    put(ws, r, 10, f'=IF(ISNUMBER(G{r}),IF(G{r}=0,"",ROUND(F{r}/G{r}*100,1)),"")')
    put(ws, r, 11, f'=IF(ISNUMBER(G{r}),IF(G{r}=0,"未开始",IF(F{r}>=G{r},"完成",IF(F{r}/G{r}>=0.9,"正常","滞后"))),"")')
    put(ws, r, 12, s['owner']); put(ws, r, 13, '')
    put(ws, r, 14, f"（T+{i}）待填报", align=LEFT)
    put(ws, r, 15, det[0][0] if det else '')
    put(ws, r, 16, '施工计划表14-1 + 32号明细联动')
dv = DataValidation(type='list', formula1='"' + ','.join(T.STATUS_LIST) + '"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f'B2:B{len(subs)+1}')
ws.freeze_panes = 'A2'
wb.save(os.path.join(OUT, '31_子系统状态总览表.xlsx'))
print('31 OK')

# ============ 32 子系统施工步骤与明细表 ============
wb = Workbook()
ws = wb.active; ws.title = '施工步骤'
header(ws, ['子系统', '序号', '施工步骤', '计划开始', '计划完成', '状态', '完成日期', '责任人', '备注'], [20, 6, 20, 10, 10, 10, 12, 16, 30])
r = 2
for s in subs:
    for j, st in enumerate(s['steps'], 1):
        put(ws, r, 1, s['name']); put(ws, r, 2, j); put(ws, r, 3, st[0], align=LEFT)
        put(ws, r, 4, st[1]); put(ws, r, 5, st[2]); put(ws, r, 6, '未开始'); put(ws, r, 7, '')
        put(ws, r, 8, s['owner']); put(ws, r, 9, '')
        r += 1
dv = DataValidation(type='list', formula1='"未开始,待施工,施工中,测试中,验收中,已完成,暂停"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f'F2:F{r-1}')
ws.freeze_panes = 'A2'

ws = wb.create_sheet('详细工程量')
header(ws, ['子系统', '项目', '单位', '计划总量', '今日完成', '累计完成', '剩余量', '完成率%', '数据来源', '备注'], [20, 18, 8, 10, 10, 10, 10, 10, 24, 40])
r = 2
for s in subs:
    for d in s['details']:
        put(ws, r, 1, s['name']); put(ws, r, 2, d[0], align=LEFT); put(ws, r, 3, d[1])
        put(ws, r, 4, d[2]); put(ws, r, 5, 0); put(ws, r, 6, 0)
        put(ws, r, 7, f'=IF(ISNUMBER(D{r}),D{r}-F{r},"")')
        put(ws, r, 8, f'=IF(ISNUMBER(D{r}),IF(D{r}=0,"",ROUND(F{r}/D{r}*100,1)),"")')
        put(ws, r, 9, d[3]); put(ws, r, 10, '工具自动联动31号今日/累计')
        r += 1
dv = DataValidation(type='list', formula1='"未开始,待施工,施工中,测试中,验收中,已完成,暂停"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f'F2:F{r-1}')
ws.freeze_panes = 'A2'

ws = wb.create_sheet('验收记录')
header(ws, ['子系统', '验收类型', '状态', '验收日期', '验收人', '验收结论', '未通过原因', '整改日期'], [20, 16, 10, 12, 12, 24, 30, 12])
r = 2
for s in subs:
    for a in T.ACCEPTANCE:
        put(ws, r, 1, s['name']); put(ws, r, 2, a[0]); put(ws, r, 3, '未开始'); put(ws, r, 4, a[1])
        put(ws, r, 5, ''); put(ws, r, 6, '', align=LEFT); put(ws, r, 7, '', align=LEFT); put(ws, r, 8, '')
        r += 1
dv = DataValidation(type='list', formula1='"未开始,通过,未通过,进行中"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f'C2:C{r-1}')
ws.freeze_panes = 'A2'
wb.save(os.path.join(OUT, '32_子系统施工步骤与明细表.xlsx'))
print('32 OK 步骤行', sum(len(s['steps']) for s in subs), '明细行', sum(len(s['details']) for s in subs), '验收行', len(subs) * len(T.ACCEPTANCE))

# ============ 05 项目问题及待办事项 ============
wb = Workbook()
ws = wb.active; ws.title = '问题清单'
header(ws, ['编号', '问题描述', '问题类型', '优先级', '提出人', '提出日期', '具体问题', '影响范围', '影响程度', '责任部门', '解决方案', '负责人', '计划解决日期', '实际解决日期', '状态', '是否关闭', '涉及合同', '备注'],
       [10, 40, 12, 10, 12, 12, 40, 18, 12, 14, 40, 12, 12, 12, 10, 10, 12, 16])
for i, p in enumerate(T.PENDING, 1):
    r = i + 1
    put(ws, r, 1, f'TISSUE-{i:03d}')
    put(ws, r, 2, p, align=LEFT); put(ws, r, 3, '甲方待确认')
    put(ws, r, 4, '高' if i in (5, 6) else '中'); put(ws, r, 5, '泰坦项目部')
    put(ws, r, 6, '2026-09-08'); put(ws, r, 7, p, align=LEFT)
    put(ws, r, 8, '工程量/工期口径'); put(ws, r, 9, '影响计划执行')
    put(ws, r, 10, '甲方/设计')
    put(ws, r, 11, '按确认口径执行，未确认前按施工计划现口径执行', align=LEFT)
    put(ws, r, 12, '项目经理'); put(ws, r, 13, 'T+2前')
    put(ws, r, 14, ''); put(ws, r, 15, '未关闭'); put(ws, r, 16, '否')
    put(ws, r, 17, '施工技术协议'); put(ws, r, 18, '《72芯集束光缆15天施工计划》附录A')
ws.freeze_panes = 'A2'
wb.save(os.path.join(OUT, '05_项目问题及待办事项.xlsx'))
print('05 OK')

print('ALL TITAN FILES GENERATED ->', OUT)
