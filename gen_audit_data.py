# -*- coding: utf-8 -*-
"""生成工程审计登记表（BG/泰坦）与泰坦施工进度检查表初始文件"""
import os, sys, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import titan_defs

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 体系根
HDR_FILL = PatternFill('solid', fgColor='1F3A5F')
BODY_FONT = Font(name='微软雅黑', size=9)
HDR_FONT = Font(name='微软雅黑', size=9, bold=True, color='FFFFFF')
BORDER = Border(*[Side(style='thin', color='B8C4D0')] * 4)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
WRAP = Alignment(vertical='center', wrap_text=True)

AUDIT_HEADERS = ['日期', '检查类型', '问题编号', '问题描述', '图片文件', '是否已处理', '处理时间', '处理说明', '登记人', '登记时间']
AUDIT_W = [12, 12, 12, 40, 30, 11, 12, 30, 10, 18]


def rebuild(ws, headers, rows, widths, center_cols=None):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.border = BORDER; cell.alignment = CENTER
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = BODY_FONT; cell.border = BORDER
            cell.alignment = CENTER if (center_cols and c in center_cols) else WRAP
    for r in range(len(rows) + 2, ws.max_row + 1):
        ws.delete_rows(len(rows) + 2)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def gen_audit(proj):
    if proj == 'titan':
        d = os.path.join(BASE, '泰坦IDC', '工程审计')
    else:
        d = os.path.join(BASE, '工程审计')
    os.makedirs(d, exist_ok=True)
    p = os.path.join(d, '33_工程审计登记表.xlsx')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '审计登记'
    rebuild(ws, AUDIT_HEADERS, [], AUDIT_W, center_cols={1, 2, 3, 6, 7, 9, 10})
    wb.save(p)
    print('审计登记表 OK ->', p)


def gen_progress_check():
    """泰坦 34 号施工进度检查表（子项级：施工日×子系统 90 行；项目未开工时全部「未开工」）"""
    plan = titan_defs.PLAN
    keys = ['fen', 'b7', 'b9', 'odf', 'duan', 'ces']
    subs = [s['name'] for s in titan_defs.SUBSYSTEMS]
    total_plan = 7216
    days = [(titan_defs.T0 + datetime.timedelta(days=i)).strftime('%Y-%m-%d') for i in range(1, 16)]
    rows = []
    for i in range(15):
        d = days[i]
        for j, sub in enumerate(subs):
            key = keys[j]
            p_daily = plan[key][i]
            cum_p = sum(plan[key][:i + 1])
            rows.append([i + 1, d, sub, p_daily, 0, -p_daily, cum_p, 0, -cum_p, '—', '未开工'])
    over_rows = [
        ['总工程量', total_plan, '单位'],
        ['累计实际完成', 0, '单位'],
        ['总体完成率', '0.0%', ''],
        ['当前应达（第—施工日）', '0.0%', '按工期累计计划'],
        ['偏差', '0.0%', '负=落后'],
        ['判定', '未开工', '超前>=100% / 达标>=98% / 预警>=90% / 落后<90%'],
        ['检查时间', datetime.datetime.now().strftime('%Y-%m-%d %H:%M'), ''],
    ]
    d = os.path.join(BASE, '泰坦IDC')
    os.makedirs(d, exist_ok=True)
    p = os.path.join(d, '34_施工进度检查表.xlsx')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '每日进度检查'
    rebuild(ws, ['施工日', '日期', '子系统', '计划当日量', '实际当日量', '当日差额', '累计计划量', '累计实际量', '累计差额', '完成率%', '判定状态'],
            rows, [7, 11, 16, 10, 10, 10, 10, 10, 10, 8, 9], center_cols=set(range(1, 12)))
    ws2 = wb.create_sheet('总体进度检查')
    rebuild(ws2, ['检查项', '数值', '说明'], over_rows, [24, 20, 46], center_cols={1, 2})
    wb.save(p)
    print('施工进度检查表 OK ->', p, '(%d 行)' % len(rows))


if __name__ == '__main__':
    gen_audit('bg')
    gen_audit('titan')
    gen_progress_check()
    print('ALL AUDIT FILES GENERATED')
