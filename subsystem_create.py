"""Create a subsystem and its linked ledgers as one recoverable operation."""
import copy
import datetime
import json
import math
import os
import shutil
import tempfile


def create_subsystem(server, data):
    proj = data.get('proj')
    if proj not in server.SUBS_OF:
        raise ValueError('项目无效')
    server.refresh_subsystems()
    name = str(data.get('name') or '').strip()
    unit = str(data.get('unit') or '').strip()
    owner = str(data.get('owner') or '').strip()
    if not name or len(name) > 60 or any(ord(c) < 32 for c in name) or name[0] in '=+-@':
        raise ValueError('子系统名称需为1至60字，不能含控制字符或以公式符号开头')
    if not unit or len(unit) > 20:
        raise ValueError('请填写计量单位（最多20字）')
    cfg = server.sub_settings(proj)
    names = set(server.SUBS_OF[proj]) | {v.get('name', k) for k, v in cfg.items()}
    if name.casefold() in {n.casefold() for n in names} or server.canonical_sub(proj, name) in names:
        raise ValueError('子系统名称已存在，请使用其他名称')
    raw = data.get('total')
    total = None if raw in (None, '') else float(raw)
    if total is not None and (not math.isfinite(total) or total < 0 or (proj == 'titan' and not total.is_integer())):
        raise ValueError('计划总量必须为非负数字，泰坦项目需为整数')
    start = datetime.date.fromisoformat(str(data.get('start') or ''))
    finish = datetime.date.fromisoformat(str(data.get('finish') or ''))
    count = (finish-start).days+1
    if not 1 <= count <= 731:
        raise ValueError('计划工期需为1至731个自然日')
    dates = [(start+datetime.timedelta(days=i)).isoformat() for i in range(count)]
    daily = dict(zip(dates, server.allocate_plan(total, [1]*count)))
    F = server._paths(proj)
    root = os.path.dirname(F['p30'])
    settings_path = os.path.join(root, '子系统计划设置.json')
    targets = [F[k] for k in ('p30', 'p31', 'p32', 'p34')] + [settings_path]
    backup = os.path.join(root, '新增子系统备份', datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    os.makedirs(backup, exist_ok=True)
    existed = {p: os.path.exists(p) for p in targets}
    for p in targets:
        if existed[p]:
            shutil.copy2(p, os.path.join(backup, os.path.basename(p)))
    books = {k: server.load_workbook(F[k]) for k in ('p30', 'p31', 'p32')}

    def append(ws, values):
        row = ws.max_row + 1
        for col, value in enumerate(values, 1):
            cell = ws.cell(row, col, value)
            if isinstance(value, str):
                cell.data_type = 's'
            if row > 2:
                cell._style = copy.copy(ws.cell(row-1, col)._style)
            cell.alignment = server.CENTER
        if ws.auto_filter.ref:
            ws.auto_filter.ref = ws.dimensions

    item = '主要工程量'
    ws = books['p30']['工程量总览']
    ids = {str(ws.cell(r, 1).value) for r in range(2, ws.max_row+1)}
    number = 1
    while f'QTY-C{number:03d}' in ids:
        number += 1
    qty = [f'QTY-C{number:03d}', name, item, unit, total] if proj == 'titan' else [f'QTY-C{number:03d}', name, item, total, unit]
    append(ws, qty + [0, 0, None, None, '', '', owner, '未开始'])
    r = ws.max_row
    if proj == 'titan':
        ws.cell(r, 8, f'=IF(ISNUMBER(E{r}),E{r}-G{r},"")')
        ws.cell(r, 9, f'=IF(E{r}>0,G{r}/E{r},"")').number_format = '0.00%'
    else:
        ws.cell(r, 8, f'=IF(D{r}>0,G{r}/D{r},"")').number_format = '0.00%'
        ws.cell(r, 9, '未开始')
    append(books['p31']['子系统状态'], [name, '未开始', start.isoformat(), finish.isoformat(), 0, 0, total, unit, None, None, None, owner, datetime.datetime.now().isoformat(timespec='seconds'), '', item, '手工新增'])
    ws31 = books['p31']['子系统状态']; r = ws31.max_row
    ws31.cell(r, 9, f'=IF(ISNUMBER(G{r}),G{r}-F{r},"")')
    ws31.cell(r, 10, f'=IF(G{r}>0,F{r}/G{r},"")').number_format = '0.00%'
    ws31.cell(r, 11, '未开始')
    append(books['p32']['详细工程量'], [name, item, unit, total, 0, 0, total, 0 if total else None, '手工新增', '', start.isoformat(), finish.isoformat(), '[]', 0])
    ws32 = books['p32']['详细工程量']; r = ws32.max_row
    ws32.cell(r, 7, f'=IF(ISNUMBER(D{r}),D{r}-F{r},"")')
    ws32.cell(r, 8, f'=IF(D{r}>0,F{r}/D{r},"")').number_format = '0.00%'
    append(books['p32']['施工步骤'], [name, 1, '施工实施', start.isoformat(), finish.isoformat(), '未开始', '', owner, ''])
    for kind in ('内部验收', '试运行', '终验收'):
        append(books['p32']['验收记录'], [name, kind, '未开始', '', '', '', '', ''])
    cfg[name] = {'custom': True, 'name': name, 'start': start.isoformat(), 'finish': finish.isoformat(), 'start_no': 1, 'offset_start': 1, 'daily': daily}
    if proj == 'bg':
        daily_sheet = books['p30']['子系统每日计划'] if '子系统每日计划' in books['p30'].sheetnames else books['p30'].create_sheet('子系统每日计划')
        if daily_sheet.cell(1, 1).value is None:
            for col, label in enumerate(['日期', '子系统', '计划量'], 1):
                daily_sheet.cell(1, col, label)
        for day, value in daily.items():
            append(daily_sheet, [day, name, value])
    try:
        with tempfile.TemporaryDirectory(dir=root, prefix='新增子系统暂存_') as staging:
            for key, book in books.items():
                book.save(os.path.join(staging, os.path.basename(F[key])))
                book.close()
            with open(os.path.join(staging, os.path.basename(settings_path)), 'w', encoding='utf-8') as stream:
                json.dump(cfg, stream, ensure_ascii=False, indent=2)
            for target in [F[k] for k in books] + [settings_path]:
                os.replace(os.path.join(staging, os.path.basename(target)), target)
            server.refresh_subsystems()
            if proj == 'titan':
                server.update_progress_check(proj)
            else:
                server.update_progress_check_bg(proj)
    except Exception:
        for book in books.values():
            book.close()
        for target in targets:
            if existed[target]:
                shutil.copy2(os.path.join(backup, os.path.basename(target)), target)
            elif os.path.exists(target):
                os.remove(target)
        server.refresh_subsystems()
        raise
    return {'ok': True, 'sub': name, 'msg': '子系统已添加，工程量、状态、计划、明细、验收及全局统计已同步'}


def delete_subsystem(server, data):
    proj=str(data.get('proj') or '');sub=str(data.get('sub') or '')
    if proj not in server.SUBS_OF or sub not in server.SUBS_OF[proj]:raise ValueError('子系统不存在')
    cfg=server.sub_settings(proj)
    if proj in ('bg','titan') and not (cfg.get(sub) or {}).get('custom'):
        raise ValueError('系统预置子系统不能删除，只能删除手工添加的子系统')
    F=server._paths(proj);root=os.path.dirname(F['p30']);settings_path=os.path.join(root,'子系统计划设置.json')
    targets=[F[k] for k in ('p30','p31','p32','p34','p05')]+[settings_path]
    backup=os.path.join(root,'删除子系统备份',datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f'));os.makedirs(backup,exist_ok=True)
    existed={p:os.path.exists(p) for p in targets}
    for p in targets:
        if existed[p]:shutil.copy2(p,os.path.join(backup,os.path.basename(p)))
    try:
        def remove_rows(ws,col):
            for r in range(ws.max_row,1,-1):
                if str(ws.cell(r,col).value or '')==sub:ws.delete_rows(r)
        w=server.load_workbook(F['p30']);remove_rows(w['工程量总览'],2);remove_rows(w['每日计量'],2)
        if '子系统每日计划' in w.sheetnames:remove_rows(w['子系统每日计划'],2)
        w.save(F['p30']);w.close()
        w=server.load_workbook(F['p31']);remove_rows(w['子系统状态'],1);w.save(F['p31']);w.close()
        w=server.load_workbook(F['p32'])
        for name in ('施工步骤','详细工程量','验收记录'):
            if name in w.sheetnames:remove_rows(w[name],1)
        w.save(F['p32']);w.close()
        if os.path.exists(F['p05']):
            w=server.load_workbook(F['p05'])
            for ws in w.worksheets:
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        if str(cell.value or '')==sub:cell.value='项目级'
            w.save(F['p05']);w.close()
        cfg.pop(sub,None)
        with open(settings_path+'.tmp','w',encoding='utf-8') as f:json.dump(cfg,f,ensure_ascii=False,indent=2)
        os.replace(settings_path+'.tmp',settings_path)
        server.SUBS_OF[proj].remove(sub)
        server.pm_core.remove_subsystem_task(server.DIR,proj,server.PROJ_NAMES[proj],server.SUBS_OF[proj],sub,'删除子系统')
        if proj=='titan':server.update_progress_check(proj)
        else:server.update_progress_check_bg(proj)
        return {'ok':True,'msg':'子系统已删除，相关工程量、计划、步骤和验收记录已移出统计'}
    except Exception:
        for p in targets:
            bak=os.path.join(backup,os.path.basename(p))
            if existed[p] and os.path.exists(bak):shutil.copy2(bak,p)
        server.refresh_subsystems();raise
