"""Subsystem scheduling and self-contained owner reports (calendar days)."""
import datetime as dt
import html
import json
import os
import re


def day(value):
    try:
        return dt.date.fromisoformat(str(value)[:10])
    except (ValueError, TypeError):
        return None


def load(root):
    path = os.path.join(root, '施工流程配置.json')
    if not os.path.exists(path):
        return {}
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def save(root, config):
    path = os.path.join(root, '施工流程配置.json')
    with open(path + '.tmp', 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    os.replace(path + '.tmp', path)


def calculate(snapshot, config=None):
    config = config or {}
    tasks = snapshot.get('tasks', [])
    profile = snapshot.get('profile', {})
    nodes = []
    warnings = []
    overrides = config.get('durations', {})
    for t in tasks:
        if t.get('source_type') != 'subsystem':
            continue
        key = t.get('source_key') or t['system_name']
        start, finish = day(t.get('plan_start')), day(t.get('plan_finish'))
        duration = (finish-start).days+1 if start and finish and finish >= start else None
        if key in overrides:
            value = overrides[key]
            if isinstance(value, bool) or not isinstance(value, int) or not 1 <= value <= 36500:
                raise ValueError('工期必须为 1～36500 的整数自然日')
            duration = value
        if duration is None:
            warnings.append(key + '：工作包缺少有效起止日期，请补充工期')
        qty = float(t.get('display_plan_qty', t.get('plan_qty', 0)) or 0)
        actual = float(t.get('display_actual_qty', t.get('actual_qty', 0)) or 0)
        nodes.append(dict(id=key, name=t.get('task_name') or key, code=t.get('wbs_code', ''),
                          owner=t.get('owner') or '待明确', duration=duration, source_start=str(start or ''),
                          source_finish=str(finish or ''), override=key in overrides,
                          progress=round(min(100, actual/qty*100), 1) if qty else None))
    by = {n['id']: n for n in nodes}
    if not nodes:
        warnings.append('尚无子系统工作包，请先创建子系统')
    # WBS dependencies crossing subsystem boundaries become subsystem FS constraints.
    code_map = {t.get('wbs_code'): t for t in tasks}
    def system(t):
        return t.get('source_key') if t.get('source_type') == 'subsystem' else t.get('system_name')
    edges = []
    if 'edges' in config:
        edges = config['edges']
    else:
        seen = set()
        for t in tasks:
            target = system(t)
            for code in re.split(r'[,，;；\s]+', str(t.get('predecessor') or '')):
                if not code:
                    continue
                parent = code_map.get(code)
                source = system(parent) if parent else None
                if source in by and target in by and source != target:
                    pair = (source, target)
                    if pair not in seen:
                        edges.append(dict(source=source, target=target, type='FS', lag=0))
                        seen.add(pair)
                elif not parent:
                    warnings.append('WBS ' + str(t.get('wbs_code')) + '：前置编码 ' + code + ' 不存在')
    if not isinstance(edges, list):
        raise ValueError('施工关系格式错误')
    incoming = {k: [] for k in by}
    outgoing = {k: [] for k in by}
    cleaned = []
    seen = set()
    for edge in edges:
        a, b, typ, lag = edge.get('source'), edge.get('target'), edge.get('type'), edge.get('lag', 0)
        if a not in by or b not in by:
            warnings.append('关系引用了已删除的子系统：' + str(a) + ' → ' + str(b) + '，请重新设置')
            continue
        if a == b:
            raise ValueError('子系统不能依赖自身：' + a)
        if typ not in ('FS', 'SS', 'FF') or isinstance(lag, bool) or not isinstance(lag, int) or abs(lag) > 36500:
            raise ValueError('关系须为 FS / SS / FF，间隔须为整数天（±36500 内）')
        identity = (a, b, typ, lag)
        if identity in seen:
            continue
        seen.add(identity)
        e = dict(source=a, target=b, type=typ, lag=lag)
        cleaned.append(e)
        incoming[b].append(e)
        outgoing[a].append(e)
    degree = {k: len(v) for k, v in incoming.items()}
    queue = [k for k in by if degree[k] == 0]
    order = []
    while queue:
        key = queue.pop(0)
        order.append(key)
        for e in outgoing[key]:
            degree[e['target']] -= 1
            if degree[e['target']] == 0:
                queue.append(e['target'])
    if len(order) != len(nodes):
        raise ValueError('施工关系存在循环：' + '、'.join(k for k in by if degree[k]))
    starts = [day(n['source_start']) for n in nodes if n['source_start']]
    base = day(profile.get('plan_start')) or (min(starts) if starts else None)
    if not base:
        warnings.append('项目开始日期未配置，图表按相对自然日显示')
    if not cleaned and len(nodes) > 1:
        warnings.append('尚未配置跨子系统施工依赖；当前按工作包日期识别先后与并行，不代表工艺依赖已确认')
    def weight(e):
        a, b = by[e['source']], by[e['target']]
        return e['lag'] + (a['duration'] if e['type'] == 'FS' else a['duration']-b['duration'] if e['type'] == 'FF' else 0)
    for key in order:
        n = by[key]
        n['level'] = max([by[e['source']]['level']+1 for e in incoming[key]] or [0])
        release = max(0, (day(n['source_start'])-base).days) if base and n['source_start'] else 0
        n['start'] = release if n['duration'] is not None else None
        for e in incoming[key]:
            parent = by[e['source']]
            if parent['start'] is None or n['start'] is None:
                n['start'] = None
            else:
                n['start'] = max(n['start'], parent['start'] + weight(e))
        n['end'] = n['start'] + n['duration'] if n['start'] is not None else None
        n['start_date'] = str(base+dt.timedelta(days=n['start'])) if base and n['start'] is not None else ''
        n['finish_date'] = str(base+dt.timedelta(days=n['end']-1)) if base and n['end'] is not None else ''
        if n['duration'] is not None and n['start'] is None:
            warnings.append(n['name'] + '：前置工期未明确，暂不能排程')
        if n['finish_date'] and n['source_finish'] and n['finish_date'] > n['source_finish']:
            warnings.append(n['name'] + '：依赖排程晚于工作包计划完成日 ' + n['source_finish'])
    known = [n for n in nodes if n['end'] is not None]
    total = max([n['end'] for n in known] or [0])
    complete = len(known) == len(nodes) and bool(nodes)
    for key in reversed(order):
        n = by[key]
        n['float'] = None
        if complete:
            latest = min([by[e['target']]['latest']-weight(e) for e in outgoing[key]] + [total-n['duration']])
            n['latest'] = latest
            n['float'] = latest-n['start']
    finish = str(base+dt.timedelta(days=total-1)) if base and total else ''
    if finish and profile.get('plan_finish') and finish > profile['plan_finish']:
        warnings.append('排程完成日 ' + finish + ' 晚于项目计划完成日 ' + profile['plan_finish'])
    parallel = []
    for i, a in enumerate(known):
        for b in known[i+1:]:
            overlap = min(a['end'], b['end'])-max(a['start'], b['start'])
            if overlap > 0:
                parallel.append(dict(source=a['id'], target=b['id'], days=overlap))
    return dict(nodes=sorted(nodes, key=lambda n: (n['start'] is None, n['start'] or 0, n['level'], n['code'])),
                edges=cleaned, parallel=parallel, total=total, complete=complete, start=str(base or ''), finish=finish,
                warnings=warnings, profile=profile, generated=dt.datetime.now().isoformat(timespec='seconds'),
                config=dict(durations={k:v for k,v in overrides.items() if k in by}, **({'edges':cleaned} if 'edges' in config else {})))


def report(model, title):
    esc = lambda x: html.escape(str(x), quote=True)
    nodes, edges = model['nodes'], model['edges']
    by = {n['id']: n for n in nodes}
    def label(n):
        return (n['start_date'] or ('第%d天' % (n['start']+1) if n['start'] is not None else '待排程')) + ' ～ ' + (n['finish_date'] or ('第%d天' % n['end'] if n['end'] is not None else '待排程'))
    def text(x, y, value, size=12, color='#334155'):
        return f'<text x="{x}" y="{y}" font-size="{size}" fill="{color}">{esc(value)}</text>'
    positions = {}
    levels = {}
    date_levels = {value:i for i,value in enumerate(sorted({n['start'] for n in nodes if n['start'] is not None}))}
    for n in nodes:
        level = n['level'] if edges else date_levels.get(n['start'], len(date_levels))
        levels.setdefault(level, []).append(n)
    width = max(1000, (max(levels, default=0)+1)*320)
    height = max(140, max([len(v) for v in levels.values()] or [1])*130+65)
    svg = [f'<svg xmlns="http://www.w3.org/2000/svg" role="img" aria-label="子系统施工关系图" viewBox="0 0 {width} {height}"><defs><marker id="arrow" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8" fill="#64748b"/></marker></defs>']
    for level, group in levels.items():
        svg.append(text(level*320+15, 20, ('前置层级 '+str(level+1)) if edges else ('同日起工 · '+(group[0]['start_date'] or ('第%d天'%(group[0]['start']+1) if group[0]['start'] is not None else '待排程'))),12))
        for i, n in enumerate(group):
            positions[n['id']] = (level*320+15, i*130+40)
    for e in edges:
        x, y = positions[e['source']]; xx, yy = positions[e['target']]
        svg.append(f'<path d="M{x+265},{y+48} C{x+290},{y+48} {xx-25},{yy+48} {xx},{yy+48}" fill="none" stroke="#64748b" marker-end="url(#arrow)"/>')
        svg.append(text(x+267, y+36, e['type']+('%+d天' % e['lag']), 10))
    for n in nodes:
        x, y = positions[n['id']]
        color = '#dc2626' if n['float'] == 0 else '#2563eb'
        svg.append(f'<rect x="{x}" y="{y}" width="265" height="100" rx="9" fill="#f8fafc" stroke="{color}"/>')
        name = n['name']
        svg.append(text(x+10, y+21, name[:19], 13))
        if len(name)>19: svg.append(text(x+10,y+38,name[19:38],11))
        svg.append(text(x+10, y+57, str(n['duration'] or '待定')+' 天 · '+('关键工作' if n['float']==0 else '子系统工作包')))
        svg.append(text(x+10, y+79, label(n), 11))
    svg.append('</svg>')
    flow = ''.join(svg)
    gh = 65+len(nodes)*45
    gantt = [f'<svg xmlns="http://www.w3.org/2000/svg" role="img" aria-label="施工工期甘特图" viewBox="0 0 1100 {gh}">']
    scale = 780/max(1, model['total'])
    for i in range(6):
        d = round(model['total']*i/5)
        x = 270+d*scale
        gantt.append(f'<path d="M{x},35 V{gh}" stroke="#e2e8f0"/>')
        gantt.append(text(x, 22, 'D+'+str(d), 11))
    for i, n in enumerate(nodes):
        y = 48+i*45
        gantt.append(text(5, y+16, n['name'][:22]))
        if n['start'] is None:
            gantt.append(text(275, y+16, '待补充工期或前置计划'))
            continue
        x, w = 270+n['start']*scale, max(2,n['duration']*scale)
        color = '#dc2626' if n['float']==0 else '#2563eb'
        gantt.append(f'<rect x="{x}" y="{y}" width="{w}" height="22" rx="3" fill="{color}" opacity=".22"/>')
        if n['progress'] is not None:
            gantt.append(f'<rect x="{x}" y="{y}" width="{w*n["progress"]/100}" height="22" fill="{color}"/>')
        gantt.append(text(x+4, y+16, str(n['duration'])+'天', 11))
    gantt.append('</svg>')
    rows = ''.join('<tr>'+''.join('<td>'+esc(v)+'</td>' for v in [n['code'],n['name'],n['owner'],n['duration'] or '待定',label(n),('%.1f%%'%n['progress']) if n['progress'] is not None else '无计量基数', '关键工作' if n['float']==0 else (str(n['float'])+'天' if n['float'] is not None else '待定'), '流程工期覆盖' if n['override'] else '工作包日期同步'])+'</tr>' for n in nodes)
    meanings = {'FS':'完成后开始','SS':'开始后开始','FF':'完成后完成'}
    relations = ''.join('<tr>'+''.join('<td>'+esc(v)+'</td>' for v in [by[e['source']]['name'],by[e['target']]['name'],e['type']+' '+meanings[e['type']],str(e['lag'])+'天'])+'</tr>' for e in edges)
    overlaps = ''.join('<tr><td>'+esc(by[e['source']]['name'])+'</td><td>'+esc(by[e['target']]['name'])+'</td><td>'+str(e['days'])+'天</td></tr>' for e in model['parallel'])
    warnings = ''.join('<li>'+esc(w)+'</li>' for w in model['warnings'])
    total_label = '总工期' if model['complete'] else '已排部分跨度（总工期待定）'
    style = 'body{font:14px "Microsoft YaHei",Arial,sans-serif;color:#172b4d;margin:24px;background:white}h1{font-size:25px}h2{font-size:18px;margin-top:28px}p{line-height:1.7}.muted{color:#64748b}.stats{display:flex;gap:15px;flex-wrap:wrap}.stat{background:#eff6ff;border-radius:8px;padding:16px;flex:1;min-width:150px}.stat b{font-size:23px;display:block;margin-top:6px}.chart{overflow:auto;border:1px solid #e2e8f0;padding:12px;border-radius:8px}.chart svg{width:100%;min-width:800px;height:auto}table{width:100%;border-collapse:collapse;font-size:12px}th,td{text-align:left;border:1px solid #dbe3ee;padding:8px;overflow-wrap:anywhere}th{background:#eff6ff}.notice{background:#fff7ed;padding:14px;line-height:1.8}button{padding:9px 18px;cursor:pointer}@page{size:A3 landscape;margin:12mm}@media print{body{margin:0}.no-print{display:none}.chart{overflow:visible;border:0}.chart svg{min-width:0}tr{break-inside:avoid}thead{display:table-header}h2{break-after:avoid}.chart{break-inside:avoid}}'
    return '<!doctype html><html lang="zh-CN"><meta charset="utf-8"><title>'+esc(title)+' · 施工进度汇报</title><style>'+style+'</style><body><button class="no-print" onclick="window.print()">打印 / 另存为 PDF</button><h1>'+esc(title)+' · 施工进度汇报</h1><p class="muted">数据时间：'+esc(model['generated'])+' ｜ 数据来源：子系统同步工作包与施工关系配置</p><div class="stats"><div class="stat">'+total_label+'<b>'+str(model['total'])+' 自然日</b></div><div class="stat">计划开始<b>'+esc(model['start'] or '相对排程')+'</b></div><div class="stat">排程完成<b>'+esc(model['finish'] or '待定')+'</b></div><div class="stat">子系统 / 并行组合<b>'+str(len(nodes))+' / '+str(len(model['parallel']))+'</b></div></div>'+('<div class="notice"><b>计划待核实事项</b><ul>'+warnings+'</ul></div>' if warnings else '')+'<p>口径：自然日，含开工与完工当天；FS 0天表示前项完成后的次日开始，SS 0天表示可同日开始，FF 0天表示完成时间不得早于前项。正间隔为滞后，负间隔为提前。按工作包最早开始日期及前置约束自动排程；无依赖的任务允许并行，未校核人员、材料和场地资源。红色为当前约束下零总时差工作；甘特条深色部分为实际完成率。</p><h2>施工流程与工期关系</h2><div class="chart">'+flow+'</div><h2>工期甘特图</h2><div class="chart">'+''.join(gantt)+'</div><h2>子系统工作包排程明细</h2><table><thead><tr><th>WBS</th><th>子系统</th><th>责任人</th><th>自然日</th><th>排程起止</th><th>完成率</th><th>总时差</th><th>工期来源</th></tr></thead><tbody>'+rows+'</tbody></table><h2>前后置关系</h2><table><tr><th>前置子系统</th><th>后续子系统</th><th>关系</th><th>间隔</th></tr>'+relations+'</table><h2>可并行施工区间</h2><p class="muted">以下为排程日期重叠的组合，实际同时施工需落实资源与作业面。</p><table><tr><th>子系统一</th><th>子系统二</th><th>重叠工期</th></tr>'+overlaps+'</table></body></html>'
