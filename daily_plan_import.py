"""Validated per-subsystem daily plan imports with whole-batch rollback."""
import base64
import csv
import datetime
import hashlib
import io
import json
import math
import os
from pathlib import Path
import re
import shutil
import zipfile
import openpyxl

LIMIT = 8 * 1024 * 1024


def parse_upload(filename, content):
    if not isinstance(content, str) or len(content) > (LIMIT*4//3+8):
        raise ValueError('文件不能超过 8MB')
    try:raw = base64.b64decode(content, validate=True)
    except Exception:raise ValueError('文件内容无效，请重新选择文件')
    if len(raw)>LIMIT:raise ValueError('文件不能超过 8MB')
    ext=Path(filename).suffix.lower()
    if ext=='.csv':
        try:text=raw.decode('utf-8-sig')
        except UnicodeDecodeError:
            try:text=raw.decode('gb18030')
            except UnicodeDecodeError:raise ValueError('CSV 请使用 UTF-8 或 GB18030 编码')
        rows=list(csv.reader(io.StringIO(text)))
    elif ext=='.xlsx':
        try:
            with zipfile.ZipFile(io.BytesIO(raw)) as archive:
                if sum(x.file_size for x in archive.infolist())>40*1024*1024:raise ValueError('Excel 解压后过大，请仅保留每日计划数据')
            book=openpyxl.load_workbook(io.BytesIO(raw),read_only=True,data_only=False)
            try:
                sheet=book['每日计划导入'] if '每日计划导入' in book.sheetnames else book.worksheets[0]
                if sheet.max_row>5001 or sheet.max_column>50:raise ValueError('最多导入 5000 行、50 列')
                rows=list(sheet.iter_rows(values_only=True))
            finally:book.close()
        except ValueError:raise
        except Exception:raise ValueError('无法读取 Excel，请使用有效的 .xlsx 文件')
    else:raise ValueError('仅支持 .xlsx 或 .csv 文件')
    if not rows or len(rows)>5001:raise ValueError('文件为空或超过 5000 行')
    headers=[str(x or '').strip().lstrip('\ufeff') for x in rows[0]]
    if headers.count('日期')!=1 or headers.count('每日计划量')!=1:
        raise ValueError('首行必须各包含一列“日期”和“每日计划量”，请使用下载模板')
    dc,qc=headers.index('日期'),headers.index('每日计划量')
    items=[];errors=[];seen=set();skipped=0
    for number,row in enumerate(rows[1:],2):
        date=row[dc] if dc<len(row) else None
        qty=row[qc] if qc<len(row) else None
        if qty is None or str(qty).strip()=='':
            if date is not None and str(date).strip():skipped+=1
            continue
        try:
            if isinstance(date,datetime.datetime):date=date.date()
            if isinstance(date,datetime.date):date=date.isoformat()
            else:
                date=str(date or '').strip()
                if not re.fullmatch(r'\d{4}-\d{2}-\d{2}',date):raise ValueError('日期须为 YYYY-MM-DD 或 Excel 日期')
                date=datetime.date.fromisoformat(date).isoformat()
            if date in seen:raise ValueError('日期重复：'+date)
            if isinstance(qty,bool):raise ValueError('计划量须为非负数字')
            value=float(qty)
            if not math.isfinite(value) or value<0:raise ValueError('计划量须为有限的非负数字')
            seen.add(date);items.append({'row':number,'date':date,'qty':value})
        except (ValueError,TypeError) as e:
            errors.append('第%d行：%s'%(number,str(e) if isinstance(e,ValueError) else '日期或数量格式错误'))
    if errors:raise ValueError('；'.join(errors[:20])+('；另有%d项错误'%(len(errors)-20) if len(errors)>20 else ''))
    if not items:raise ValueError('没有可导入数量。请在“每日计划量”列填写数字；空白数量不会写入')
    return sorted(items,key=lambda x:x['date']),skipped


def preview(server, proj, sub, items, skipped=0):
    if proj not in server.SUBS_OF or sub not in server.SUBS_OF[proj]:raise ValueError('项目或子系统无效')
    schedule=server.sub_schedule(proj,sub)
    cfg=server.sub_settings(proj).get(sub,{})
    restricted='daily' in cfg or proj=='titan'
    errors=[];rows=[];merged=dict(schedule)
    for item in items:
        date,qty=item['date'],item['qty']
        if restricted and date not in schedule:errors.append('第%d行：%s 不在当前子系统计划工期内'%(item['row'],date))
        if proj=='titan' and not float(qty).is_integer():errors.append('第%d行：泰坦每日计划数量必须为整数'%item['row'])
        rows.append({'date':date,'before':schedule.get(date),'after':qty})
        merged[date]=qty
    if errors:raise ValueError('；'.join(errors[:20]))
    total=sum(merged.values())
    if not math.isfinite(total):raise ValueError('合计数量超出有效数字范围')
    # Fingerprint includes dependent files so another plan/actual edit cannot silently be overwritten.
    files=server._paths(proj)
    book=server.load_workbook(files['p30'])
    try:
        sheet=book['工程量总览'];col=5 if proj=='titan' else 4
        before_total=sum(float(sheet.cell(r,col).value or 0) for r in range(2,sheet.max_row+1) if server.canonical_sub(proj,sheet.cell(r,2).value)==sub)
    finally:book.close()
    paths=[files[k] for k in ('p30','p31','p32','p34')]+[os.path.join(os.path.dirname(files['p30']),'子系统计划设置.json')]
    hashes={p:hashlib.sha256(Path(p).read_bytes()).hexdigest() if os.path.exists(p) else None for p in paths}
    token=hashlib.sha256(json.dumps([proj,sub,schedule,cfg,hashes],sort_keys=True,ensure_ascii=False).encode()).hexdigest()
    return {'ok':True,'rows':rows,'count':len(rows),'changed':sum(r['before']!=r['after'] for r in rows),
            'skipped':skipped,'before_total':before_total,'after_total':total,'token':token}


def commit(server,proj,sub,items,token,skipped=0):
    check=preview(server,proj,sub,items,skipped)
    if token!=check['token']:raise ValueError('计划或关联数据已变化，请重新选择文件并预览后导入')
    if not check['changed']:return {**check,'msg':'文件与现有每日计划一致，无需修改'}
    files=server._paths(proj)
    root=os.path.dirname(files['p30']);cfg_path=os.path.join(root,'子系统计划设置.json')
    targets=[files[k] for k in ('p30','p31','p32','p34')]+[cfg_path]
    backup=os.path.join(root,'每日计划导入备份',datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    os.makedirs(backup)
    existed={p:os.path.exists(p) for p in targets}
    for p in targets:
        if existed[p]:shutil.copy2(p,os.path.join(backup,os.path.basename(p)))
    cfg=server.sub_settings(proj)
    try:
        updates={r['date']:int(r['qty']) if proj=='titan' else r['qty'] for r in items}
        if sub in cfg and 'daily' in cfg[sub]:
            cfg[sub]['daily'].update(updates)
            with open(cfg_path+'.import.tmp','w',encoding='utf-8') as f:json.dump(cfg,f,ensure_ascii=False,indent=2)
            os.replace(cfg_path+'.import.tmp',cfg_path)
        else:
            book=server.load_workbook(files['p30'])
            try:
                if proj=='titan':
                    sheet=book['每日计划'];col=[3,5,7,9,11,13][server.SUBS_OF[proj].index(sub)]
                    for r in range(2,sheet.max_row+1):
                        date=str(sheet.cell(r,2).value or '')[:10]
                        if str(sheet.cell(r,1).value or '').startswith('第') and date in updates:sheet.cell(r,col).value=updates[date]
                else:
                    sheet=book['子系统每日计划'] if '子系统每日计划' in book.sheetnames else book.create_sheet('子系统每日计划')
                    if sheet.cell(1,1).value is None:
                        for c,value in enumerate(['日期','子系统','计划量'],1):sheet.cell(1,c).value=value
                    found={str(sheet.cell(r,1).value or '')[:10]:r for r in range(2,sheet.max_row+1) if sheet.cell(r,2).value==sub}
                    for date,qty in updates.items():
                        if date in found:sheet.cell(found[date],3).value=qty
                        else:sheet.append([date,sub,qty])
                book.save(files['p30'])
            finally:book.close()
        result=server.save_sub_plan({'proj':proj,'sub':sub,'total':check['after_total'],'preserve_daily':True})
        if not result.get('ok'):raise ValueError(result.get('err') or '计划总量同步失败')
    except Exception:
        for path in targets:
            if existed[path]:shutil.copy2(os.path.join(backup,os.path.basename(path)),path)
            elif os.path.exists(path):os.remove(path)
        raise
    return {**check,'msg':'已导入 %d 个日期，计划总量更新为 %s，统计与图表已同步'%(check['count'],format(check['after_total'],'g'))}


def handle(server,data):
    proj,sub=data.get('proj'),data.get('sub')
    items,skipped=parse_upload(data.get('filename',''),data.get('content',''))
    if data.get('action')=='preview':return preview(server,proj,sub,items,skipped)
    if data.get('action')=='commit':return commit(server,proj,sub,items,data.get('token'),skipped)
    raise ValueError('未知每日计划导入操作')
