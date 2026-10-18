"""Authoritative device counts and idempotent projection into project ledgers."""
import collections,datetime,hashlib,json,os,pathlib,shutil,copy,tempfile

MARK='点位图自动同步'
CATALOG=['枪机','半球','球机','AP','广播','门禁','停车场道闸','人行通道','接入机柜','汇聚机柜']
ACTIVE=set()

def package(device):
    c=device.get('category','')
    if c in ('防爆摄像机','人车分流摄像机','全局摄像机','人脸识别摄像机','全景球机') or any(t in c for t in ('防爆','人车分流','全局','人脸识别','全景')):return '专用摄像机安装调试','commissioned'
    if any(t in c for t in ('枪机','防油烟','热成像')):return '监控枪机安装调试','commissioned'
    if '半球' in c:return '监控半球安装调试','commissioned'
    if '球机' in c:return '监控球机安装调试','commissioned'
    if 'AP' in c.upper() or '无线' in c:return '无线AP安装调试','commissioned'
    if '机柜' in c:return '网络机柜安装','installed'
    if '信息点' in c or '插座' in c:return '信息点安装调试','commissioned'
    return ('设备-' if c[:1] in '=+-@' else '')+c[:42]+'安装调试','commissioned'

def summary(data):
    groups={name:[] for name in CATALOG}
    raw=collections.defaultdict(list)
    for d in data.get('devices',[]):
        c=d.get('category') or '其他设备';raw[c].append(d)
        name= '球机' if '球机' in c else '半球' if '半球' in c else '枪机' if any(t in c for t in ('枪机','防油烟','热成像')) else 'AP' if 'AP' in c.upper() else '机柜（类型待确认）' if c=='机柜' else '专用摄像机' if any(t in c for t in ('防爆','人车分流','全局','人脸识别')) else c
        groups.setdefault(name,[]).append(d)
    def row(name,items):
        n=len(items);i=sum(bool(d.get('installed')) for d in items);c=sum(bool(d.get('commissioned')) for d in items)
        return dict(category=name,total=n,installed=i,commissioned=c,pending_install=n-i,pending_commission=i-c,installation_rate=round(i/n*100,2) if n else None,commissioning_rate=round(c/n*100,2) if n else None)
    return {'categories':[row(k,v) for k,v in groups.items()], 'raw_categories':[row(k,v) for k,v in raw.items()], 'total':row('全部设备',data.get('devices',[]))}

def _json(path,default):
    return json.loads(path.read_text(encoding='utf-8-sig')) if path.exists() else default

def sync_info(root):
    root=pathlib.Path(root);cfg=_json(root/'点位图统计同步设置.json',{})
    state=_json(root/'点位图全局同步.json',{})
    return dict(enabled=bool(cfg.get('enabled')),synced_at=state.get('synced_at',''),status=state.get('status','待同步'),message=state.get('message',''))

def completion_dates(root):
    import point_map
    dates={}
    with point_map.connect(root) as db:
        for r in db.execute('SELECT id,before_data,after_data,updated_at FROM history ORDER BY seq'):
            before=json.loads(r['before_data'] or '{}');after=json.loads(r['after_data'])
            for field in ('installed','commissioned'):
                if bool(after.get(field)) and not bool(before.get(field)):
                    dates[r['id'],field]=datetime.datetime.fromisoformat(r['updated_at']).astimezone().date().isoformat()
                elif not after.get(field):dates.pop((r['id'],field),None)
    return dates

def _save_json(path,value):
    temp=path.with_name(path.name+'.tmp');temp.write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf8');os.replace(temp,path)

def sync(server,proj,force=False):
    if proj in ACTIVE:return dict(ok=True,changed=False)
    ACTIVE.add(proj)
    try:return _sync(server,proj,force)
    except Exception as e:
        root=pathlib.Path(server.project_root(proj));path=root/'点位图全局同步.json';state=_json(path,{})
        state.update(status='同步失败',message=str(e));_save_json(path,state)
        raise
    finally:ACTIVE.discard(proj)

def _sync(server,proj,force=False):
    import point_map
    root=pathlib.Path(server.project_root(proj));config=_json(root/'点位图统计同步设置.json',{})
    if not config.get('enabled'):return dict(ok=True,enabled=False)
    data=point_map.snapshot(root);stats=summary(data);state_path=root/'点位图全局同步.json';previous=_json(state_path,{})
    today=datetime.date.today().isoformat();F=server._paths(proj)
    paths=[pathlib.Path(F[k]) for k in ('p30','p31','p32','p34')]+[root/'子系统计划设置.json',root/'项目管理核心数据.db']
    def stamps():return {p.name:p.stat().st_mtime_ns if p.exists() else None for p in paths if p.suffix!='.db'}
    fingerprint=hashlib.sha256(json.dumps([(d['id'],d.get('category'),bool(d.get('installed')),bool(d.get('commissioned'))) for d in data['devices']],sort_keys=True,ensure_ascii=False).encode()).hexdigest()+today
    if not force and previous.get('status')=='已同步' and previous.get('fingerprint')==fingerprint and previous.get('files')==stamps():return dict(ok=True,enabled=True,changed=False,statistics=stats)
    groups=collections.defaultdict(list)
    fields={}
    for d in data['devices']:
        name,field=package(d);groups[name].append(d);fields[name]=field
    # Keep previously managed packages, including zero totals after deleting the last device.
    for name,field in previous.get('packages',{}).items():groups.setdefault(name,[]);fields.setdefault(name,field)
    backup=root/'点位图同步备份'/datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f');backup.mkdir(parents=True)
    existed={p:p.exists() for p in paths+[state_path]}
    for p in paths+[state_path]:
        if p.exists():shutil.copy2(p,backup/p.name)
    books={k:server.load_workbook(F[k]) for k in ('p30','p31','p32')}
    cfg=server.sub_settings(proj);dates=completion_dates(root)
    def find(ws,col,value):return next((r for r in range(2,ws.max_row+1) if str(ws.cell(r,col).value or '')==value),None)
    def append(ws,values):
        r=ws.max_row+1
        for c,value in enumerate(values,1):
            cell=ws.cell(r,c,value)
            if r>2:cell._style=copy.copy(ws.cell(r-1,c)._style)
            if isinstance(value,str):cell.data_type='s'
        if ws.auto_filter.ref:ws.auto_filter.ref=ws.dimensions
        return r
    try:
        qty=books['p30']['工程量总览'];flows=books['p30']['每日计量'];status=books['p31']['子系统状态'];details=books['p32']['详细工程量']
        managed=set(groups)
        # These packages are counted exclusively from device state. Manual rows remain archived in the backup.
        for r in range(flows.max_row,1,-1):
            if str(flows.cell(r,2).value or '') in managed and str(flows.cell(r,3).value or '') in managed:flows.delete_rows(r)
        for name,devices in groups.items():
            field=fields[name];total=len(devices);complete=[d for d in devices if d.get(field)];actual=len(complete)
            daily=collections.Counter(dates.get((d['id'],field)) for d in complete);daily.pop(None,None)
            day_qty=daily.get(today,0);installed=sum(bool(d.get('installed')) for d in devices);commissioned=sum(bool(d.get('commissioned')) for d in devices)
            status_name='已完成' if total and actual==total else '施工中' if installed or actual else '未开始'
            note=f'{MARK}；总数{total}，已安装{installed}，已调试{commissioned}；完成口径：'+('已安装' if field=='installed' else '已调试')
            qr=next((r for r in range(2,qty.max_row+1) if str(qty.cell(r,2).value or '')==name and str(qty.cell(r,3).value or '')==name),None)
            if qr is None:qr=append(qty,['QTY-PM-'+hashlib.sha1(name.encode()).hexdigest()[:10],name,name,0,'台',0,0])
            for c,v in {4:total,5:'台',6:day_qty,7:actual,8:f'=IF(D{qr}>0,G{qr}/D{qr},"")',9:status_name,10:note,11:MARK,13:status_name}.items():qty.cell(qr,c).value=v
            qty.cell(qr,8).number_format='0.0%'
            sr=find(status,1,name)
            if sr is None:sr=append(status,[name,'未开始','','',0,0,0,'台'])
            for c,v in {2:status_name,5:day_qty,6:actual,7:total,8:'台',9:f'=G{sr}-F{sr}',10:f'=IF(G{sr}>0,F{sr}/G{sr},"")',11:status_name,13:datetime.datetime.now().isoformat(timespec='seconds'),14:note,15:name,16:MARK}.items():status.cell(sr,c).value=v
            status.cell(sr,10).number_format='0.0%'
            dr=next((r for r in range(2,details.max_row+1) if str(details.cell(r,1).value or '')==name and str(details.cell(r,2).value or '')==name),None)
            if dr is None:dr=append(details,[name,name,'台',0,0,0,0,0,MARK,'','','','[]'])
            for c,v in {3:'台',4:total,5:day_qty,6:actual,7:f'=D{dr}-F{dr}',8:f'=IF(D{dr}>0,F{dr}/D{dr},"")',9:MARK,10:note}.items():details.cell(dr,c).value=v
            details.cell(dr,8).number_format='0.0%'
            running=actual-sum(daily.values())
            for date,count in sorted(daily.items()):
                running+=count;append(flows,[date,name,name,count,'台',running,MARK,'根据当前有效点位完成状态计算'])
            if name not in server.SUBS_OF[proj]:
                cfg[name]={'custom':True,'name':name,'start':'','finish':'','daily':{},'schedule_status':'待排期','pointmap_managed':True}
                steps=books['p32']['施工步骤'];append(steps,[name,1,'设备安装','','','已完成' if total and installed==total else '施工中' if installed else '未开始','','',MARK]);append(steps,[name,2,'设备调试','','','已完成' if total and commissioned==total else '施工中' if commissioned else '未开始','','',MARK])
                for kind in ('内部验收','试运行','终验收'):append(books['p32']['验收记录'],[name,kind,'未开始','','','','',''])
            steps=books['p32']['施工步骤']
            for r in range(2,steps.max_row+1):
                if steps.cell(r,1).value!=name or steps.cell(r,9).value!=MARK:continue
                count=installed if steps.cell(r,3).value=='设备安装' else commissioned
                steps.cell(r,6).value='已完成' if total and count==total else '施工中' if count else '未开始'
        # Same statistics in each ledger for installed/debug quantities without double counting primary work quantities.
        for book in books.values():
            ws=book['设备安装调试统计'] if '设备安装调试统计' in book.sheetnames else book.create_sheet('设备安装调试统计')
            ws.delete_rows(1,ws.max_row)
            ws.append(['设备类别','总数量','已安装','未安装','已调试','已安装待调试','安装完成率','调试完成率'])
            for row in stats['categories']+[stats['total']]:
                if row['total']==0 and row['category'] not in CATALOG:continue
                r=ws.max_row+1;ws.append([row['category'],row['total'],row['installed'],row['pending_install'],row['commissioned'],row['pending_commission'],f'=IF(B{r}>0,C{r}/B{r},"")',f'=IF(B{r}>0,E{r}/B{r},"")'])
                ws.cell(r,1).data_type='s'
                ws.cell(r,7).number_format=ws.cell(r,8).number_format='0.0%'
            for c in 'ABCDEFGH':ws.column_dimensions[c].width=23 if c=='A' else 18
            reference=book.worksheets[0]
            for c in range(1,9):ws.cell(1,c)._style=copy.copy(reference.cell(1,min(c,reference.max_column))._style)
            ws.row_dimensions[1].height=28
            ws.freeze_panes='B2';ws.auto_filter.ref=ws.dimensions
        with tempfile.TemporaryDirectory(dir=root,prefix='点位图同步暂存_') as temp:
            for key,book in books.items():book.save(pathlib.Path(temp)/pathlib.Path(F[key]).name);book.close()
            for key in books:os.replace(pathlib.Path(temp)/pathlib.Path(F[key]).name,F[key])
        _save_json(root/'子系统计划设置.json',cfg);server.refresh_subsystems()
        if proj=='titan':server.update_progress_check(proj)
        else:server.update_progress_check_bg(proj)
        # pm_snapshot projects the same authoritative quantities into the global WBS database.
        server.pm_snapshot(proj)
        _save_json(state_path,dict(fingerprint=fingerprint,files=stamps(),packages=fields,synced_at=datetime.datetime.now().isoformat(timespec='seconds'),status='已同步',message='设备统计、30/31/32/34台账和全局工作包已同步',statistics=stats))
        return dict(ok=True,enabled=True,changed=True,statistics=stats)
    except Exception:
        for book in books.values():book.close()
        for p in paths+[state_path]:
            if (backup/p.name).exists():shutil.copy2(backup/p.name,p)
            elif not existed[p] and p.exists():p.unlink()
        server.refresh_subsystems();raise
