"""Editable low voltage design layers above an immutable CAD underlay."""
import collections
import copy
import datetime
import json
import math
import uuid

CATALOG = [
    ('bullet','枪机','视频监控','CCTV-G','devices'),
    ('dome','半球','视频监控','CCTV-D','devices'),
    ('ptz','球机','视频监控','CCTV-P','devices'),
    ('ap','AP','无线网络','AP','devices'),
    ('speaker','广播','公共广播','PA','devices'),
    ('access','门禁','门禁系统','AC','devices'),
    ('barrier','停车场道闸','停车场管理','BG','devices'),
    ('pedestrian','人行通道','人行通道','PG','devices'),
    ('access_cabinet','接入机柜','网络系统','CAB-A','devices'),
    ('aggregation_cabinet','汇聚机柜','网络系统','CAB-D','devices'),
    ('ethernet','网线','综合布线','UTP','lines'),
    ('fiber','光缆','综合布线','FO','lines'),
    ('tray','桥架','桥架系统','TRAY','lines'),
]


def _base():
    import point_map
    return point_map


def _raw(root):
    with _base().connect(root) as db:
        columns={r[1] for r in db.execute('PRAGMA table_info(history)')}
        if 'action' not in columns:
            db.execute("ALTER TABLE history ADD COLUMN action TEXT DEFAULT 'update'")
            db.commit()
        return {r['id']:(json.loads(r['data']),r['revision'],r['updated_at'])
                for r in db.execute('SELECT * FROM records')}


def design_snapshot(root):
    result=_base().seed(root)
    from point_import import configuration
    result.setdefault('meta',{}).update(configuration(root))
    records=_raw(root);known=set();trash=[]
    for kind in ('devices','lines'):
        for item in result.setdefault(kind,[]):
            known.add(item['id'])
            data,revision,stamp=records.get(item['id'],({},0,''))
            item.update(installed=False,commissioned=False,laid=False,measured_length_m=None,
                        note='',revision=revision,updated_at=stamp,record_kind=kind)
            item.update(data)
        for id,(data,revision,stamp) in records.items():
            if id not in known and data.get('custom') and data.get('record_kind')==kind:
                result[kind].append(dict(data,id=id,revision=revision,updated_at=stamp))
        trash.extend(i for i in result[kind] if i.get('_deleted'))
        result[kind]=[i for i in result[kind] if not i.get('_deleted')]
    devices={i['id']:i for i in result['devices']}
    for line in result['lines']:
        pts=line.get('points',[])
        if len(pts)<2:continue
        for field,index in [('from_id',0),('to_id',-1)]:
            device=devices.get(line.get(field))
            if device:pts[index]=[device['x'],device['y']]
        if line.get('custom') or line.get('geometry_override') or line.get('from_id') or line.get('to_id'):
            line['drawing_length_m']=round(sum(math.dist(a,b) for a,b in zip(pts,pts[1:])),3)
        line['from_name']=devices.get(line.get('from_id'),{}).get('name','')
        line['to_name']=devices.get(line.get('to_id'),{}).get('name','')
    result.update(ok=True,trash=trash,catalog=[dict(id=id,name=name,system=system,prefix=prefix,kind=kind)
                  for id,name,system,prefix,kind in CATALOG])
    result['design_summary']={
        'device_types':dict(collections.Counter(i['category'] for i in result['devices'])),
        'routes':dict(collections.Counter(i['category'] for i in result['lines'])),
        'connected_lines':sum(bool(i.get('from_id') and i.get('to_id')) for i in result['lines']),
        'route_lengths_m':dict((category,round(sum(i.get('drawing_length_m',0) for i in result['lines'] if i['category']==category),3)) for category in {i['category'] for i in result['lines']})}
    from point_statistics import summary, sync_info
    result['device_statistics']=summary(result)
    result['global_sync']=sync_info(root)
    if result['global_sync']['enabled']:
        result['meta']['notes']=[n for n in result['meta'].get('notes',[]) if '不会自动反写' not in n]
        result['meta']['notes'].append('设备安装及调试状态自动同步到项目总览、子系统和全局工作包；安装调试包按已调试计完成，机柜安装包按已安装计完成。')
    return result


def _write(root,id,data,revision,action):
    stamp=datetime.datetime.now(datetime.timezone.utc).isoformat(timespec='seconds')
    encoded=json.dumps(data,ensure_ascii=False)
    with _base().connect(root) as db:
        with db:
            db.execute('BEGIN IMMEDIATE')
            before=db.execute('SELECT * FROM records WHERE id=?',(id,)).fetchone()
            if (before['revision'] if before else 0)!=revision:
                raise ValueError('记录已被其他窗口修改，请刷新后重试')
            db.execute('INSERT OR REPLACE INTO records VALUES(?,?,?,?)',(id,encoded,revision+1,stamp))
            db.execute('INSERT INTO history(id,before_data,after_data,updated_at,action) VALUES(?,?,?,?,?)',
                       (id,before['data'] if before else None,encoded,stamp,action))
    return dict(ok=True,id=id,revision=revision+1,msg='已保存')


def design_mutate(root,payload):
    pm=_base();current=design_snapshot(root);action=payload.get('action')
    if action not in ('create','update','delete','restore','duplicate','undo'):
        raise ValueError('未知设计操作')
    all_items=current['devices']+current['lines']+current['trash']
    id=str(payload.get('id') or '');item=next((r for r in all_items if r['id']==id),None)
    patch=payload.get('patch',{})
    if not isinstance(patch,dict):raise ValueError('更新内容无效')
    creating=action in ('create','duplicate')
    if action!='create':
        if item is None:raise ValueError('记录不存在，请刷新')
        if payload.get('revision')!=item['revision']:raise ValueError('记录已被其他窗口修改，请刷新后重试')
        if item.get('_deleted') and action not in ('restore','undo'):raise ValueError('记录已删除，请先恢复')
    kind=payload.get('kind') if action=='create' else item['record_kind']
    if kind not in ('devices','lines'):raise ValueError('记录类型无效')
    catalog=next((c for c in CATALOG if c[0]==patch.get('catalog_id')),None)
    if catalog and catalog[4]!=kind:raise ValueError('设备类型与图层不一致')
    if action=='duplicate':
        patch={k:v for k,v in item.items() if k not in ('id','revision','updated_at','name','handle','source','installed','commissioned','laid','measured_length_m','from_id','to_id','from_name','to_name','record_kind','custom','_deleted')}
        if kind=='devices':patch['points']=[[item['x']+2,item['y']-2]]
        else:patch['points']=[[p[0]+2,p[1]-2] for p in item['points']]
        patch['name']=item['name']+'-副本'
        base=patch['name'];number=2
        while any(i['name']==patch['name'] for i in all_items):patch['name']=base+str(number);number+=1
    if creating:
        id='manual-'+uuid.uuid4().hex;revision=0
        points=pm._points(patch.get('points'),kind=='lines')
        name=patch.get('name') or ''
        if not name.strip():
            prefix=catalog[3] if catalog else ('DEV' if kind=='devices' else 'LINE')
            n=1
            while any(i['name']==f'{prefix}-{n:03d}' for i in all_items):n+=1
            name=f'{prefix}-{n:03d}'
        patch['name']=name
        data=dict(id=id,custom=True,record_kind=kind,source='弱电设计层',points=points,
                  installed=False,commissioned=False,laid=False,measured_length_m=None,note='',
                  zone=patch.get('zone') or '待分区',category=patch.get('category') or (catalog[1] if catalog else ''),
                  system=patch.get('system') or (catalog[2] if catalog else ''),orientation=0)
    else:
        revision=item['revision'];data=copy.deepcopy(_raw(root).get(id,({},0,''))[0])
    if action in ('delete','undo'):
        if action=='undo':
            with pm.connect(root) as db:
                previous=db.execute("SELECT before_data FROM history WHERE id=? ORDER BY seq DESC LIMIT 1",(id,)).fetchone()
            if not previous:raise ValueError('该记录没有可撤销的修改')
            data=json.loads(previous[0]) if previous[0] else (dict(data,_deleted=True) if item.get('custom') else {})
        else:data['_deleted']=True
        if data.get('_deleted') and kind=='devices' and any(id in (line.get('from_id'),line.get('to_id')) for line in current['lines']):
            raise ValueError('该设备仍有连接线路，请先删除线路或调整线路端点')
        if data.get('_deleted') and kind=='lines' and any(line.get('tray_id')==id for line in current['lines']):
            raise ValueError('该桥架仍关联线缆，请先取消线缆的桥架关联')
        return _write(root,id,data,revision,action)
    if action=='restore':
        data['_deleted']=False
        if any(i['id']!=id and i['name']==item['name'] for i in current[kind]):raise ValueError('同编号记录已存在，请先调整编号')
        if kind=='lines':
            for key in ('from_id','to_id','tray_id'):
                if item.get(key) and not any(i['id']==item[key] for i in current['lines' if key=='tray_id' else 'devices']):
                    raise ValueError('请先恢复该线路关联的设备或桥架')
        return _write(root,id,data,revision,action)
    for key,limit in [('name',120),('category',60),('zone',80),('system',80),('model',120),('spec',180),('ip',64),('note',2000),('catalog_id',40)]:
        if key in patch:
            value=pm._text(patch[key],limit)
            if key in ('name','category') and not value:raise ValueError('编号及类别不能为空')
            if key=='name' and any(i['id']!=id and i['name']==value for i in current[kind]):raise ValueError('编号已存在')
            data[key]=value
    for key in ('installed','commissioned','laid'):
        if key in patch:
            if not isinstance(patch[key],bool):raise ValueError('状态必须为勾选值')
            if (kind=='lines')!=(key=='laid'):raise ValueError('状态不适用于该记录')
            data[key]=patch[key]
    for key in ('flip_x','flip_y'):
        if key in patch:
            if kind!='devices' or not isinstance(patch[key],bool):raise ValueError('翻转状态无效')
            data[key]=patch[key]
    merged=dict(item or {},**data)
    if merged.get('commissioned') and not merged.get('installed'):raise ValueError('安装完成后才能标记调试完成')
    for key,maximum,minimum in [('orientation',3600,-3600),('width_mm',10000,0),('height_mm',10000,0),('elevation_m',1000,-20),('cores',432,0),('measured_length_m',1e7,0)]:
        if key in patch:
            value=patch[key]
            if key=='measured_length_m' and kind!='lines':raise ValueError('只有线路可填写长度')
            if value is not None and (isinstance(value,bool) or not isinstance(value,(float,int)) or not math.isfinite(value) or not minimum<=value<=maximum):raise ValueError('长度、角度或规格数值无效')
            if key=='cores' and value is not None and not float(value).is_integer():raise ValueError('光缆芯数须为整数')
            data[key]=value
    if 'points' in patch:
        points=pm._points(patch['points'],kind=='lines');data.update(points=points,geometry_override=True)
        if kind=='devices':data.update(x=points[0][0],y=points[0][1])
    if kind=='lines':
        for key in ('from_id','to_id','tray_id'):
            if key in patch:
                value=pm._text(patch[key],100)
                valid=current['lines'] if key=='tray_id' else current['devices']
                linked=next((d for d in valid if d['id']==value),None)
                if value and (linked is None or (key=='tray_id' and linked.get('category')!='桥架')):raise ValueError('连接设备或桥架不存在')
                if value==id:raise ValueError('不能连接自身')
                data[key]=value
        merged=dict(item or {},**data)
        if merged.get('from_id') and merged.get('from_id')==merged.get('to_id'):raise ValueError('线路两端不能是同一设备')
    return _write(root,id,data,revision,action)


def export_dxf(root):
    """R12 CAD exchange: source DWG xref plus editable symbols and route lines."""
    data=design_snapshot(root);out=[]
    def emit(*values):out.extend(str(v) for v in values)
    def text(value):return str(value).replace('\n',' ').replace('\r',' ')
    origin=data['meta'].get('cad_origin_mm',[47068000,-39337500])
    def xy(p):return [round(p[0]*1000+origin[0],3),round(p[1]*1000+origin[1],3)]
    emit(0,'SECTION',2,'HEADER',9,'$ACADVER',1,'AC1009',9,'$DWGCODEPAGE',3,'ANSI_936',9,'$INSBASE',10,0,20,0,30,0,0,'ENDSEC')
    emit(0,'SECTION',2,'TABLES',0,'TABLE',2,'LAYER',70,5)
    for layer,color in [('LV_DEVICE',3),('LV_UTP',5),('LV_FIBER',6),('LV_TRAY',2),('LV_LABEL',7)]:emit(0,'LAYER',2,layer,70,0,62,color,6,'CONTINUOUS')
    emit(0,'ENDTAB',0,'ENDSEC',0,'SECTION',2,'BLOCKS')
    source=data['meta'].get('cad_source_path','')
    if source:emit(0,'BLOCK',8,'0',2,'CAD_SOURCE',70,4,10,0,20,0,30,0,3,'CAD_SOURCE',1,source,0,'ENDBLK',8,'0')
    emit(0,'ENDSEC',0,'SECTION',2,'ENTITIES')
    if source:emit(0,'INSERT',8,'0',2,'CAD_SOURCE',10,0,20,0,30,0,41,1,42,1,43,1)
    for d in data['devices']:
        x,y=xy([d['x'],d['y']]);emit(0,'CIRCLE',8,'LV_DEVICE',10,x,20,y,30,0,40,350)
        emit(0,'TEXT',8,'LV_LABEL',10,x+450,20,y+450,30,0,40,250,1,text(d['name']+' '+d['category']),50,0)
    for line in data['lines']:
        layer='LV_TRAY' if line['category']=='桥架' else 'LV_FIBER' if '光' in line['category'] else 'LV_UTP'
        for a,b in zip(line['points'],line['points'][1:]):
            a,b=xy(a),xy(b);emit(0,'LINE',8,layer,10,a[0],20,a[1],30,0,11,b[0],21,b[1],31,0)
        mid=xy(line['points'][len(line['points'])//2]);emit(0,'TEXT',8,'LV_LABEL',10,mid[0],20,mid[1]+200,30,0,40,200,1,text(line['name'])+' '+str(line.get('drawing_length_m',0))+'m',50,0)
    emit(0,'ENDSEC',0,'EOF')
    return ('\r\n'.join(out)+'\r\n').encode('gb18030',errors='replace')
