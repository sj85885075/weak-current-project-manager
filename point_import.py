"""Project drawing uploads, isolated CAD conversion, and atomic underlay switching."""
import datetime,hashlib,json,math,os,pathlib,shutil,subprocess,sys,threading,uuid,xml.etree.ElementTree as ET

MAX_FILE=150*1024*1024
JOBS={};LOCK=threading.Lock();CONVERSION_LOCK=threading.Lock()
EXTENSIONS={'.dwg','.dxf','.pdf','.svg','.png','.jpg','.jpeg'}

def pdf_engine():
    runtime=pathlib.Path(__file__).parent/'pointmap_runtime'
    if runtime.exists() and str(runtime) not in sys.path:sys.path.insert(0,str(runtime))
    import pymupdf
    return pymupdf

def configuration(root):
    path=pathlib.Path(root)/'图纸底图配置.json'
    return json.loads(path.read_text(encoding='utf8')) if path.exists() else {}

def asset_path(root):
    root=pathlib.Path(root);config=configuration(root)
    if not config:return root/'CAD原图底层.svg'
    filename=config.get('cad_underlay',{}).get('file','')
    candidate=(root/filename).resolve()
    if not candidate.is_relative_to((root/'图纸导入').resolve()) or candidate.suffix!='.svg':raise ValueError('底图路径无效')
    return candidate

def _number(value,default):
    n=float(default if value in ('',None) else value)
    if not math.isfinite(n) or abs(n)>1e8:raise ValueError('图纸比例或位置无效')
    return n

def upload(root,name,stream,size,options):
    root=pathlib.Path(root);name=pathlib.Path(str(name).replace('\\','/')).name
    suffix=pathlib.Path(name).suffix.lower()
    if suffix not in EXTENSIONS:raise ValueError('请选择 DWG、DXF、PDF、SVG、PNG 或 JPG 图纸')
    if not 0<size<=MAX_FILE:raise ValueError('图纸文件应在150MB以内')
    with LOCK:
        if any(j['root']==str(root) and j['status'] in ('uploading','converting') for j in JOBS.values()):raise ValueError('当前项目已有图纸正在导入')
        id=uuid.uuid4().hex;folder=root/'图纸导入'/id;folder.mkdir(parents=True)
        job={'ok':True,'id':id,'root':str(root),'filename':name,'status':'uploading','message':'接收图纸文件'};JOBS[id]=job
    try:
        source=folder/('source'+suffix)
        with source.open('wb') as f:
            remaining=size
            while remaining:
                chunk=stream.read(min(1024*1024,remaining))
                if not chunk:raise ValueError('上传未完成，请重新导入')
                f.write(chunk);remaining-=len(chunk)
        job.update(status='converting',message='正在转换图纸，较大的 CAD 图纸可能需要几分钟')
        threading.Thread(target=_convert,args=(root,folder,source,options,job),daemon=True).start()
        return {k:v for k,v in job.items() if k!='root'}
    except Exception as e:
        job.update(status='failed',message=str(e));raise

def status(root,id):
    with LOCK:job=JOBS.get(id)
    if not job or job['root']!=str(root):raise ValueError('导入任务不存在，请重新选择文件')
    return {k:v for k,v in job.items() if k!='root'}

def _atomic(path,value):
    tmp=path.with_suffix('.tmp');tmp.write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf8');os.replace(tmp,path)

def _convert(root,folder,source,options,job):
    with CONVERSION_LOCK:_convert_impl(root,folder,source,options,job)

def _convert_impl(root,folder,source,options,job):
    try:
        import point_map
        seed=point_map.seed(root);meta=seed.get('meta',{});old=configuration(root)
        meta.update(old);cad=source.suffix in ('.dwg','.dxf')
        if cad:
            svg,bounds,origin,units=_cad(source,folder,meta,options)
        else:
            svg,ratio,pages=_document(source,options)
            previous=meta.get('cad_underlay',{}).get('bounds',[0,0,100,100])
            width=_number(options.get('width'),previous[2]-previous[0]);x=_number(options.get('x'),previous[0]);y=_number(options.get('y'),previous[1])
            if not .001<=width<=1e7:raise ValueError('图纸实际宽度应大于0')
            bounds=[x,y,x+width,y+width/ratio];origin=meta.get('cad_origin_mm',[0,0]);units=meta.get('drawing_units','mm')
            job['pages']=pages
        dest=folder/'underlay.svg';dest.write_text(svg,encoding='utf8')
        config={'cad_underlay':{'file':str(dest.relative_to(root)).replace('\\','/'),'bounds':bounds,'source':job['filename'],'version':job['id']},'cad_origin_mm':origin,'drawing_units':units,'imported_at':datetime.datetime.now(datetime.timezone.utc).isoformat(),'imported_filename':job['filename'],'import_sha256':hashlib.sha256(source.read_bytes()).hexdigest()}
        if cad:config['cad_source_path']=str(source)
        elif meta.get('cad_source_path'):config['cad_source_path']=''
        path=root/'图纸底图配置.json'
        if path.exists():shutil.copy2(path,folder/'previous_configuration.json')
        else:_atomic(folder/'previous_configuration.json',{'_initial':True})
        _atomic(path,config)
        job.update(status='done',message='图纸导入完成；设备、线路及施工状态已保留',bounds=bounds)
    except Exception as e:job.update(status='failed',message='导入失败：'+str(e))

def _document(source,options):
    pdf=pdf_engine();page_index=int(options.get('page') or 1)-1
    if source.suffix=='.svg':
        raw=source.read_bytes()
        if b'<!DOCTYPE' in raw.upper() or b'<!ENTITY' in raw.upper():raise ValueError('SVG 包含不支持的外部定义')
        root=ET.fromstring(raw)
        allowed={'svg','g','defs','path','rect','line','polyline','polygon','circle','ellipse','text','tspan','use','symbol','clipPath','mask','linearGradient','radialGradient','stop','title','desc'}
        for parent in root.iter():
            for child in list(parent):
                if child.tag.split('}')[-1] not in allowed:parent.remove(child)
            for key,value in list(parent.attrib.items()):
                if key.lower().startswith('on') or 'javascript:' in value.lower() or ('href' in key and not value.startswith('#')) or ('url(' in value.lower() and 'url(#' not in value.lower()):del parent.attrib[key]
        raw=ET.tostring(root,encoding='utf8')
    else:raw=source.read_bytes()
    doc=pdf.open(stream=raw,filetype=source.suffix.lstrip('.'))
    if doc.needs_pass:raise ValueError('请先解密 PDF 再导入')
    if not 0<=page_index<len(doc):raise ValueError(f'页码超出范围，文件共{len(doc)}页')
    page=doc[page_index];ratio=page.rect.width/page.rect.height
    if source.suffix in ('.png','.jpg','.jpeg'):
        import base64
        mime='image/png' if source.suffix=='.png' else 'image/jpeg'
        svg=f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {page.rect.width} {page.rect.height}"><image width="100%" height="100%" href="data:{mime};base64,{base64.b64encode(raw).decode()}"/></svg>'
    else:svg=page.get_svg_image(text_as_path=True)
    return svg,ratio,len(doc)

def _cad(source,folder,meta,options):
    candidates=[os.environ.get('POINTMAP_CAD_CONSOLE',''),r'D:\CAD2026\AutoCAD 2026\accoreconsole.exe']
    console=next((p for p in candidates if p and pathlib.Path(p).is_file()),None)
    if not console:raise ValueError('未找到 AutoCAD 转换器，请安装 AutoCAD 或导入导出的 PDF/SVG')
    # Commands are fixed; uploaded drawing names and paths never become executable Lisp.
    script='''(setvar "FILEDIA" 0)
(setvar "BACKGROUNDPLOT" 0)
(setvar "SECURELOAD" 2)
(setvar "FONTALT" "simsun.ttf")
(command "_.UCS" "_W")
(command "_.ZOOM" "_E")
(setq pmlo (getvar "EXTMIN") pmhi (getvar "EXTMAX") pmu (getvar "INSUNITS"))
(setq pmf (open "coordinates.txt" "w"))
(foreach pmn (list (car pmlo) (cadr pmlo) (car pmhi) (cadr pmhi) pmu) (write-line (rtos pmn 2 9) pmf))
(close pmf)
(entmakex (list '(0 . "LWPOLYLINE") '(100 . "AcDbEntity") '(8 . "0") '(100 . "AcDbPolyline") '(90 . 4) '(70 . 1) (cons 10 (list (car pmlo) (cadr pmlo))) (cons 10 (list (car pmhi) (cadr pmlo))) (cons 10 (list (car pmhi) (cadr pmhi))) (cons 10 (list (car pmlo) (cadr pmhi)))))
(command "_.-PLOT" "_Y" "" "DWG To PDF.pc3" "ISO_A0_(841.00_x_1189.00_MM)" "_M" "_P" "_N" "_W" pmlo pmhi "_F" "_C" "_Y" "monochrome.ctb" "_N" "_A" "drawing.pdf" "_N" "_Y")
(command "_.QUIT" "_Y")
'''
    (folder/'convert.scr').write_text(script,encoding='ascii')
    if options.get('same_extent')=='1' and meta.get('cad_underlay',{}).get('bounds'):
        bounds=[float(n) for n in meta['cad_underlay']['bounds']]
        origin=[float(n) for n in meta.get('cad_origin_mm',[0,0])]
        coords=[bounds[0]+origin[0]/1000,bounds[1]+origin[1]/1000,bounds[2]+origin[0]/1000,bounds[3]+origin[1]/1000]
        assert all(math.isfinite(n) for n in coords)
        override='(setq pmfactor (cond ((= pmu 1) 0.0254) ((= pmu 2) 0.3048) ((= pmu 5) 0.01) ((= pmu 6) 1.0) ((= pmu 7) 1000.0) (T 0.001)))\n'
        override+=f'(setq pmlo (list (/ {coords[0]:.9f} pmfactor) (/ {coords[1]:.9f} pmfactor)) pmhi (list (/ {coords[2]:.9f} pmfactor) (/ {coords[3]:.9f} pmfactor)))\n'
        script=script.replace('(setq pmf (open',override+'(setq pmf (open')
        (folder/'convert.scr').write_text(script,encoding='ascii')
    flags=getattr(subprocess,'CREATE_NO_WINDOW',0)
    env=dict(os.environ);env['ACADLSPASDOC']='0'
    if (folder/'drawing.pdf').exists():(folder/'drawing.pdf').unlink()
    with (folder/'conversion.log').open('wb') as log:
        try:subprocess.run([console,'/i',str(source.resolve()),'/s',str((folder/'convert.scr').resolve())],cwd=folder,stdout=log,stderr=subprocess.STDOUT,timeout=240,creationflags=flags,env=env)
        except subprocess.TimeoutExpired:raise ValueError('CAD 转换超过4分钟，请检查文件能否在 AutoCAD 正常打开')
    path=folder/'drawing.pdf'
    if not path.exists():raise ValueError('AutoCAD 未生成图纸，请查看转换日志或先导出 PDF 后导入')
    numbers=[float(s) for s in (folder/'coordinates.txt').read_text().splitlines()];x0,y0,x1,y1,unit=numbers
    factor={0:.001,1:.0254,2:.3048,4:.001,5:.01,6:1,7:1000}.get(int(unit))
    if not factor:raise ValueError('CAD单位不支持，请将图纸单位改为毫米、厘米或米')
    if x1<=x0 or y1<=y0:raise ValueError('CAD图形范围无效')
    doc=pdf_engine().open(path);page=doc[0]
    frame=max(page.get_drawings(),key=lambda d:d['rect'].get_area())['rect']
    if abs(frame.width/frame.height-(x1-x0)/(y1-y0))>.02*frame.width/frame.height:raise ValueError('CAD转换比例校核失败，请导入PDF并指定实际宽度')
    svg=ET.fromstring(page.get_svg_image(text_as_path=True));svg.set('viewBox',f'{frame.x0} {frame.y0} {frame.width} {frame.height}')
    svg.set('width',str(frame.width));svg.set('height',str(frame.height))
    origin=meta.get('cad_origin_mm') or [x0*factor*1000,y0*factor*1000]
    bounds=[x0*factor-origin[0]/1000,y0*factor-origin[1]/1000,x1*factor-origin[0]/1000,y1*factor-origin[1]/1000]
    return ET.tostring(svg,encoding='unicode'),bounds,origin,'mm'
