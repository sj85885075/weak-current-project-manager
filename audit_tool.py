import server, org_deliver, pm_core, tempfile, shutil, json, io, hashlib, datetime, traceback, base64
from pathlib import Path
from decimal import Decimal
import openpyxl
results=[];data_checks=[]
def check(label,fn):
    try:fn();results.append((label,'通过',''))
    except Exception as e:results.append((label,'失败',str(e)));print('FAIL',label,traceback.format_exc())
def require(v,msg='断言失败'):
    if not v:raise AssertionError(msg)
original=server.DIR
tracked=[]
for proj in ('bg','titan'):
    folder=Path(server._paths(proj)['p30']).parent
    for path in folder.iterdir():
        if path.is_file() and path.suffix in ('.xlsx','.json','.db'):tracked.append(path)
hashes={str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in tracked}
for proj in ('bg','titan'):
    F=server._paths(proj);w30=server.load_workbook(F['p30']);w31=server.load_workbook(F['p31']);w32=server.load_workbook(F['p32']);flows={};duplicates=[]
    for r in list(w30['每日计量'].values)[1:]:
        if not r[0]:continue
        key=(str(r[0])[:10],str(r[1]),str(r[2]));flows[key]=flows.get(key,0)+float(r[3] or 0)
    for sub in server.SUBS_OF[proj]:
        schedule=server.sub_schedule(proj,sub);plan=sum(schedule.values());row=next(r for r in w31['子系统状态'].values if r[0]==sub)
        detail=sum(float(r[5] or 0) for r in w32['详细工程量'].values if r[0]==sub)
        actual=sum(v for (d,s,i),v in flows.items() if s==sub)
        bad=[d for d,v in schedule.items() if v<0 or (proj=='titan' and not float(v).is_integer())]
        data_checks.append({'项目':proj,'子系统':sub,'计划合计':plan,'31累计':row[5],'32累计':detail,'流水累计':actual,'异常计划日期':bad,'累计一致':abs(float(row[5] or 0)-detail)<1e-6 and abs(actual-detail)<1e-6})
    for w in (w30,w31,w32):w.close()
with tempfile.TemporaryDirectory() as root:
    Path(root,'泰坦IDC').mkdir()
    for proj in ('bg','titan'):
        source=Path(server._paths(proj)['p30']).parent;dest=Path(root,'泰坦IDC' if proj=='titan' else '')
        for path in source.iterdir():
            if path.is_file() and path.suffix in ('.xlsx','.json','.db'):shutil.copy2(path,dest/path.name)
    server.DIR=root;org_deliver.ROOT=root;original_projects_file=server.CUSTOM_PROJECTS_FILE;server.CUSTOM_PROJECTS_FILE=str(Path(root,'自定义项目.json'))
    def project_create_test():
        result=server.create_project({'name':'验收新增项目','project_no':'TEST-001','project_type':'系统集成项目','plan_start':'2026-10-01','plan_finish':'2026-12-31'})
        require(result['ok']);key=result['project']['key'];require(key in server.SUBS_OF and server.PROJ_NAMES[key]=='验收新增项目' and not server.SUBS_OF[key])
        data=server.bootstrap(key);require(data['meta']['proj']==key and data['meta']['name']=='验收新增项目' and not data['sys_status']);require(Path(server._paths(key)['p30']).exists())
        snap=server.pm_snapshot(key);require(snap['profile']['project_no']=='TEST-001')
        from subsystem_create import create_subsystem,delete_subsystem
        require(create_subsystem(server,{'proj':key,'name':'验收新子系统','unit':'项','total':10,'owner':'测试组','start':'2026-10-01','finish':'2026-10-05'})['ok']);require('验收新子系统' in server.SUBS_OF[key])
        require(delete_subsystem(server,{'proj':key,'sub':'验收新子系统'})['ok']);require('验收新子系统' not in server.SUBS_OF[key])
        server.SUBS_OF.pop(key,None);server.PROJ_NAMES.pop(key,None);server.BASE_SUBS.pop(key,None)
    check('新增项目/独立目录/参数初始化',project_create_test)
    for proj in ('bg','titan'):
        sub=server.SUBS_OF[proj][0]
        check(proj+' 全量读取',lambda:require(len(server.bootstrap(proj)['sys_status'])==len(server.SUBS_OF[proj])))
        def people_test():
            person={'date':'2026-10-08','name':'验收测试员','pos':'测试员','src':'中方','team':'验收组','wt':'测试','area':'2F','content':'隔离验收','hours':8,'note':'备注保留'}
            server.save_payload({'proj':proj,'today':person['date'],'attendance':{person['date']:{'total':1,'cn':1}},'people':[person]})
            d=server.bootstrap(proj);p=next(p for p in d['people'] if p['name']==person['name'])
            require(all(p[k]==v for k,v in person.items()),str(p));require(d['attendance'][person['date']]['total']==1)
            require(any(p['name']=='验收测试员' and p['group']=='验收组' and p['area']=='2F' for p in org_deliver.load_org(proj)['people']))
        check(proj+' 人员/备注/新日期/组织架构联动',people_test)
        def child_test():
            b={'proj':proj,'sub':sub,'kind':'detail','action':'add','name':'验收计量项','unit':'条','total':100,'start':'2026-10-08','finish':'2026-10-12'}
            require(server.save_sub_item(b)['ok']);require(server.sub_actual(proj,'2026-10-08',{sub:{b['name']:{'today':12}}})['ok'])
            server.sub_actual(proj,'2026-10-09',{sub:{b['name']:{'today':20}}});server.sub_actual(proj,'2026-10-08',{sub:{b['name']:{'today':5}}})
            require(server.save_sub_item({**b,'action':'edit','old_name':b['name'],'name':'验收改名项'})['ok'])
            d=server.bootstrap(proj);x=next(x for x in d['details'] if x['item']=='验收改名项');require(x['cum']==25 and x['ps']=='2026-10-08')
            require(next(r['today'] for r in server.sub_actual(proj,'2026-10-08')['details'] if r['item']=='验收改名项')==5)
            require(any(r[1]=='2026-10-09' and r[2]==sub and r[4]>=20 for r in d['progress_check']),'历史实际日期未进入进度统计')
            require(server.save_sub_item({'proj':proj,'sub':sub,'kind':'detail','action':'delete','old_name':'验收改名项'})['ok'])
            require(not any(x['item']=='验收改名项' for x in server.bootstrap(proj)['details']))
            for action,name,old in [('add','验收步骤',''),('edit','验收改名步骤','验收步骤'),('delete','','验收改名步骤')]:
                require(server.save_sub_item({'proj':proj,'sub':sub,'kind':'step','action':action,'name':name,'old_name':old,'start':'2026-10-08','finish':'2026-10-12'})['ok'])
        check(proj+' 子项增改删/日期/历史修正/归档',child_test)
        def plan_test():
            require(server.save_sub_plan({'proj':proj,'sub':sub,'total':1001})['ok'])
            cfg=server.save_sub_settings({'proj':proj,'sub':sub,'name':'验收子系统','start':'2026-10-01','finish':'2026-10-20','start_no':101});require(cfg['ok'])
            before=server.sub_schedule(proj,sub);expected=sum(before.values())
            require(server.save_daily_plan({'proj':proj,'sub':sub,'date':'2026-10-02','qty':123})['ok'])
            after=server.sub_schedule(proj,sub);require(after['2026-10-02']==123);require(all(after[k]==before[k] for k in before if k!='2026-10-02'))
            for _ in range(2):require(server.save_daily_plan({'proj':proj,'sub':sub,'date':'2026-10-02','qty':123})['ok'])
            d=server.bootstrap(proj);require(d['sub_settings'][sub]['name']=='验收子系统');require(d['sys_status'][sub]['ps']=='2026-10-01')
            require(any(r[1]=='2026-10-02' and r[2]==sub and r[3]==123 for r in d['progress_check']),'每日计划未同步进度检查')
            require(server.save_sub_plan({'proj':proj,'sub':sub,'total':0})['ok']);require(sum(server.sub_schedule(proj,sub).values())==0)
        check(proj+' 工期/名称/整数合计/单日精确/重复保存/清零',plan_test)
        def pm_test():
            args=(root,proj,'验收',server.SUBS_OF[proj]);con=pm_core.ensure(*args);con.execute('delete from actual_logs');con.execute('delete from wbs_tasks');con.commit();con.close()
            require(pm_core.save_profile(*args,{'plan_start':'2026-10-01','plan_finish':'2026-10-20'})['ok']);require(pm_core.generate_wbs(*args,{sub:100})['ok'])
            t=pm_core.snapshot(*args)['tasks'][0];require(pm_core.record_actual(*args,t['id'],'2026-10-02',20)['ok']);require(pm_core.update_task(*args,t['id'],{'plan_qty':40})['ok']);require(pm_core.snapshot(*args)['tasks'][0]['actual_qty']==20)
            require(pm_core.delete_task(*args,t['id'])['ok']);require(not any(x['id']==t['id'] for x in pm_core.snapshot(*args)['tasks']))
            content=base64.b64encode(pm_core.wbs_template()).decode('ascii')
            preview=pm_core.import_wbs(*args,'WBS.xlsx',content,True);require(preview['ok'] and preview['rows']==3,str(preview))
            imported=pm_core.import_wbs(*args,'WBS.xlsx',content,False);require(imported['ok'],str(imported))
            snap=pm_core.snapshot(*args);summary=next(x for x in snap['tasks'] if x['wbs_code']=='02.01');leaf=next(x for x in snap['tasks'] if x['wbs_code']=='02.01.01-R1')
            require(summary['is_summary'] and leaf['parent_id']==summary['id']);require(not pm_core.record_actual(*args,summary['id'],'2026-10-03',1)['ok'])
            require(pm_core.record_actual(*args,leaf['id'],'2026-10-03',3)['ok']);summary=next(x for x in pm_core.snapshot(*args)['tasks'] if x['wbs_code']=='02.01');require(summary['display_actual_qty']==3)
            linked=server.pm_snapshot(proj);require(any(x.get('source_type')=='subsystem' for x in linked['tasks']),'子系统未同步到一级WBS')
            source_task=next(x for x in linked['tasks'] if x.get('source_type')=='subsystem')
            body,name=server.export_work_package(proj,source_task['id']);book=openpyxl.load_workbook(io.BytesIO(body))
            require({'工作包说明','工程量明细','施工步骤','验收记录','每日计划与实际'}.issubset(book.sheetnames),'同步工作包导出内容不完整');require(name.endswith('_工作包.xlsx'));book.close()
            body,name=server.export_work_package(proj,leaf['id']);book=openpyxl.load_workbook(io.BytesIO(body))
            require('实际记录' in book.sheetnames and book['实际记录'].max_row>=2,'手工工作包实际流水未导出');book.close()
            body,name=server.export_work_package_summary(proj);book=openpyxl.load_workbook(io.BytesIO(body))
            require(book.sheetnames[:3]==['工作包汇总','工作包明细','每日实际明细'] and book['工作包汇总'].max_row>1 and book['工作包明细'].max_row>1,'工作包汇总或明细导出不完整');require(name.endswith('.xlsx'));book.close()
            rename_target=source_task.get('source_key')
            require(server.save_sub_settings({'proj':proj,'sub':rename_target,'name':'WBS名称同步验收','start':'2026-10-01','finish':'2026-10-20','start_no':201})['ok'])
            linked=server.pm_snapshot(proj);require(not pm_core.record_actual(*args,source_task['id'],'2026-10-04',1)['ok'])
            renamed=next(x for x in linked['tasks'] if x.get('source_key')==rename_target);require(renamed['system_name']=='WBS名称同步验收' and renamed['task_name']=='WBS名称同步验收','子系统名称未同步WBS')
        check(proj+' 项目参数/WBS/实际/修改/删除',pm_test)
        def daily_audit():
            r={'date':'2026-10-08','title':'验收测试','qty':2}
            require(server.save_daily({'proj':proj,'record':r})['ok']);require(server.save_daily({'proj':proj,'record':r})['ok'])
            rows=server.load_daily(proj)['records'];require(len({x[0] for x in rows})==len(rows),'每日记录ID重复')
            require(server.save_audit({'proj':proj,'date':'2026-10-08','desc':'隔离测试问题'})['ok'])
        check(proj+' 每日记录/审计登记',daily_audit)
        def experience_test():
            marker='验收经验库联动'
            server.save_payload({'proj':proj,'today':'2026-10-08',
                'issues_new':[{'sub':sub,'desc':marker+'问题','prio':'P1','owner':'验收负责人','due':'2026-10-10','resp':'施工方','solution':'立即整改'}],
                'todos_new':[{'sub':sub,'item':marker+'待办','pri':'P1','owner':'验收负责人','due':'2026-10-10','note':'跟踪关闭'}],
                'risks_new':[{'sub':sub,'desc':marker+'风险','category':'进度','probability':'高','impact':'高','trigger':'计划偏差','strategy':'减轻','response':'增加资源','owner':'验收负责人'}]})
            d=server.bootstrap(proj)
            issue=next(r for r in d['issues'] if marker+'问题' in str(r))
            require(issue[7 if proj=='titan' else 4]==sub,'问题未关联子系统')
            server.save_payload({'proj':proj,'today':'2026-10-08','issues_updates':{str(issue[0]):{'handled':'已解决','solution':'更换材料并复验通过'}}})
            issue=next(r for r in server.bootstrap(proj)['issues'] if r[0]==issue[0])
            require(issue[18]=='已处理' and issue[14 if proj=='titan' else 15]=='已关闭','解决状态未同步')
            require(issue[10 if proj=='titan' else 12]=='更换材料并复验通过','解决措施未写回')
            todo=next(r for r in d['todos'] if marker+'待办' in str(r));require(todo[9]==sub,'待办未关联子系统')
            risk=next(r for r in d['risks'] if marker+'风险' in str(r));require(risk[1]==sub and risk[6]=='高','风险识别或评级未同步')
        check(proj+' 子系统问题/待办/风险与经验库联动',experience_test)
        tables=['org','people','attendance','qty','metrics','sys_status','steps','details','acceptance','issues','todos','risks','audit','progress_check','progress_total','progress_sub']+(['titan_plan'] if proj=='titan' else [])
        for table in tables:
            def export_test(table=table):
                title,hdr,rows=org_deliver.collect_table(proj,table)
                require(all(len(r)==len(hdr) for r in rows),'导出列数不一致')
                for fmt in ('xlsx','csv','pdf','docx'):
                    path=org_deliver.export_to(proj,table,fmt);require(Path(path).stat().st_size>0)
            check(proj+' 导出 '+table+'（四种格式）',export_test)
        def export_sub():
            body,name=server.export_sub_progress(proj,sub);w=openpyxl.load_workbook(io.BytesIO(body));require(w.active.max_column==13);require(w.active.cell(3,6).value.startswith('='));dates=[str(w.active.cell(r,2).value)[:10] for r in range(3,w.active.max_row+1)];cfg=server.sub_settings(proj)[sub];require(dates[0]==cfg['start'] and dates[-1]==cfg['finish'] and len(dates)==(datetime.date.fromisoformat(cfg['finish'])-datetime.date.fromisoformat(cfg['start'])).days+1,'导出日期未与工期同步');w.close()
        check(proj+' 子系统13列施工表',export_sub)
        for entry in org_deliver.DELIVER_TYPES:
            doc_type=entry[0]
            check(proj+' 交付生成 '+doc_type,lambda doc_type=doc_type:require(org_deliver.gen_deliver(proj,doc_type,'xlsx')['ok']))
    server.DIR=original;org_deliver.ROOT=original;server.CUSTOM_PROJECTS_FILE=original_projects_file
unchanged=all(hashlib.sha256(Path(p).read_bytes()).hexdigest()==h for p,h in hashes.items())
results.append(('真实项目文件哈希保持不变','通过' if unchanged else '失败',''))
report={'time':datetime.datetime.now().isoformat(),'results':results,'real_data':data_checks,'source_unchanged':unchanged}
Path('功能数据联动验收结果.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print('TOTAL',len(results),'PASS',sum(x[1]=='通过' for x in results),'FAIL',sum(x[1]=='失败' for x in results))
print('REAL_DATA_INCONSISTENT',[x for x in data_checks if not x['累计一致'] or x['异常计划日期']])

