/* 资源、成本与子系统质量台账。 */
(()=>{
  const titles={resources:'资源管理',costs:'成本管理',quality:'质量管理'};
  const fields={
    resources:[['name','资源名称','text',true],['source','获取途径',['甲方提供','采购到货','内部调拨','租赁','其他'],true],['category','资源类型',['线材','设备','辅材','工具','其他'],true],['quantity','数量','number',true],['unit','单位','text',true],['acquired','购买 / 到货日期','date'],['scope','使用范围 / 子系统','text',true],['use_start','安装 / 使用开始','date'],['use_end','安装 / 使用结束','date'],['owner','负责人','text'],['storage_location','存放地点','text'],['note','备注','text']],
    costs:[['name','支出内容','text',true],['category','开支类别',['人工','耗材','交通','话费','商务','设备','其他'],true],['amount','金额','number',true],['currency','币种',['CNY','THB','USD','IDR'],true],['date','发生日期','date',true],['scope','使用范围 / 子系统','text'],['owner','经办人','text'],['note','备注 / 凭证编号','text']],
    quality:[['sub','子系统'],['status','施工质量'],['type','验收类型'],['result','验收状态'],['date','验收日期'],['who','验收人'],['concl','验收结论'],['reason','未通过原因'],['fixdate','整改完成日期']]
  };
  let data={resources:[],costs:[]}, loadedProject='', serial=0, costProfile=null;
  const charts={}, filters={};
  for(const [kind,title] of Object.entries(titles)){
    const link=document.createElement('a');link.href='#'+kind;link.dataset.route=kind;link.innerHTML='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" aria-hidden="true"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8M8 12h8M8 16h5"/></svg>'+title;
    $('#nav').insertBefore(link,$('#nav a[data-route="data"]'));
    const page=document.createElement('section');page.id='view-'+kind;page.className='page hidden';
    page.innerHTML=`<div class="page-head"><h2>${title}</h2><span class="sub">${kind==='quality'?'同步当前项目子系统验收记录；全部适用验收通过才判定达标':'按当前项目独立保存，统计随明细更新'}</span></div><div class="toolbar"><input data-search placeholder="搜索明细" aria-label="搜索明细"><button class="btn ghost" data-refresh>刷新</button><button class="btn ghost" data-export>导出 CSV</button>${kind==='quality'?'':'<button class="btn" data-add>＋ 新增明细</button>'}</div><div data-kpi class="grid g4"></div><div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:14px;margin:14px 0"><div class="card"><div class="chart-title" data-title-a></div><div data-chart-a style="height:300px"></div></div><div class="card"><div class="chart-title" data-title-b></div><div data-chart-b style="height:300px"></div></div></div><div class="card"><div style="overflow:auto"><table data-table></table></div></div>`;
    $('#view-data').parentNode.insertBefore(page,$('#view-data'));
    page.querySelector('[data-search]').oninput=e=>{filters[kind]=e.target.value.trim().toLowerCase();render(kind);};
    page.querySelector('[data-refresh]').onclick=async()=>{if(kind==='quality')await loadAll(ui.proj);else await load();render(kind);};
    page.querySelector('[data-export]').onclick=()=>exportRows(kind);
    if(kind!=='quality')page.querySelector('[data-add]').onclick=()=>edit(kind);
    page.querySelector('[data-table]').onclick=async e=>{
      const b=e.target.closest('button');if(!b)return;
      if(b.dataset.sub){ui.curSub=b.dataset.sub;location.hash='systems';return;}
      const row=(data[kind]||[]).find(r=>r.id===b.dataset.id);if(!row)return;
      if(b.dataset.action==='edit')edit(kind,row);
      else if(confirm('确定删除这条明细？')){try{const r=await api('/api/management',{proj:ui.proj,kind,action:'delete',id:row.id});if(!r.ok)throw Error(r.err);await load();toast(r.msg);}catch(e){toast(e.message);}}
    };
  }
  const costPage=$('#view-costs'), toolbar=costPage.querySelector('.toolbar');
  toolbar.insertAdjacentHTML('beforeend','<button class="btn ghost" data-total>输入 / 修改项目总价</button><button class="btn ghost" data-template>下载支出模板</button><button class="btn" data-import>导入支出清单</button>');
  const salaryNote=document.createElement('p');salaryNote.className='hint';salaryNote.dataset.salarySync='';toolbar.after(salaryNote);
  costPage.querySelector('[data-total]').onclick=async()=>{
    const proj=ui.proj;
    try{
      const pm=await api('/api/pm?proj='+encodeURIComponent(proj));
      if(proj!==ui.proj)return;if(pm.ok===false)throw Error(pm.err);
      dialog.innerHTML='<h3>项目总价</h3><form><label>总价 *</label><input name="amount" type="number" min="0.01" step="any" required value="'+esc(pm.profile?.contract_amount||'')+'"><label>币种</label><select name="currency">'+['IDR','CNY','USD','THB'].map(c=>'<option '+(c===(pm.profile?.contract_currency||'IDR')?'selected':'')+'>'+c+'</option>').join('')+'</select><p class="hint">与项目基础参数中的合同金额同步保存。不同币种支出需要折算后才能计算利润。</p><p data-error role="alert"></p><button class="btn" type="submit">保存总价</button> <button type="button" data-cancel class="btn ghost">取消</button></form>';
      dialog.querySelector('[data-cancel]').onclick=()=>dialog.close();
      dialog.querySelector('form').onsubmit=async e=>{
        e.preventDefault();const b=e.target.querySelector('[type=submit]');b.disabled=true;
        try{if(proj!==ui.proj)throw Error('项目已切换，请重新打开');const form=new FormData(e.target),amount=Number(form.get('amount'));if(!Number.isFinite(amount)||amount<=0)throw Error('请输入大于零的总价');const r=await api('/api/pm',{proj,action:'save_profile',profile:{contract_amount:amount,contract_currency:form.get('currency')}});if(!r.ok)throw Error(r.err);dialog.close();await load();toast(r.msg);}catch(err){dialog.querySelector('[data-error]').textContent=err.message;}finally{b.disabled=false;}
      };dialog.showModal();
    }catch(err){toast(err.message);}
  };
  costPage.querySelector('[data-template]').onclick=()=>{
    const csv=fields.costs.map(f=>f[1]).join(',')+'\r\n';
    const url=URL.createObjectURL(new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'})),a=document.createElement('a');a.href=url;a.download='支出清单导入模板.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
  };
  costPage.querySelector('[data-import]').onclick=()=>{
    const proj=ui.proj;let content='',ready=false;
    dialog.innerHTML='<h3>导入支出清单</h3><p>下载模板后填写，另存为 CSV UTF-8。金额大于零，日期使用 YYYY-MM-DD。类别：人工、耗材、交通、话费、商务、设备、其他；币种：IDR、CNY、USD、THB。最多 5000 条、2 MB。</p><input type="file" accept=".csv" aria-label="支出 CSV 文件"><p class="hint">先预览再导入；完全相同的记录自动跳过。有错误时整批不保存。</p><div data-preview style="max-height:300px;overflow:auto"></div><p data-error role="alert"></p><button class="btn" data-check>校验并预览</button> <button class="btn" data-confirm disabled>确认导入</button> <button class="btn ghost" data-cancel>取消</button>';
    const check=dialog.querySelector('[data-check]'),save=dialog.querySelector('[data-confirm]'),input=dialog.querySelector('input');
    input.onchange=()=>{ready=false;save.disabled=true;content='';dialog.querySelector('[data-preview]').textContent='';};
    dialog.querySelector('[data-cancel]').onclick=()=>dialog.close();
    check.onclick=async()=>{check.disabled=true;save.disabled=true;ready=false;
      try{if(proj!==ui.proj)throw Error('项目已切换，请重新打开');const file=input.files[0];if(!file||!file.name.toLowerCase().endsWith('.csv'))throw Error('请选择 CSV 文件');if(file.size>2000000)throw Error('文件超过 2 MB');content=new TextDecoder('utf-8',{fatal:true}).decode(await file.arrayBuffer());const r=await api('/api/management',{proj,kind:'costs',action:'import',content,dry_run:true});if(!r.ok)throw Error(r.err);dialog.querySelector('[data-error]').textContent='';dialog.querySelector('[data-preview]').innerHTML='<p>'+esc(r.msg)+'；待导入合计：'+esc(Object.entries(r.totals).map(([c,v])=>c+' '+v.toLocaleString()).join('，')||'0')+'</p><p>预览前 20 条</p><table>'+r.preview.map(row=>'<tr>'+fields.costs.map(f=>'<td>'+esc(row[f[0]])+'</td>').join('')+'</tr>').join('')+'</table>';ready=r.count>0;save.disabled=!ready;}catch(err){dialog.querySelector('[data-error]').textContent=err.message;}finally{check.disabled=false;}
    };
    save.onclick=async()=>{if(!ready)return;save.disabled=true;check.disabled=true;input.disabled=true;
      try{if(proj!==ui.proj)throw Error('项目已切换，请重新打开');const r=await api('/api/management',{proj,kind:'costs',action:'import',content,dry_run:false});if(!r.ok)throw Error(r.err);dialog.close();await load();toast(r.msg);}catch(err){dialog.querySelector('[data-error]').textContent=err.message;save.disabled=false;}finally{check.disabled=false;input.disabled=false;}
    };dialog.showModal();
  };
  function quality(){
    const subs=Object.keys(state.sys_status||{});
    for(const a of state.acceptance||[])if(!subs.includes(a.sub))subs.push(a.sub);
    return subs.flatMap(sub=>{
      const list=(state.acceptance||[]).filter(a=>a.sub===sub), applicable=list.filter(a=>a.status!=='不适用');
      const status=list.some(a=>a.status==='未通过')?'未达标':!list.length?'待验收':!applicable.length?'不适用':applicable.every(a=>a.status==='已通过')?'达标':'待验收';
      return (list.length?list:[{}]).map(a=>({...a,sub,status,result:a.status||'未开始'}));
    });
  }
  function rows(kind){const list=kind==='quality'?quality():data[kind]||[];return list.filter(r=>!filters[kind]||Object.values(r).join(' ').toLowerCase().includes(filters[kind]));}
  function groups(list,key,value){const out={};list.forEach(r=>{const k=typeof key==='function'?key(r):r[key]||'未填写';out[k]=(out[k]||0)+(value?Number(r[value]||0):1);});return out;}
  function plot(kind,slot,title,values){
    const host=$('#view-'+kind);host.querySelector('[data-title-'+slot+']').textContent=title;
    const c=charts[kind+slot]||(charts[kind+slot]=echarts.init(host.querySelector('[data-chart-'+slot+']')));
    c.setOption({tooltip:{trigger:'axis'},grid:{left:65,right:20,bottom:85,top:20,containLabel:true},xAxis:{type:'category',data:Object.keys(values),axisLabel:{rotate:25}},yAxis:{type:'value',minInterval:slot==='a'&&kind!=='costs'?1:undefined},series:[{type:'bar',data:Object.values(values),itemStyle:{color:slot==='a'?'#2E4E7E':'#E58A39'},barMaxWidth:48}],graphic:Object.keys(values).length?[]:[{type:'text',left:'center',top:'middle',style:{text:'暂无数据',fill:'#8895a7'}}]},true);c.resize();
  }
  function render(kind){
    if(!titles[kind])return;
    const host=$('#view-'+kind), list=rows(kind);let stats=[];
    if(kind==='quality'){
      const unique=[...new Map(list.map(r=>[r.sub,r])).values()],g=groups(unique,'status');
      stats=[['子系统',unique.length],['达标',g['达标']||0],['未达标',g['未达标']||0],['待验收',g['待验收']||0]];
      plot(kind,'a','子系统质量分布（个）',g);plot(kind,'b','验收记录状态（条）',groups(list,'result'));
    }else if(kind==='resources'){
      stats=[['资源记录',list.length],['甲方提供',list.filter(r=>r.source==='甲方提供').length],['采购到货',list.filter(r=>r.source==='采购到货').length],['已登记使用时间',list.filter(r=>r.use_start).length]];
      plot(kind,'a','获取途径分布（条）',groups(list,'source'));plot(kind,'b','资源类型分布（条；不同单位数量不合计）',groups(list,'category'));
    }else{
      const totals=groups(list,'currency','amount');stats=[['支出记录',list.length],...Object.entries(totals).map(([k,v])=>[k+' 合计',v.toLocaleString(undefined,{maximumFractionDigits:2})])];
      if(true){
        const amount=Number(costProfile?.contract_amount), currency=costProfile?.contract_currency||'IDR', allCosts=data.costs||[];
        const total=allCosts.filter(r=>r.currency===currency).reduce((sum,r)=>sum+Number(r.amount||0),0);
        const mixed=allCosts.some(r=>r.currency!==currency);
        const valid=costProfile&&Number.isFinite(amount)&&amount>0;
        const profit=amount-total, money=v=>currency+' '+v.toLocaleString(undefined,{maximumFractionDigits:2});
        stats.push(['项目总价（'+currency+'）',valid?money(amount):'总价待确认'],['当前利润（全项目）',!valid?'总价待确认':mixed?'存在其他币种，待折算':money(profit)],['利润率',valid&&!mixed?(profit/amount*100).toFixed(2)+'%':'待核算']);
        let note=host.querySelector('[data-profit-note]');
        if(!note){note=document.createElement('p');note.dataset.profitNote='';note.className='hint';host.querySelector('[data-kpi]').after(note);}
        note.textContent='利润 = 项目总价 − 全项目已录入成本；利润率 = 利润 ÷ 项目总价。搜索仅筛选支出统计和明细，盈亏始终按全项目计算。后续新增成本将影响利润。';
      }else{host.querySelector('[data-profit-note]')?.remove();}
      plot(kind,'a','分类开支（按币种分组）',groups(list,r=>r.category+' / '+r.currency,'amount'));plot(kind,'b','月度开支（按币种分组）',Object.fromEntries(Object.entries(groups(list,r=>r.date.slice(0,7)+' / '+r.currency,'amount')).sort()));
    }
    host.querySelector('[data-kpi]').innerHTML=stats.map(([label,value])=>`<div class="card kpi"><div>${esc(label)}</div><div class="val">${esc(value)}</div></div>`).join('');
    host.querySelector('[data-table]').innerHTML='<thead><tr>'+fields[kind].map(f=>'<th>'+f[1]+'</th>').join('')+'<th>操作</th></tr></thead><tbody>'+list.map(r=>'<tr>'+fields[kind].map(f=>'<td>'+esc(r[f[0]]??'—')+'</td>').join('')+'<td style="white-space:nowrap">'+(r.source==='personnel_salary_v1'?'<span class="hint">人员工资自动同步</span>':kind==='quality'?'<button class="btn ghost sm" data-sub="'+esc(r.sub)+'">进入子系统</button>':'<button class="btn ghost sm" data-action="edit" data-id="'+esc(r.id)+'">修改</button> <button class="btn ghost sm" data-action="delete" data-id="'+esc(r.id)+'">删除</button>')+'</td></tr>').join('')+(list.length?'':'<tr><td colspan="'+(fields[kind].length+1)+'" style="padding:30px">暂无明细'+(filters[kind]?'符合搜索条件':(kind==='quality'?'，请在子系统填写验收信息':'，点击“新增明细”开始登记'))+'</td></tr>')+'</tbody>';
  }
  const dialog=document.createElement('dialog');dialog.style.cssText='width:700px;max-width:92vw;border:1px solid #ccd5df;border-radius:12px;padding:24px';document.body.appendChild(dialog);
  function edit(kind,row={}){
    const project=ui.proj;
    dialog.innerHTML='<h3>'+titles[kind]+' · '+(row.id?'修改':'新增')+'明细</h3><form><div class="form-row f2">'+fields[kind].map(([key,label,type,required])=>'<div><label for="management-'+key+'">'+label+(required?' *':'')+'</label>'+(Array.isArray(type)?'<select id="management-'+key+'" name="'+key+'">'+type.map(v=>'<option'+(v===row[key]?' selected':'')+'>'+esc(v)+'</option>').join('')+'</select>':'<input id="management-'+key+'" name="'+key+'" type="'+type+'" value="'+esc(row[key]||'')+'" '+(required?'required ':'')+(type==='number'?'min="0.000001" step="any" ':'')+(key==='scope'?'list="management-subs" ':key==='storage_location'?'list="management-storage" placeholder="选择已有地点或输入新地点" ':'')+'>')+'</div>').join('')+'</div><datalist id="management-subs">'+Object.keys(state.sys_status||{}).map(s=>'<option value="'+esc(s)+'">').join('')+'</datalist><datalist id="management-storage">'+[...new Set((data.resources||[]).map(r=>r.storage_location).filter(Boolean))].map(v=>'<option value="'+esc(v)+'">').join('')+'</datalist><p data-error style="color:#b42318" role="alert"></p><button class="btn" type="submit">保存明细</button> <button class="btn ghost" type="button" data-cancel>取消</button></form>';
    dialog.querySelector('[data-cancel]').onclick=()=>dialog.close();
    dialog.querySelector('form').onsubmit=async e=>{e.preventDefault();const button=dialog.querySelector('[type=submit]');button.disabled=true;try{if(project!==ui.proj)throw Error('项目已切换，请重新打开表单');const r=await api('/api/management',{proj:project,kind,action:'save',id:row.id,row:Object.fromEntries(new FormData(e.target))});if(!r.ok)throw Error(r.err);dialog.close();await load();toast(r.msg);}catch(error){dialog.querySelector('[data-error]').textContent=error.message;}finally{button.disabled=false;}};
    dialog.showModal();
  }
  function exportRows(kind){const quote=v=>'"'+String(v??'').replace(/^[=+@-]/,"'$&").replaceAll('"','""')+'"';const csv=[fields[kind].map(f=>quote(f[1])).join(','),...rows(kind).map(r=>fields[kind].map(f=>quote(r[f[0]])).join(','))].join('\r\n');const url=URL.createObjectURL(new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download=ui.proj+'_'+titles[kind]+'.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
  async function load(){const proj=ui.proj,request=++serial;if(loadedProject!==proj){data={resources:[],costs:[]};costProfile=null;render(location.hash.slice(1));}try{const [result,project]=await Promise.all([api('/api/management?proj='+encodeURIComponent(proj)),api('/api/pm?proj='+encodeURIComponent(proj))]);if(proj!==ui.proj||request!==serial)return;if(result.ok===false)throw Error(result.err);costProfile=project?.profile||null;data=result;const salary=result.salary_sync;salaryNote.textContent=salary?'人员工资已同步 '+salary.count+' 条（按应发计入人工成本）；'+Object.entries(salary.totals||{}).map(([c,v])=>v.toLocaleString('zh-CN')+' '+c).join(' / ')+(salary.warnings?.length?'。待核对：'+salary.warnings.join('；'):''):'人员工资尚未同步，请重启小工具加载新版服务';loadedProject=proj;render(location.hash.slice(1));}catch(e){toast('管理台账读取失败：'+e.message);}}
  const previousRender=renderAll;renderAll=function(){previousRender();if(titles[location.hash.slice(1)])load();};
  const previousRoute=route;route=function(){previousRoute();const kind=location.hash.slice(1);if(titles[kind]){$('#pageTitle').textContent=titles[kind];render(kind);load();}};
  registerResize(()=>Object.values(charts).forEach(c=>{if(c.getDom().offsetWidth)c.resize();}));
  route();
})();

/* Project-scoped contract change ledger; shared by every project. */
(()=>{
  const page=document.createElement('section');page.id='view-changes';page.className='page hidden';
  $('#view-overview').parentNode.appendChild(page);
  page.innerHTML='<div class="page-head"><h2>变更管理</h2><span class="sub">按项目独立登记范围、工程量、费用边界及签署状态</span></div><div class="toolbar"><input data-search aria-label="搜索变更" placeholder="编号、名称、内容"><select data-status aria-label="变更状态"><option value="">全部状态</option>'+['待确认','待签字','已确认','实施中','已完成','已取消'].map(v=>'<option>'+v+'</option>').join('')+'</select><button class="btn ghost" data-refresh>刷新</button><button class="btn ghost" data-export>导出CSV</button><button class="btn" data-add>新增变更</button></div><div class="grid g4" data-kpi></div><p class="hint">变更金额为对合同费用的调整，采购及人工支出在成本管理中登记。待核算金额留空。</p><div class="card" style="overflow:auto"><table data-table></table></div><p data-error role="alert" style="color:#b42318"></p>';
  const style=document.createElement('style');style.textContent='#view-changes td{white-space:normal;min-width:100px;vertical-align:top}#view-changes details{min-width:300px;max-width:650px}#view-changes p{white-space:pre-wrap;overflow-wrap:anywhere}.change-editor{width:880px;max-width:94vw;max-height:90vh;overflow:auto;border:1px solid #ccd5df;border-radius:12px;padding:24px}.change-editor textarea{width:100%;min-height:95px}.change-editor label{display:block;margin-top:8px}';document.head.appendChild(style);
  const dialog=document.createElement('dialog');dialog.className='change-editor';document.body.appendChild(dialog);
  let data=[],loadedProject='',serial=0;
  const fields=[['code','变更编号','text',true],['name','变更名称','text',true],['status','变更状态',['待确认','待签字','已确认','实施中','已完成','已取消']],['date','变更日期（原件未填则留空）','date'],['system','所属系统','text'],['owner','负责人','text'],['original','原设计方案','area'],['description','变更内容','area',true],['materials','设备材料及数量','area'],['cost_rule','费用边界','area'],['cost_status','费用确认状态',['待核算','部分不增费，其余待核算','已确认','不涉及费用']],['amount','整体变更金额（待核算留空，可为减项负数）','number'],['currency','币种',['IDR','CNY','USD','THB']],['schedule_impact','工期影响','text'],['responsibility','责任及实施条件','area'],['signed_by','双方签署代表','text'],['signed_date','双方签署日期','date'],['source_doc','来源文件','text'],['source_id','资料库文件编号（可留空）','text'],['note','备注及待确认事项','area']];
  const rows=()=>data.filter(r=>(!page.querySelector('[data-status]').value||r.status===page.querySelector('[data-status]').value)&&Object.values(r).join(' ').toLowerCase().includes(page.querySelector('[data-search]').value.trim().toLowerCase()));
  function render(){
    const list=rows();
    const stats=[['变更总数',data.length],['待确认 / 待签字',data.filter(r=>['待确认','待签字'].includes(r.status)).length],['实施中',data.filter(r=>r.status==='实施中').length],['已完成',data.filter(r=>r.status==='已完成').length]];
    page.querySelector('[data-kpi]').innerHTML=stats.map(([k,v])=>'<div class="card kpi"><div>'+esc(k)+'</div><div class="val">'+v+'</div></div>').join('');
    page.querySelector('[data-table]').innerHTML='<thead><tr>'+['编号 / 名称','系统 / 状态','日期 / 签署','费用边界 / 金额','变更详情','操作'].map(v=>'<th>'+v+'</th>').join('')+'</tr></thead><tbody>'+list.map(r=>'<tr><td><b>'+esc(r.code)+'</b><br>'+esc(r.name)+'</td><td>'+esc(r.system||'项目级')+'<br>'+esc(r.status)+'</td><td>'+esc(r.date||'变更日期待确认')+'<br>'+esc(r.signed_by||'双方尚未签字')+'<br>'+esc(r.signed_date||'')+'</td><td>'+esc(r.cost_status)+'<br>'+esc(r.amount===''||r.amount==null?'整体金额待核算':Number(r.amount).toLocaleString()+' '+r.currency)+'<p>'+esc(r.cost_rule)+'</p></td><td><details><summary>查看原方案、调整及责任</summary>'+[['原设计',r.original],['变更内容',r.description],['设备材料',r.materials],['责任及实施条件',r.responsibility],['工期影响',r.schedule_impact],['备注',r.note]].map(([k,v])=>'<b>'+k+'</b><p>'+esc(v||'待确认')+'</p>').join('')+'<b>来源文件</b><p>'+esc(r.source_doc||'未关联')+'</p>'+(r.source_id?'<a href="/api/materials/file?'+new URLSearchParams({proj:ui.proj,id:r.source_id})+'" download>下载来源文件</a>':'')+'</details></td><td><button class="btn ghost sm" data-edit="'+esc(r.id)+'">修改</button> <button class="btn ghost sm" data-delete="'+esc(r.id)+'">删除</button></td></tr>').join('')+(list.length?'':'<tr><td colspan="6">暂无变更记录</td></tr>')+'</tbody>';
  }
  function overview(){
    let box=$('#changeOverview');if(!box){box=document.createElement('div');box.id='changeOverview';box.className='card';box.style.margin='14px 0';$('#view-overview').prepend(box);}
    box.innerHTML='<a href="#changes" style="font-weight:bold">变更管理</a><span style="margin-left:18px">'+(loadedProject===ui.proj?'共 '+data.length+' 项；待确认 / 待签字 '+data.filter(r=>['待确认','待签字'].includes(r.status)).length+' 项':'正在读取…')+'</span>';
  }
  async function load(){
    const proj=ui.proj,seq=++serial;if(loadedProject!==proj){data=[];loadedProject='';render();overview();}
    try{const r=await api('/api/management?proj='+encodeURIComponent(proj));if(seq!==serial||proj!==ui.proj)return;if(r.ok===false)throw Error(r.err);if(!Array.isArray(r.changes))throw Error('变更服务尚未加载，请重启小工具');data=r.changes;loadedProject=proj;page.querySelector('[data-error]').textContent='';render();overview();}
    catch(err){if(seq!==serial||proj!==ui.proj)return;page.querySelector('[data-error]').textContent='读取失败：'+err.message;const box=$('#changeOverview');if(box)box.textContent='变更管理读取失败：'+err.message;}
  }
  function edit(row={}){
    const proj=ui.proj;dialog.innerHTML='<h3>变更管理 · '+(row.id?'修改':'新增')+'</h3><form><div class="form-row f2">'+fields.map(([key,label,type,required])=>'<div><label>'+esc(label)+(required?' *':'')+'</label>'+(Array.isArray(type)?'<select name="'+key+'">'+type.map(v=>'<option '+(v===(row[key]||(key==='currency'?'IDR':''))?'selected':'')+'>'+esc(v)+'</option>').join('')+'</select>':type==='area'?'<textarea name="'+key+'" '+(required?'required':'')+'>'+esc(row[key]||'')+'</textarea>':'<input name="'+key+'" type="'+type+'" value="'+esc(row[key]??'')+'" '+(required?'required ':'')+(type==='number'?'step="any"':'')+'>')+'</div>').join('')+'</div><p data-error role="alert" style="color:#b42318"></p><button class="btn" type="submit">保存变更</button> <button class="btn ghost" type="button" data-cancel>取消</button></form>';
    dialog.querySelector('[data-cancel]').onclick=()=>dialog.close();
    dialog.querySelector('form').onsubmit=async e=>{e.preventDefault();const button=e.target.querySelector('[type=submit]');button.disabled=true;try{if(proj!==ui.proj)throw Error('项目已切换，请重新打开');const r=await api('/api/management',{proj,kind:'changes',action:'save',id:row.id,row:Object.fromEntries(new FormData(e.target))});if(!r.ok)throw Error(r.err);dialog.close();await load();toast(r.msg);}catch(err){dialog.querySelector('[data-error]').textContent=err.message;}finally{button.disabled=false;}};dialog.showModal();
  }
  page.querySelector('[data-add]').onclick=()=>edit();page.querySelector('[data-refresh]').onclick=load;
  for(const selector of ['[data-search]','[data-status]'])page.querySelector(selector).addEventListener('input',render);
  page.querySelector('[data-table]').onclick=async e=>{const editButton=e.target.closest('[data-edit]'),deleteButton=e.target.closest('[data-delete]');if(editButton){edit(data.find(r=>r.id===editButton.dataset.edit));return;}if(deleteButton){const proj=ui.proj,row=data.find(r=>r.id===deleteButton.dataset.delete);if(!row||!confirm('删除变更 '+row.code+'？修改历史会保留。'))return;deleteButton.disabled=true;try{const r=await api('/api/management',{proj,kind:'changes',action:'delete',id:row.id});if(!r.ok)throw Error(r.err);await load();toast(r.msg);}catch(err){toast(err.message);}finally{deleteButton.disabled=false;}}};
  page.querySelector('[data-export]').onclick=()=>{const quote=v=>'"'+String(v??'').replace(/^[=+@-]/,"'$&").replaceAll('"','""')+'"';const csv=[fields.map(f=>quote(f[1])).join(','),...rows().map(r=>fields.map(f=>quote(r[f[0]])).join(','))].join('\r\n');const url=URL.createObjectURL(new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'})),a=document.createElement('a');a.href=url;a.download=ui.proj+'_变更管理.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);};
  window.renderChanges=()=>{render();load();};
  const previousOverview=renderOverview;renderOverview=function(){previousOverview();overview();load();};
  const previousRoute=route;route=function(){if(location.hash==='#progress'){location.replace('#changes');return;}previousRoute();};
  if(['#overview','#changes','#progress',''].includes(location.hash))route();
})();


/* Contract collection schedule, scoped to the selected project. */
(()=>{
  const link=document.createElement('a');link.href='#receivables';link.dataset.route='receivables';link.style.paddingLeft='38px';link.textContent='↳ 回款管理';$('#nav [data-route="costs"]').after(link);
  $('#view-costs .toolbar').insertAdjacentHTML('beforeend','<a class="btn ghost" href="#receivables">回款管理</a>');
  const page=document.createElement('section');page.id='view-receivables';page.className='page hidden';$('#view-costs').after(page);
  page.innerHTML='<div class="page-head"><h2>回款管理</h2><span class="sub">按合同节点管理应收与实际回款</span></div><div class="toolbar"><button class="btn ghost" data-refresh>刷新</button><button class="btn" data-plan>配置合同 / 节点日期</button><button class="btn ghost" data-export>导出回款周期明细</button><a class="btn ghost" href="#costs">返回成本管理</a></div><p data-error role="alert" style="color:#b42318"></p><div data-summary class="grid g4"></div><div class="card" data-contract style="margin:14px 0;white-space:pre-wrap;overflow-wrap:anywhere"></div><div class="card" style="overflow:auto"><h3>回款周期明细</h3><table data-stages></table></div><div class="card" style="overflow:auto;margin-top:14px"><h3>实际回款记录</h3><button class="btn" data-receipt>登记回款</button><table data-receipts></table></div><p class="hint">已登记回款仅统计实收金额。未登记不代表尚未收款，请核对银行流水；逾期提示按未核销余额计算。PPh 23 仅针对安装费，扣税基数和凭证需另行核对，本页不自动抵扣应收。</p>';
  const dialog=document.createElement('dialog');dialog.style.cssText='width:760px;max-width:94vw;max-height:90vh;overflow:auto;border:1px solid #ccd5df;border-radius:12px;padding:24px';document.body.append(dialog);
  let data={plan:null,stages:[],receipts:[],version:0},loaded='',serial=0,loading=false;
  const fmt=v=>Number(v||0).toLocaleString(undefined,{maximumFractionDigits:2});
  function render(){
    const plan=data.plan,t=data.totals||{},currency=plan?.currency||'';
    page.querySelector('[data-summary]').innerHTML=[['合同应收',t.planned],['已登记回款',t.received],['待核销余额',t.balance],['逾期未核销',t.overdue]].map(([k,v])=>'<div class="card kpi"><div>'+k+(currency?'（'+esc(currency)+'）':'')+'</div><div class="val">'+(plan?fmt(v):'—')+'</div></div>').join('');
    page.querySelector('[data-contract]').textContent=plan?'合同依据：'+(plan.source||'手工配置')+'\n回款基数：'+fmt(plan.amount)+' '+plan.currency+'；分期比例 50% / 45% / 5%。\n'+(plan.note||''):'当前项目尚未配置合同回款计划。';
    page.querySelector('[data-receipt]').disabled=loading||!plan;page.querySelector('[data-export]').disabled=loading||!plan;page.querySelector('[data-plan]').disabled=loading;
    page.querySelector('[data-stages]').innerHTML='<thead><tr>'+['期次 / 比例','应收金额','触发条件','触发日期','到期日','已登记 / 余额','状态','操作'].map(v=>'<th>'+v+'</th>').join('')+'</tr></thead><tbody>'+data.stages.map(s=>'<tr><td>'+esc(s.name)+'<br>'+s.percent+'%</td><td>'+fmt(s.amount)+' '+esc(currency)+'</td><td style="min-width:260px;white-space:normal">'+esc(s.condition)+'</td><td>'+esc(s.trigger_date||'待确认')+'</td><td>'+esc(s.due_date||'待触发后计算')+'</td><td>'+fmt(s.received)+' / '+fmt(s.balance)+'</td><td>'+esc(s.status)+'</td><td><button class="btn ghost sm" data-stage="'+s.code+'">登记回款</button></td></tr>').join('')+(plan?'':'<tr><td colspan="8">尚无回款周期明细</td></tr>')+'</tbody>';
    const names=Object.fromEntries(data.stages.map(s=>[s.code,s.name]));
    page.querySelector('[data-receipts]').innerHTML='<thead><tr>'+['实收日期','期次','实收金额','凭证 / 流水号','备注','操作'].map(v=>'<th>'+v+'</th>').join('')+'</tr></thead><tbody>'+data.receipts.map(r=>'<tr><td>'+esc(r.date)+'</td><td>'+esc(names[r.stage])+'</td><td>'+fmt(r.amount)+' '+esc(currency)+'</td><td>'+esc(r.reference)+'</td><td style="white-space:normal">'+esc(r.note)+'</td><td><button class="btn ghost sm" data-edit="'+r.id+'">修改</button> <button class="btn ghost sm" data-delete="'+r.id+'">删除</button></td></tr>').join('')+(data.receipts.length?'':'<tr><td colspan="6">暂无已登记回款，请按银行流水核对后登记。</td></tr>')+'</tbody>';
  }
  async function load(){
    const proj=ui.proj,seq=++serial;loading=true;if(loaded!==proj)data={plan:null,stages:[],receipts:[],version:0};render();
    try{const r=await api('/api/management?proj='+encodeURIComponent(proj));if(proj!==ui.proj||seq!==serial)return;if(r.ok===false)throw Error(r.err);if(!r.collections)throw Error('回款服务尚未加载');data=r.collections;loaded=proj;page.querySelector('[data-error]').textContent='';}
    catch(err){if(proj===ui.proj&&seq===serial){loaded='';data={plan:null,stages:[],receipts:[],version:0};page.querySelector('[data-error]').textContent='读取失败：'+err.message;}}
    finally{if(proj===ui.proj&&seq===serial){loading=false;render();}}
  }
  const input=(name,label,value='',type='text',required=false)=>'<div><label>'+esc(label)+'</label><input name="'+name+'" type="'+type+'" value="'+esc(value)+'" '+(required?'required ':'')+(type==='number'?'min="0.01" step="0.01"':'')+'></div>';
  function form(title,body,save){
    if(loading||loaded!==ui.proj){toast('请等待当前项目读取完成');return;}
    const proj=ui.proj,version=data.version;
    dialog.innerHTML='<h3>'+title+'</h3><form><div class="form-row f2">'+body+'</div><p data-error role="alert" style="color:#b42318"></p><button class="btn" type="submit">保存</button> <button class="btn ghost" type="button" data-cancel>取消</button></form>';
    dialog.querySelector('[data-cancel]').onclick=()=>dialog.close();dialog.querySelector('form').onsubmit=async e=>{e.preventDefault();const b=e.target.querySelector('[type=submit]');b.disabled=true;try{if(proj!==ui.proj)throw Error('项目已切换，请重新打开表单');const payload=save(new FormData(e.target));const r=await api('/api/management',{proj,kind:'receivables',expected_version:version,...payload});if(!r.ok)throw Error(r.err);dialog.close();await load();toast(r.msg);}catch(err){dialog.querySelector('[data-error]').textContent=err.message;}finally{b.disabled=false;}};dialog.showModal();
  }
  function editPlan(){const p=data.plan||{};form('合同回款计划 / 触发日期',input('amount','合同应收总额（含税，独立保存为回款基数）',p.amount||'','number',true)+'<div><label>币种</label><select name="currency">'+['IDR','CNY','USD','THB'].map(c=>'<option '+(c===(p.currency||'IDR')?'selected':'')+'>'+c+'</option>').join('')+'</select></div>'+input('source','合同名称及条款来源',p.source||'')+input('advance','合同签署日期（50%，+14日）',p.dates?.advance||'','date')+input('second','设备到货、安装及 BAST 全部满足日期（45%，+14日）',p.dates?.second||'','date')+input('retention','配置、测试、培训及最终验收全部满足日期（5%，+3个月）',p.dates?.retention||'','date')+'<div style="grid-column:1/-1"><label>备注 / 待确认事项</label><textarea name="note" style="width:100%;min-height:90px">'+esc(p.note||'')+'</textarea><p class="hint">当前分期模板为 50% / 45% / 5%。月份按日历月计算，月末取目标月份最后一天。触发日期留空时，该期到期日保持待定。合同总价后续变化时，请在此核对并修改回款基数。</p></div>',f=>({action:'save_plan',row:{amount:f.get('amount'),currency:f.get('currency'),source:f.get('source'),note:f.get('note'),dates:Object.fromEntries(['advance','second','retention'].map(k=>[k,f.get(k)]))}}));}
  function receipt(row={}){form(row.id?'修改回款记录':'登记实际回款','<div><label>回款期次</label><select name="stage">'+data.stages.map(s=>'<option value="'+s.code+'" '+(s.code===row.stage?'selected':'')+'>'+esc(s.name)+'</option>').join('')+'</select></div>'+input('amount','实收金额（'+data.plan.currency+'）',row.amount||'','number',true)+input('date','实收日期',row.date||'','date',true)+input('reference','银行流水 / 凭证编号',row.reference||'')+input('note','备注',row.note||''),f=>({action:'save_receipt',id:row.id,row:Object.fromEntries(f)}));}
  page.querySelector('[data-refresh]').onclick=load;page.querySelector('[data-plan]').onclick=editPlan;page.querySelector('[data-receipt]').onclick=()=>receipt();
  page.querySelector('[data-stages]').onclick=e=>{const b=e.target.closest('[data-stage]');if(b)receipt({stage:b.dataset.stage});};
  page.querySelector('[data-receipts]').onclick=async e=>{const b=e.target.closest('button');if(!b)return;const id=b.dataset.edit||b.dataset.delete,row=data.receipts.find(r=>r.id===id);if(!row)return;if(b.dataset.edit){receipt(row);return;}if(!confirm('删除这笔回款记录？余额将重新计算，修改历史保留。'))return;const proj=ui.proj;b.disabled=true;try{const r=await api('/api/management',{proj,kind:'receivables',action:'delete_receipt',id,expected_version:data.version});if(!r.ok)throw Error(r.err);if(proj===ui.proj)await load();toast(r.msg);}catch(err){toast(err.message);}finally{b.disabled=false;}};
  page.querySelector('[data-export]').onclick=()=>{const q=v=>'"'+String(v??'').replace(/^[=+@-]/,"'$&").replaceAll('"','""')+'"';const rows=[['期次','比例%','币种','应收金额','触发条件','触发日期','到期日','已登记回款','待核销余额','状态','合同依据'],...data.stages.map(s=>[s.name,s.percent,data.plan.currency,s.amount,s.condition,s.trigger_date,s.due_date,s.received,s.balance,s.status,data.plan.source])];const url=URL.createObjectURL(new Blob(['\ufeff'+rows.map(r=>r.map(q).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8'})),a=document.createElement('a');a.href=url;a.download=ui.proj+'_回款周期明细.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);};
  const oldRoute=route;route=function(){if(location.hash!=='#receivables'){oldRoute();return;}document.querySelectorAll('.page').forEach(p=>p.classList.add('hidden'));page.classList.remove('hidden');$('#pageTitle').textContent='回款管理';document.querySelectorAll('#nav a').forEach(a=>a.classList.toggle('active',a.dataset.route==='receivables'));load();};
  const oldRender=renderAll;renderAll=function(){oldRender();if(location.hash==='#receivables')load();};
  render();if(location.hash==='#receivables')route();
})();
