/* 资源、成本与子系统质量台账。 */
(()=>{
  const titles={resources:'资源管理',costs:'成本管理',quality:'质量管理'};
  const fields={
    resources:[['name','资源名称','text',true],['source','获取途径',['甲方提供','采购到货','内部调拨','租赁','其他'],true],['category','资源类型',['线材','设备','辅材','工具','其他'],true],['quantity','数量','number',true],['unit','单位','text',true],['acquired','获取日期','date'],['scope','使用范围 / 子系统','text',true],['use_start','安装 / 使用开始','date'],['use_end','安装 / 使用结束','date'],['owner','负责人','text'],['note','备注','text']],
    costs:[['name','支出内容','text',true],['category','开支类别',['人工','耗材','交通','话费','商务','设备','其他'],true],['amount','金额','number',true],['currency','币种',['CNY','THB','USD','IDR'],true],['date','发生日期','date',true],['scope','使用范围 / 子系统','text'],['owner','经办人','text'],['note','备注 / 凭证编号','text']],
    quality:[['sub','子系统'],['status','施工质量'],['type','验收类型'],['result','验收状态'],['date','验收日期'],['who','验收人'],['concl','验收结论'],['reason','未通过原因'],['fixdate','整改完成日期']]
  };
  let data={resources:[],costs:[]}, loadedProject='', serial=0;
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
      plot(kind,'a','分类开支（按币种分组）',groups(list,r=>r.category+' / '+r.currency,'amount'));plot(kind,'b','月度开支（按币种分组）',Object.fromEntries(Object.entries(groups(list,r=>r.date.slice(0,7)+' / '+r.currency,'amount')).sort()));
    }
    host.querySelector('[data-kpi]').innerHTML=stats.map(([label,value])=>`<div class="card kpi"><div>${esc(label)}</div><div class="val">${esc(value)}</div></div>`).join('');
    host.querySelector('[data-table]').innerHTML='<thead><tr>'+fields[kind].map(f=>'<th>'+f[1]+'</th>').join('')+'<th>操作</th></tr></thead><tbody>'+list.map(r=>'<tr>'+fields[kind].map(f=>'<td>'+esc(r[f[0]]??'—')+'</td>').join('')+'<td style="white-space:nowrap">'+(kind==='quality'?'<button class="btn ghost sm" data-sub="'+esc(r.sub)+'">进入子系统</button>':'<button class="btn ghost sm" data-action="edit" data-id="'+esc(r.id)+'">修改</button> <button class="btn ghost sm" data-action="delete" data-id="'+esc(r.id)+'">删除</button>')+'</td></tr>').join('')+(list.length?'':'<tr><td colspan="'+(fields[kind].length+1)+'" style="padding:30px">暂无明细'+(filters[kind]?'符合搜索条件':(kind==='quality'?'，请在子系统填写验收信息':'，点击“新增明细”开始登记'))+'</td></tr>')+'</tbody>';
  }
  const dialog=document.createElement('dialog');dialog.style.cssText='width:700px;max-width:92vw;border:1px solid #ccd5df;border-radius:12px;padding:24px';document.body.appendChild(dialog);
  function edit(kind,row={}){
    const project=ui.proj;
    dialog.innerHTML='<h3>'+titles[kind]+' · '+(row.id?'修改':'新增')+'明细</h3><form><div class="form-row f2">'+fields[kind].map(([key,label,type,required])=>'<div><label for="management-'+key+'">'+label+(required?' *':'')+'</label>'+(Array.isArray(type)?'<select id="management-'+key+'" name="'+key+'">'+type.map(v=>'<option'+(v===row[key]?' selected':'')+'>'+esc(v)+'</option>').join('')+'</select>':'<input id="management-'+key+'" name="'+key+'" type="'+type+'" value="'+esc(row[key]||'')+'" '+(required?'required ':'')+(type==='number'?'min="0.000001" step="any" ':'')+(key==='scope'?'list="management-subs" ':'')+'>')+'</div>').join('')+'</div><datalist id="management-subs">'+Object.keys(state.sys_status||{}).map(s=>'<option value="'+esc(s)+'">').join('')+'</datalist><p data-error style="color:#b42318" role="alert"></p><button class="btn" type="submit">保存明细</button> <button class="btn ghost" type="button" data-cancel>取消</button></form>';
    dialog.querySelector('[data-cancel]').onclick=()=>dialog.close();
    dialog.querySelector('form').onsubmit=async e=>{e.preventDefault();const button=dialog.querySelector('[type=submit]');button.disabled=true;try{if(project!==ui.proj)throw Error('项目已切换，请重新打开表单');const r=await api('/api/management',{proj:project,kind,action:'save',id:row.id,row:Object.fromEntries(new FormData(e.target))});if(!r.ok)throw Error(r.err);dialog.close();await load();toast(r.msg);}catch(error){dialog.querySelector('[data-error]').textContent=error.message;}finally{button.disabled=false;}};
    dialog.showModal();
  }
  function exportRows(kind){const quote=v=>'"'+String(v??'').replace(/^[=+@-]/,"'$&").replaceAll('"','""')+'"';const csv=[fields[kind].map(f=>quote(f[1])).join(','),...rows(kind).map(r=>fields[kind].map(f=>quote(r[f[0]])).join(','))].join('\r\n');const url=URL.createObjectURL(new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download=ui.proj+'_'+titles[kind]+'.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
  async function load(){const proj=ui.proj,request=++serial;if(loadedProject!==proj){data={resources:[],costs:[]};render(location.hash.slice(1));}try{const result=await api('/api/management?proj='+encodeURIComponent(proj));if(proj!==ui.proj||request!==serial)return;if(result.ok===false)throw Error(result.err);data=result;loadedProject=proj;render(location.hash.slice(1));}catch(e){toast('管理台账读取失败：'+e.message);}}
  const previousRender=renderAll;renderAll=function(){previousRender();if(titles[location.hash.slice(1)])load();};
  const previousRoute=route;route=function(){previousRoute();const kind=location.hash.slice(1);if(titles[kind]){$('#pageTitle').textContent=titles[kind];render(kind);load();}};
  registerResize(()=>Object.values(charts).forEach(c=>{if(c.getDom().offsetWidth)c.resize();}));
  route();
})();
