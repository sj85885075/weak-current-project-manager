/* 项目治理模块：统一读取 SQLite 项目参数、过程库、WBS 与变更历史。 */
(function(){
  let pm={};
  let taskFilter='';
  const val=(id,v)=>{ const el=$(id); if(el) el.value=v==null?'':v; };
  const n=v=>Number(v||0);
  const pct=(a,b)=>b?Math.min(100,a/b*100):0;
  async function loadPm(){ const proj=ui.proj; const data=await api('/api/pm?proj='+proj); if(proj!==ui.proj)return; pm=data;if(location.hash==='#pm')renderPm();renderLinked(); }
  function chip(text,kind){ return '<span class="chip '+(kind||'')+'">'+esc(text)+'</span>'; }
  function renderPm(){
    const p=pm.profile||{}, d=pm.dashboard||{}, tasks=pm.tasks||[];
    val('#pmNo',p.project_no); val('#pmType',p.project_type); val('#pmOwner',p.owner_unit); val('#pmContractor',p.contractor);
    val('#pmManager',p.manager); val('#pmTech',p.technical_lead); val('#pmStart',p.plan_start); val('#pmFinish',p.plan_finish);
    val('#pmAmount',p.contract_amount); val('#pmPeople',p.personnel_count); val('#pmHours',p.work_hours||8); val('#pmDays',p.work_days||6);
    const unitText=Object.entries(d.unit_totals||{}).map(([u,x])=>u+' '+x.actual_qty+'/'+x.plan_qty).join(' · ')||'尚未配置工程量';
    $('#pmKpi').innerHTML=[['WBS节点',d.task_count||0,'含汇总层'],['叶子工作包',d.work_package_count||0,'可登记实际'],['实际完成率',(d.actual_rate||0)+'%','计划 '+(d.plan_rate||0)+'%'],['进度状态',d.status||'待配置',(d.variance||0)+'% 偏差']].map(x=>'<div class="card kpi"><div class="klabel">'+x[0]+'</div><div class="kvalue">'+x[1]+'</div><div class="ksub">'+x[2]+'</div></div>').join('');
    const importStatus=$('#wbsImportStatus');if(importStatus&&!importStatus.dataset.message)importStatus.textContent='分单位完成量：'+unitText;
    $('#pmGroups').innerHTML=(pm.groups||[]).map(g=>'<div style="padding:8px 0;border-bottom:1px solid var(--line)"><b>'+esc(g.code)+'</b><div class="hint" style="margin-top:3px">'+esc(g.description)+'</div></div>').join('');
    $('#pmDomains').innerHTML=(pm.domains||[]).map((g,i)=>'<span class="chip '+(i%3===0?'blue':'')+'" style="margin:4px 5px 4px 0;display:inline-block">'+esc(g.name)+'</span>').join('');
    $('#pmProcesses').innerHTML='<thead><tr><th>过程组</th><th>知识领域</th><th>管理过程</th><th>默认工期</th><th>输入</th><th>输出文件</th></tr></thead><tbody>'+(pm.processes||[]).map(x=>'<tr><td>'+esc(x.process_group)+'</td><td>'+esc(x.domain)+'</td><td><b>'+esc(x.name)+'</b></td><td>'+esc(x.default_days)+' 天</td><td>'+esc(x.inputs||'')+'</td><td>'+esc(x.outputs||'')+'</td></tr>').join('')+'</tbody>';
    const shown=tasks.filter(t=>!taskFilter||[t.wbs_code,t.system_name,t.task_name,t.owner].join(' ').toLowerCase().includes(taskFilter));
    $('#pmTasks').innerHTML='<thead><tr><th>WBS</th><th>类型</th><th>子系统 / 工作包</th><th>责任人</th><th>计划起止</th><th>计划量</th><th>累计实际</th><th>完成率</th><th>状态</th><th>今日实际</th><th>操作</th></tr></thead><tbody>'+shown.map(t=>{
      const plan=n(t.display_plan_qty==null?t.plan_qty:t.display_plan_qty),actual=n(t.display_actual_qty==null?t.actual_qty:t.display_actual_qty),rate=pct(actual,plan); const kind=rate>=100?'green':(d.status==='进度落后'?'red':'blue');
      const synced=t.source_type==='subsystem';
      const actualEditor=synced?'<span class="hint">子系统同步</span>':(t.is_summary?'<span class="hint">下级汇总</span>':'<div style="display:flex;gap:5px"><input type="number" min="0" step="any" data-pm-qty="'+esc(t.id)+'" style="width:70px" placeholder="数量"><button class="btn sm" data-pm-add="'+esc(t.id)+'">录入</button></div>');
      const download=(t.is_summary&&!synced)?'':' <button class="btn ghost sm" data-pm-download="'+esc(t.id)+'">下载工作包</button>';
      const actions=(synced?'<button class="btn ghost sm" data-pm-sub="'+esc(t.source_key||t.system_name)+'">进入子系统</button>':'<button class="btn ghost sm" data-pm-edit="'+esc(t.id)+'">修改</button> <button class="btn ghost sm" data-pm-delete="'+esc(t.id)+'">删除</button>')+download;
      const displayUnit=t.display_unit==null?t.unit:t.display_unit;
      return '<tr><td class="num">'+esc(t.wbs_code)+'</td><td>'+chip(t.is_summary?'汇总节点':(t.task_type||'工作包'),t.is_summary?'blue':'')+(synced?'<br><span class="hint">数据源：子系统</span>':'')+'</td><td style="padding-left:'+Math.min(48,Math.max(0,n(t.level)-1)*14)+'px"><b>'+esc(t.task_name)+'</b><br><span class="hint">'+esc(t.system_name||'')+(t.deliverable?' · 交付物：'+esc(t.deliverable):'')+'</span></td><td>'+esc(t.owner||'—')+'</td><td>'+esc(t.plan_start||'—')+'<br>'+esc(t.plan_finish||'—')+'</td><td>'+esc(plan)+' '+esc(displayUnit||'')+'</td><td>'+esc(actual)+' '+esc(displayUnit||'')+'</td><td>'+chip(rate.toFixed(1)+'%',kind)+'</td><td>'+chip(t.status,kind)+'</td><td>'+actualEditor+'</td><td>'+actions+'</td></tr>';
    }).join('')+'</tbody>';
    $('#pmHistory').innerHTML='<thead><tr><th>时间</th><th>操作</th><th>对象</th><th>操作者</th></tr></thead><tbody>'+(pm.history||[]).map(x=>'<tr><td>'+esc(x.created_at)+'</td><td>'+esc(x.action)+'</td><td>'+esc(x.entity_type)+'</td><td>'+esc(x.operator||'—')+'</td></tr>').join('')+'</tbody>';
  }
  function profilePayload(){return {project_no:$('#pmNo').value.trim(),project_type:$('#pmType').value.trim(),owner_unit:$('#pmOwner').value.trim(),contractor:$('#pmContractor').value.trim(),manager:$('#pmManager').value.trim(),technical_lead:$('#pmTech').value.trim(),plan_start:$('#pmStart').value,plan_finish:$('#pmFinish').value,contract_amount:n($('#pmAmount').value),personnel_count:n($('#pmPeople').value),work_hours:n($('#pmHours').value)||8,work_days:n($('#pmDays').value)||6};}
  $('#btnPmSave').onclick=async()=>{const r=await api('/api/pm',{proj:ui.proj,action:'save_profile',profile:profilePayload()});toast(r.ok?r.msg:(r.err||'保存失败'));if(r.ok) await loadPm();};
  $('#btnPmGenerate').onclick=async()=>{
    const quantity_map={}; (state.qty||[]).forEach(x=>{quantity_map[x.sub]=Math.max(n(quantity_map[x.sub]),n(x.total));});
    const r=await api('/api/pm',{proj:ui.proj,action:'generate_wbs',quantity_map}); toast(r.ok?r.msg:(r.err||'生成失败')); if(r.ok) await loadPm();
  };
  const editor=document.createElement('dialog');
  editor.style.cssText='border:1px solid var(--line);border-radius:12px;padding:24px;width:560px;max-width:90vw';
  editor.innerHTML='<h3>修改WBS任务</h3><div class="form-row f2"><div><label>任务名称</label><input id="peName"></div><div><label>任务类型</label><select id="peType"><option>汇总节点</option><option>工作包</option><option>活动</option><option>里程碑</option></select></div><div><label>责任人 / 班组</label><input id="peOwner"></div><div><label>单位</label><input id="peUnit"></div><div><label>计划开始</label><input id="peStart" type="date"></div><div><label>计划完成</label><input id="peFinish" type="date"></div><div><label>计划工期（自然日，含首尾）</label><input id="peDays" type="number" min="1" step="1"></div><div><label>计划完成量 / 总量</label><input id="peQty" type="number" min="0" step="any"></div><div><label>权重</label><input id="peWeight" type="number" min="0" step="any"></div><div><label>前置WBS编码</label><input id="pePredecessor" placeholder="多个编码用逗号分隔"></div></div><div class="form-row" style="margin-top:10px"><div><label>工作说明</label><textarea id="peDescription"></textarea></div><div><label>交付物</label><input id="peDeliverable"></div><div><label>验收标准</label><textarea id="peAcceptance"></textarea></div></div><p class="hint">实际完成量和历史流水不会因修改计划而清除。WBS编码及上级关系通过导入模板维护。</p><button class="btn" id="peSave">保存计划</button> <button class="btn ghost" id="peClear">清空计划参数</button> <button class="btn ghost" id="peCancel">取消</button>';
  document.body.appendChild(editor);
  let editId='';
  function duration(){const a=$('#peStart').value,b=$('#peFinish').value;val('#peDays',a&&b?Math.round((Date.parse(b)-Date.parse(a))/86400000)+1:'');}
  $('#peStart').onchange=duration;$('#peFinish').onchange=duration;
  $('#peDays').onchange=()=>{const a=$('#peStart').value,d=n($('#peDays').value);if(a&&Number.isInteger(d)&&d>0){val('#peFinish',new Date(Date.parse(a)+(d-1)*86400000).toISOString().slice(0,10));}};
  $('#peCancel').onclick=()=>editor.close();
  $('#peClear').onclick=()=>{['#peStart','#peFinish','#peDays','#peQty'].forEach(id=>val(id,''));};
  $('#peSave').onclick=async()=>{const r=await api('/api/pm',{proj:ui.proj,action:'update_task',task_id:editId,task:{task_name:$('#peName').value,task_type:$('#peType').value,owner:$('#peOwner').value,unit:$('#peUnit').value,plan_start:$('#peStart').value,plan_finish:$('#peFinish').value,plan_qty:$('#peQty').value,weight:$('#peWeight').value,predecessor:$('#pePredecessor').value,description:$('#peDescription').value,deliverable:$('#peDeliverable').value,acceptance_criteria:$('#peAcceptance').value}});toast(r.ok?r.msg:r.err);if(r.ok){editor.close();await loadPm();}};
  $('#pmTasks').onclick=async e=>{
    const edit=e.target.dataset.pmEdit,del=e.target.dataset.pmDelete,id=e.target.dataset.pmAdd,sub=e.target.dataset.pmSub,download=e.target.dataset.pmDownload;
    if(download){location.href='/api/work_package_export?proj='+encodeURIComponent(ui.proj)+'&task_id='+encodeURIComponent(download);return;}
    if(sub){ui.curSub=sub;location.hash='systems';return;}
    if(edit){const t=(pm.tasks||[]).find(t=>t.id===edit);if(!t)return;editId=edit;val('#peName',t.task_name);val('#peType',t.task_type||'工作包');val('#peOwner',t.owner);val('#peUnit',t.unit);val('#peStart',t.plan_start);val('#peFinish',t.plan_finish);val('#peQty',t.plan_qty);val('#peWeight',t.weight==null?1:t.weight);val('#pePredecessor',t.predecessor);val('#peDescription',t.description);val('#peDeliverable',t.deliverable);val('#peAcceptance',t.acceptance_criteria);duration();editor.showModal();return;}
    if(del){const t=(pm.tasks||[]).find(t=>t.id===del);if(!confirm('删除计划任务“'+t.task_name+'”？该任务将退出统计，实际记录和变更历史保留。'))return;const r=await api('/api/pm',{proj:ui.proj,action:'delete_task',task_id:del});toast(r.ok?r.msg:r.err);if(r.ok)await loadPm();return;}
    if(!id)return;const input=document.querySelector('[data-pm-qty="'+id+'"]');const qty=n(input.value);if(!Number.isFinite(qty)||qty<=0){toast('请输入大于 0 的实际完成量');return;}const r=await api('/api/pm',{proj:ui.proj,action:'record_actual',task_id:id,work_date:ui.date||todayStr(),qty});toast(r.ok?r.msg:(r.err||'录入失败'));if(r.ok)await loadPm();
  };
  $('#pmTaskFilter').oninput=e=>{taskFilter=e.target.value.trim().toLowerCase();renderPm();};
  $('#btnWbsTemplate').onclick=()=>{location.href='/api/wbs_template';};
  $('#btnWorkPackageSummary').onclick=()=>{location.href='/api/work_package_summary?proj='+encodeURIComponent(ui.proj);};
  $('#btnWbsImport').onclick=()=>$('#wbsImportFile').click();
  $('#wbsImportFile').onchange=async e=>{
    const file=e.target.files[0];if(!file)return;
    try{
      if(file.size>8*1024*1024)throw new Error('文件不能超过 8MB');
      const bytes=new Uint8Array(await file.arrayBuffer());let binary='';for(let i=0;i<bytes.length;i+=0x8000)binary+=String.fromCharCode(...bytes.subarray(i,i+0x8000));const content=btoa(binary);
      const preview=await api('/api/pm',{proj:ui.proj,action:'import_wbs',filename:file.name,content,dry_run:true});
      const status=$('#wbsImportStatus');status.dataset.message='1';
      if(!preview.ok){status.textContent='校验失败：'+(preview.errors||[preview.err||'未知错误']).slice(0,5).join('；');toast('WBS导入校验未通过');return;}
      status.textContent='校验通过：共'+preview.rows+'项，新增'+preview.new+'项，更新'+preview.update+'项。';
      if(!confirm(status.textContent+' 确认写入当前项目吗？'))return;
      const result=await api('/api/pm',{proj:ui.proj,action:'import_wbs',filename:file.name,content,dry_run:false});
      if(!result.ok)throw new Error((result.errors||[result.err||'导入失败']).join('；'));
      toast(result.msg);status.textContent=result.msg;await loadPm();
    }catch(err){toast(err.message||'导入失败');$('#wbsImportStatus').textContent='导入失败：'+(err.message||'未知错误');}
    finally{e.target.value='';}
  };
  const linkedCharts=[];
  function renderLinked(){
    ['#view-overview','#view-progress'].forEach((selector,i)=>{
      const host=$(selector);if(!host||host.classList.contains('hidden'))return;
      let card=host.querySelector('[data-pm-linked]');
      if(!card){card=document.createElement('div');card.className='card';card.dataset.pmLinked='1';card.style.marginBottom='14px';card.innerHTML='<div class="chart-title">WBS 计划与实际统计</div><div class="hint" data-pm-summary></div><div data-pm-chart style="height:300px"></div>';host.appendChild(card);}
      const d=pm.dashboard||{},p=pm.profile||{};
      card.querySelector('[data-pm-summary]').textContent='计划起止 '+(p.plan_start||'未配置')+' ~ '+(p.plan_finish||'未配置')+' · 节点 '+(d.task_count||0)+' · 叶子工作包 '+(d.work_package_count||0)+' · 完成率 '+(d.actual_rate||0)+'% · 偏差 '+(d.variance||0)+'%（按叶子任务、分单位计量）';
      const chartTasks=(pm.tasks||[]).filter(t=>!t.is_summary);
      const chart=linkedCharts[i]||(linkedCharts[i]=echarts.init(card.querySelector('[data-pm-chart]')));
      chart.setOption({tooltip:{trigger:'axis'},legend:{data:['计划量','实际量']},grid:{bottom:80,containLabel:true},xAxis:{type:'category',data:chartTasks.map(t=>t.wbs_code+' '+t.task_name),axisLabel:{rotate:25}},yAxis:{type:'value'},series:[{name:'计划量',type:'bar',data:chartTasks.map(t=>n(t.plan_qty))},{name:'实际量',type:'bar',data:chartTasks.map(t=>n(t.actual_qty))}]},true);chart.resize();
    });
  }
  registerResize(()=>linkedCharts.forEach(c=>{if(c&&c.getDom().offsetWidth)c.resize();}));
  const oldRenderAll=renderAll;
  renderAll=function(){oldRenderAll();if(['','overview','progress','pm'].includes(location.hash.slice(1)))loadPm().catch(e=>toast('计划统计读取失败：'+e.message));};
  const oldRoute=route;
  route=function(){if(location.hash.replace('#','')!=='pm'){const result=oldRoute();if(['','overview','progress'].includes(location.hash.slice(1)))loadPm().catch(e=>toast(e.message));linkedCharts.forEach(c=>{if(c&&c.getDom().offsetWidth)c.resize();});return result;}document.querySelectorAll('.page').forEach(p=>p.classList.add('hidden'));$('#view-pm').classList.remove('hidden');$('#pageTitle').textContent='项目管理体系';document.querySelectorAll('#nav a').forEach(a=>a.classList.toggle('active',a.dataset.route==='pm'));loadPm().catch(e=>toast('项目管理数据读取失败：'+e.message));};
  
})();
