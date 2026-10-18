/* 每日记录模块：数据由 server.py 写入 35_项目每日记录表.xlsx。 */
(function(){
  let records=[];
  async function loadDaily(){
    const r=await api('/api/daily?proj='+ui.proj); records=r.records||[]; renderDaily(); renderDailyActions();
  }
  function renderDaily(){
    const date=$('#dailyDate').value||ui.date||todayStr(); $('#dailyDate').value=date;
    const rows=records.slice().sort((a,b)=>String(b[1]).localeCompare(String(a[1])));
    const today=rows.filter(r=>String(r[1])===date), tests=today.filter(r=>r[2]==='测试');
    $('#dailyKpi').innerHTML=[['今日记录',today.length,'统一日志'],['今日测试',tests.reduce((s,r)=>s+Number(r[5]||0),0),'芯/链路'],['已完成',today.filter(r=>r[7]==='完成').length,'项'],['待协调',today.filter(r=>['待确认','存在问题'].includes(r[7])).length,'项']].map(x=>'<div class="card kpi"><div class="klabel">'+x[0]+'</div><div class="kvalue">'+x[1]+'</div><div class="ksub">'+x[2]+'</div></div>').join('');
    $('#dailyTable').innerHTML='<thead><tr><th>日期</th><th>类型</th><th>子系统</th><th>工作内容</th><th>完成量</th><th>结果</th><th>责任人</th><th>位置及备注</th><th>关联报告</th></tr></thead><tbody>'+rows.map(r=>'<tr><td>'+esc(r[1])+'</td><td>'+esc(r[2])+'</td><td>'+esc(r[3])+'</td><td><b>'+esc(r[4])+'</b></td><td>'+esc(r[5])+' '+esc(r[6])+'</td><td>'+esc(r[7])+'</td><td>'+esc(r[8])+'</td><td>'+esc(r[9])+'</td><td>'+esc(r[10])+'</td></tr>').join('')+'</tbody>';
  }
  const oldRoute=route;
  route=function(){
    if(location.hash.replace('#','')!=='daily') return oldRoute();
    document.querySelectorAll('.page').forEach(p=>p.classList.add('hidden')); $('#view-daily').classList.remove('hidden'); $('#pageTitle').textContent='每日记录'; document.querySelectorAll('#nav a').forEach(a=>a.classList.toggle('active',a.dataset.route==='daily')); loadDaily().catch(e=>toast('每日记录读取失败：'+e.message));
  };
  $('#btnDailySave').onclick=async()=>{const rec={date:$('#dailyDate').value,type:$('#dailyType').value,system:$('#dailySystem').value,title:$('#dailyTitle').value.trim(),qty:$('#dailyQty').value,unit:$('#dailyUnit').value.trim(),result:$('#dailyResult').value,owner:$('#dailyOwner').value.trim(),note:$('#dailyNote').value.trim(),report:$('#dailyReport').value.trim()};if(!rec.date||!rec.title){toast('请填写日期和工作内容');return;}const r=await api('/api/daily',{proj:ui.proj,record:rec});if(r.ok){toast(r.msg);$('#dailyTitle').value='';$('#dailyQty').value=0;$('#dailyNote').value='';await loadDaily();}else toast(r.err||'保存失败');};
  $('#dailyDate').onchange=renderDaily;
  
})();
