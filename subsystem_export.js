/* Full subsystem snapshot: editable tables plus full-range ECharts images. */
async function exportSubsystemBundle(sub, button) {
  const oldText=button.textContent;
  button.disabled=true;button.textContent='正在生成导出包…';
  const temp=[];
  try {
    const card=button.closest('.sub-card')||document.querySelector('#sysPanels .sub-card');
    if(!card)throw Error('请先选择子系统并等待页面加载完成');
    const name=subName(sub),date=systemWorkDate,period=$('#subChartPeriod').selectedOptions[0].textContent;
    const tables=[];
    const cellValue=cell=>{
      const clone=cell.cloneNode(true);
      clone.querySelectorAll('button').forEach(x=>x.remove());
      const original=cell.querySelectorAll('input,select,textarea');
      clone.querySelectorAll('input,select,textarea').forEach((x,i)=>{
        const input=original[i];
        x.replaceWith(document.createTextNode(input.tagName==='SELECT'?input.selectedOptions[0]?.textContent||'':input.value));
      });
      clone.querySelectorAll('br').forEach(x=>x.replaceWith(document.createTextNode('\n')));
      return clone.textContent.trim();
    };
    const overview=[['字段','内容'],['子系统',name],['填报日期',date],['统计周期',period],['导出口径','当前页面输入快照；图形和计量流水为已保存数据'],['图形口径','完整时间范围；乐观/最可能/悲观分别为累计计划的120%/100%/80%，上限为计划总量；另含实际进度']];
    card.querySelectorAll('input:not([type=hidden]),select,textarea').forEach(input=>{
      if(input.closest('table'))return;
      const label=input.previousElementSibling?.tagName==='LABEL'?input.previousElementSibling.textContent:input.id;
      if(label)overview.push([label,input.tagName==='SELECT'?input.selectedOptions[0]?.textContent||'':input.value]);
    });
    tables.push({name:'子系统概览',rows:overview});
    const names=['每日计划','详细工程量','施工步骤','验收记录'];
    card.querySelectorAll('table').forEach((table,i)=>{
      const indices=[...table.querySelectorAll('thead th')].map((th,j)=>/操作/.test(th.textContent)?-1:j).filter(j=>j>=0);
      tables.push({name:names[i]||'表格'+(i+1),rows:[...table.rows].map(row=>indices.map(j=>cellValue(row.cells[j])))});
    });
    const metrics=(state.metrics||[]).filter(r=>r[1]===sub);
    tables.push({name:'全部每日计量流水',rows:[['日期','子系统','工作内容','今日完成量','单位','累计完成量','记录人','备注'],...metrics]});
    const plans=(state.sub_schedules||{})[sub]||{};
    const actual={};metrics.forEach(r=>{const d=String(r[0]).slice(0,10);actual[d]=(actual[d]||0)+(Number(r[3])||0);});
    const dates=[...new Set([...Object.keys(plans),...Object.keys(actual)])].sort();
    let cp=0,ca=0;const total=Object.values(plans).reduce((a,v)=>a+(Number(v)||0),0);
    tables.push({name:'每日进度汇总',rows:[['日期','计划量','实际量','累计计划','累计实际','计划达成率','实际达成率','进度偏差（百分点）'],...dates.map(d=>{
      cp+=Number(plans[d])||0;ca+=actual[d]||0;
      return [d,plans[d]??'',d<=date?(actual[d]??''):'',cp,d<=date?ca:'',total?cp/total:'',total&&d<=date?ca/total:'',total&&d<=date?(ca-cp)/total*100:''];
    })]});
    const charts=[];
    for(const original of subsystemCharts){
      if(original.isDisposed())continue;
      const option=original.getOption(),dom=original.getDom();
      const chartName=dom.parentElement.querySelector('.chart-title')?.textContent||'统计图';
      const rowCount=option.yAxis?.[0]?.type==='category'?(option.yAxis[0].data||[]).length:0;
      const host=document.createElement('div');host.style.cssText='position:fixed;left:-20000px;top:0;width:1600px;height:'+Math.max(520,rowCount*32+140)+'px;';
      document.body.appendChild(host);const chart=echarts.init(host);temp.push({host,chart});
      option.animation=false;option.dataZoom=[];
      option.grid=(option.grid||[]).map(g=>({...g,bottom:75}));
      chart.setOption(option,{notMerge:true});
      charts.push({name:chartName,image:chart.getDataURL({type:'png',pixelRatio:1.5,backgroundColor:'#ffffff'})});
      const axes=option.xAxis?.[0]?.data||[];
      const series=(option.series||[]).filter(s=>s.type!=='heatmap');
      if(series.length)tables.push({name:chartName+'数据',rows:[['日期',...series.map(s=>s.name)],...axes.map((d,i)=>[d,...series.map(s=>{const v=s.data?.[i];return v&&typeof v==='object'?v.value:v??'';})])]});
    }
    const response=await fetch('/api/subsystem_export',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({proj:ui.proj,sub,name,date,period,tables,charts})});
    if(!response.ok){let msg='导出失败（'+response.status+'）';try{msg=(await response.json()).err||msg;}catch{}throw Error(msg);}
    const url=URL.createObjectURL(await response.blob()),a=document.createElement('a');
    a.href=url;a.download=name.replace(/[\\/:*?"<>|]/g,'_')+'_'+date+'_全部表格与图形.zip';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
    toast('已导出全部表格、图形和完整报告');
  }catch(e){toast('导出失败：'+e.message);}
  finally{temp.forEach(({host,chart})=>{chart.dispose();host.remove();});button.disabled=false;button.textContent=oldText;}
}
