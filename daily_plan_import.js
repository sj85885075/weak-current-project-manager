(function(){
  const dialog=document.createElement('dialog');
  dialog.style.cssText='width:760px;max-width:94vw;max-height:85vh;overflow:auto;padding:24px;border:1px solid #cbd5e1;border-radius:12px';
  document.body.appendChild(dialog);
  window.bindDailyPlanImport=function(card,sub){
    const download=card.querySelector('[data-daily-template]'),button=card.querySelector('[data-daily-import]'),fileInput=card.querySelector('[data-daily-file]'),status=card.querySelector('[data-daily-import-status]');
    const proj=ui.proj;
    download.onclick=()=>{location.href='/api/daily_plan_template';};
    button.onclick=()=>fileInput.click();
    fileInput.onchange=async()=>{
      const file=fileInput.files[0];if(!file)return;
      button.disabled=true;status.textContent='正在校验文件…';
      try{
        if(!/\.(xlsx|csv)$/i.test(file.name))throw new Error('请选择 Excel .xlsx 或 CSV 文件');
        if(file.size>8*1024*1024)throw new Error('文件不能超过 8MB');
        const bytes=new Uint8Array(await file.arrayBuffer());let binary='';
        for(let i=0;i<bytes.length;i+=32768)binary+=String.fromCharCode(...bytes.subarray(i,i+32768));
        const payload={proj,sub,filename:file.name,content:btoa(binary)};
        const result=await api('/api/daily_plan_import',{...payload,action:'preview'});
        if(!result.ok)throw new Error(result.err||'校验失败');
        if(ui.proj!==proj||ui.curSub!==sub)throw new Error('当前项目或子系统已切换，请在目标子系统重新导入');
        status.textContent='校验通过，等待确认导入';
        dialog.innerHTML='<h2>每日计划导入预览</h2><p><b>'+esc(subName(sub))+'</b> · '+esc(file.name)+'</p><p>有效日期 '+result.count+' 个，变更 '+result.changed+' 个，跳过空白数量 '+result.skipped+' 行。计划总量：<b>'+esc(result.before_total)+' → '+esc(result.after_total)+'</b></p><p class="hint">只覆盖下面列出的日期，未列出的日期及历史实际完成量保持不变；提交前自动备份计划与关联统计。</p><div style="max-height:360px;overflow:auto"><table><thead><tr><th>日期</th><th>现有计划量</th><th>导入计划量</th><th>变化量</th></tr></thead><tbody>'+result.rows.map(r=>'<tr><td>'+esc(r.date)+'</td><td>'+esc(r.before==null?'未设置':r.before)+'</td><td>'+esc(r.after)+'</td><td>'+esc(Number((r.after-(r.before||0)).toFixed(8)))+'</td></tr>').join('')+'</tbody></table></div><p data-import-error style="color:#dc2626" role="alert"></p><button class="btn" data-import-confirm>确认导入</button> <button class="btn ghost" data-import-cancel>取消</button>';
        dialog.querySelector('[data-import-cancel]').onclick=()=>{dialog.close();status.textContent='已取消，原计划未修改';};
        dialog.querySelector('[data-import-confirm]').onclick=async()=>{
          const confirmButton=dialog.querySelector('[data-import-confirm]');confirmButton.disabled=true;
          try{
            if(ui.proj!==proj||ui.curSub!==sub)throw new Error('当前项目或子系统已切换，请重新导入');
            const saved=await api('/api/daily_plan_import',{...payload,action:'commit',token:result.token});
            if(!saved.ok)throw new Error(saved.err||'导入失败');
            dialog.close();toast(saved.msg);
            if(ui.proj===proj){await loadAll(proj);const details=document.querySelector('#dailyPlanImportDetails');if(details)details.open=true;const newStatus=document.querySelector('[data-daily-import-status]');if(newStatus)newStatus.textContent=saved.msg;}
          }catch(e){dialog.querySelector('[data-import-error]').textContent=e.message;}
          finally{confirmButton.disabled=false;}
        };
        dialog.showModal();
      }catch(e){status.textContent='导入未执行：'+e.message;toast(e.message);}
      finally{button.disabled=false;fileInput.value='';}
    };
  };
})();
