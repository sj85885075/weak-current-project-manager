let subsystemCreateDialog;
function openCreateSubsystem(){
  if(!subsystemCreateDialog){
    subsystemCreateDialog=document.createElement('dialog');
    subsystemCreateDialog.style.cssText='width:620px;max-width:90vw;padding:24px;border:1px solid #d5dce5;border-radius:12px';
    document.body.appendChild(subsystemCreateDialog);
  }
  const proj=ui.proj;
  subsystemCreateDialog.innerHTML='<form id="createSubsystemForm"><h3>添加子系统</h3><p class="hint">新增到当前项目。保存后同步总览、工程量、进度检查、图表和导出。</p><div class="form-row f2"><div><label>子系统名称 *</label><input name="name" required maxlength="60" placeholder="例如：新增光纤测试区域"></div><div><label>计量单位 *</label><input name="unit" required maxlength="20" placeholder="芯 / 条 / 台 / 米"></div><div><label>计划总量</label><input name="total" type="number" min="0" step="'+(isTitan()?'1':'any')+'" placeholder="可留空，后续填写"></div><div><label>负责人</label><input name="owner" maxlength="60"></div><div><label>计划开始 *</label><input name="start" type="date" required></div><div><label>计划结束 *</label><input name="finish" type="date" required></div></div><p class="hint">计划总量按自然日均匀分配。自动建立“主要工程量”、一个施工步骤及三项验收记录，可在添加后调整。</p><p id="createSubsystemError" style="color:#b42318" role="alert"></p><button class="btn" type="submit">添加并同步全局</button> <button class="btn ghost" type="button" id="cancelCreateSubsystem">取消</button></form>';
  const form=subsystemCreateDialog.querySelector('form');
  form.elements.start.value=systemWorkDate;
  form.elements.finish.value=new Date(Date.parse(systemWorkDate)+14*86400000).toISOString().slice(0,10);
  $('#cancelCreateSubsystem').onclick=()=>subsystemCreateDialog.close();
  form.onsubmit=async e=>{
    e.preventDefault();const button=form.querySelector('[type=submit]');button.disabled=true;
    $('#createSubsystemError').textContent='';
    try{
      const data=Object.fromEntries(new FormData(form));data.proj=proj;
      const result=await api('/api/subsystem_create',data);
      if(!result.ok)throw Error(result.err||'添加失败');
      subsystemCreateDialog.close();
      ui.curSub=result.sub;
      await loadAll(proj);
      ui.curSub=result.sub;renderSystems();
      toast(result.msg);
    }catch(err){$('#createSubsystemError').textContent=err.message;}
    finally{button.disabled=false;}
  };
  subsystemCreateDialog.showModal();form.elements.name.focus();
}
