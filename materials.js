/* 项目资料库：批量上传、分类、预览与本机 CAD 打开。 */
(()=>{
  const categories=['甲方资料','商务资料','技术资料','图纸资料','项目资料'];
  let files=[],project='',request=0,uploading=false;
  const page=document.createElement('section');page.className='page hidden';page.id='view-materials';
  page.innerHTML=`<div class="page-head"><h2>资料管理</h2><span class="sub">按项目分类归档 · 支持批量上传 · 同名文件分别保留</span></div>
    <div class="grid" style="grid-template-columns:repeat(5,minmax(0,1fr));margin-bottom:16px" data-summary></div>
    <div class="card" style="margin-bottom:16px"><div class="chart-title">上传资料</div><div class="toolbar" style="flex-wrap:wrap">
    <label>上传分类 <select data-category>${categories.map(c=>'<option>'+c+'</option>').join('')}</select></label>
    <label>资料文件 <input type="file" multiple data-files style="max-width:300px"></label>
    <label>备注 <input data-note maxlength="2000" placeholder="版本、用途或来源说明"></label><button class="btn" data-upload>上传资料</button></div>
    <p class="hint">支持各种文件格式，单个文件最大 100 MB。PDF、图片和文本可在线预览；表格、文档和演示文件优先使用 WPS 打开，CAD 图纸调用本机默认 CAD 软件。</p><p data-status role="status" style="margin-top:10px"></p></div>
    <div class="toolbar"><label>分类筛选 <select data-filter><option value="">全部资料</option>${categories.map(c=>'<option>'+c+'</option>').join('')}</select></label><input data-search placeholder="搜索文件名或备注" aria-label="搜索文件名或备注"><button class="btn ghost" data-refresh>刷新资料</button><span class="hint" data-count></span></div>
    <div class="card" style="overflow:auto"><table><thead><tr><th>文件名</th><th>分类</th><th>格式</th><th>大小</th><th>上传时间</th><th>备注</th><th>操作</th></tr></thead><tbody data-rows></tbody></table></div>`;
  $('#view-data').parentNode.insertBefore(page,$('#view-data'));
  const link=document.createElement('a');link.href='#materials';link.dataset.route='materials';link.innerHTML='<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" aria-hidden="true"><path d="M3 6h7l2 3h9v11H3z"/></svg>资料管理';$('#nav').insertBefore(link,$('#nav a[data-route="data"]'));
  const $m=s=>page.querySelector(s);
  const size=n=>n<1024?n+' B':n<1048576?(n/1024).toFixed(1)+' KB':(n/1048576).toFixed(1)+' MB';
  const url=(f,preview=false)=>'/api/materials/file?'+new URLSearchParams({proj:project,id:f.id,...(preview?{preview:'1'}:{})});
  function render(){
    $m('[data-summary]').innerHTML=categories.map(c=>`<div class="card kpi"><div class="label">${c}</div><div class="val">${files.filter(f=>f.category===c).length}<small> 份</small></div></div>`).join('');
    const query=$m('[data-search]').value.toLowerCase(),category=$m('[data-filter]').value;
    const shown=files.filter(f=>(!category||f.category===category)&&(!query||(f.name+' '+f.note).toLowerCase().includes(query)));
    $m('[data-count]').textContent='当前 '+shown.length+' 份 / 共 '+files.length+' 份';
    $m('[data-rows]').innerHTML=shown.map(f=>`<tr><td style="text-align:left;max-width:280px;overflow-wrap:anywhere"><b>${esc(f.name)}</b></td><td>${esc(f.category)}</td><td>${esc(f.name.includes('.')?f.name.split('.').pop().toUpperCase():'—')}</td><td>${size(f.size)}</td><td>${esc(f.uploaded.replace('T',' '))}</td><td style="max-width:220px;overflow-wrap:anywhere">${esc(f.note)}</td><td style="min-width:260px">${f.preview?`<button class="btn ghost sm" data-action="preview" data-id="${f.id}">预览</button> `:''}${f.local?`<button class="btn ghost sm" data-action="open" data-id="${f.id}">${f.cad?'用 CAD 打开':f.wps?'用 WPS 打开':'本机打开'}</button> `:''}<a class="btn ghost sm" style="display:inline-block;text-decoration:none" href="${esc(url(f))}" download>下载</a> <button class="btn ghost sm" data-action="edit" data-id="${f.id}">修改分类 / 备注</button></td></tr>`).join('')||'<tr><td colspan="7" style="padding:32px">暂无符合条件的资料，请选择分类并上传文件。</td></tr>';
  }
  async function load(){
    const proj=ui.proj,token=++request;
    if(project!==proj){files=[];project=proj;render();}
    try{const result=await api('/api/materials?proj='+encodeURIComponent(proj));if(token!==request||proj!==ui.proj)return;if(!result.ok)throw Error(result.err);files=result.files;project=proj;render();}
    catch(e){if(proj===ui.proj)$m('[data-status]').textContent='读取失败：'+e.message;}
  }
  $m('[data-search]').oninput=render;$m('[data-filter]').onchange=render;$m('[data-refresh]').onclick=load;
  $m('[data-upload]').onclick=async()=>{
    if(uploading)return;
    const selected=[...$m('[data-files]').files],proj=ui.proj,category=$m('[data-category]').value,note=$m('[data-note]').value;
    if(!selected.length){toast('请先选择要上传的文件');return;}
    if(selected.some(f=>!f.size||f.size>100*1024*1024)){toast('文件须大于 0 字节且不超过 100 MB');return;}
    uploading=true;$m('[data-upload]').disabled=true;$m('[data-files]').disabled=true;let successes=0;const failures=[];
    for(const [i,file] of selected.entries()){
      $m('[data-status]').textContent=`正在上传 ${i+1}/${selected.length}：${file.name}`;
      try{const response=await fetch('/api/materials/upload?'+new URLSearchParams({proj,category,name:file.name,note}),{method:'POST',headers:{'Content-Type':'application/octet-stream'},body:file});const result=await response.json();if(!result.ok)throw Error(result.err);successes++;}catch(e){failures.push(file.name+'：'+e.message);}
    }
    uploading=false;$m('[data-upload]').disabled=false;$m('[data-files]').disabled=false;$m('[data-files]').value='';
    $m('[data-status]').textContent=`上传完成：成功 ${successes} 份，失败 ${failures.length} 份。`+failures.join('；')+(ui.proj!==proj?' 已保存至上传时选择的项目。':'');await load();
  };
  const dialog=document.createElement('dialog');dialog.style.cssText='width:1000px;max-width:94vw;max-height:92vh;padding:20px;border:1px solid #ccd5df;border-radius:10px';document.body.appendChild(dialog);
  dialog.addEventListener('close',()=>{dialog.innerHTML='';});
  $m('[data-rows]').onclick=async e=>{
    const button=e.target.closest('button');if(!button)return;
    const f=files.find(f=>f.id===button.dataset.id);if(!f||project!==ui.proj)return;
    const proj=project,operation=button.dataset.action;
    if(operation==='preview'){
      dialog.innerHTML=`<div class="toolbar"><b>${esc(f.name)}</b><span class="spacer"></span><button class="btn ghost" data-close>关闭预览</button></div><iframe title="${esc(f.name)}" src="${esc(url(f,true))}" style="width:100%;height:70vh;border:1px solid #ddd;background:white"></iframe>`;
      dialog.querySelector('[data-close]').onclick=()=>dialog.close();dialog.showModal();return;
    }
    if(operation==='edit'){
      dialog.innerHTML=`<h3>修改资料分类和备注</h3><p style="margin:12px 0">${esc(f.name)}</p><form><label>分类 <select name="category">${categories.map(c=>`<option${f.category===c?' selected':''}>${c}</option>`).join('')}</select></label><label style="display:block;margin:14px 0">备注 <textarea name="note" maxlength="2000">${esc(f.note)}</textarea></label><p data-error role="alert"></p><button class="btn" type="submit">保存</button> <button class="btn ghost" type="button" data-close>取消</button></form>`;
      dialog.querySelector('[data-close]').onclick=()=>dialog.close();dialog.querySelector('form').onsubmit=async event=>{event.preventDefault();const save=event.target.querySelector('[type=submit]');save.disabled=true;try{if(ui.proj!==proj)throw Error('项目已切换，请重新打开资料');const r=await api('/api/materials/action',{proj,id:f.id,action:'edit',...Object.fromEntries(new FormData(event.target))});if(!r.ok)throw Error(r.err);dialog.close();await load();toast(r.msg);}catch(err){dialog.querySelector('[data-error]').textContent=err.message;}finally{save.disabled=false;}};dialog.showModal();return;
    }
    button.disabled=true;try{const result=await api('/api/materials/action',{proj,id:f.id,action:'open'});if(!result.ok)throw Error(result.err);toast(result.msg);}catch(err){$m('[data-status]').textContent=err.message;toast(err.message);}finally{button.disabled=false;}
  };
  const oldRender=renderAll;renderAll=function(){oldRender();if(location.hash==='#materials')load();};
  const oldRoute=route;route=function(){oldRoute();if(location.hash==='#materials'){$('#pageTitle').textContent='资料管理';load();}};
  
  if(location.hash==='#materials')route();
})();
