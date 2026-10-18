const fs=require('fs'),vm=require('vm'),assert=require('assert');
const html=fs.readFileSync(__dirname+'/index.html','utf8');
const elements={};
const ctx={orgDraft:{people:[{pid:'1',name:'历史',status:'已删除'},{pid:'2',name:'甲',status:'在职'},{pid:'3',name:'乙',status:'在职'}],groups:[],jobs:[]},state:{people:[]},ORG_ROLES:[],ORG_STATUS:[],esc:s=>String(s),document:{getElementById:()=>({})},toast:()=>{},renderOrgStats:()=>{},renderOrgTree:()=>{},renderOrgGroupsJobs:()=>{},confirm:()=>true,deepCopy:o=>JSON.parse(JSON.stringify(o)),saved:[],
  $:id=>elements[id]||(elements[id]={dataset:{},addEventListener(){}})};
ctx.saveOrg=async()=>ctx.saved.push(JSON.parse(JSON.stringify(ctx.orgDraft)));
vm.createContext(ctx);
for(const [start,end] of [['function renderOrgTable(){','function renderOrgGroupsJobs(){'],['function bindOrgInputs(){','function renderOrg(){'],['async function clearOrgPeople(){','/* ===== 甲方交付文件中心']]){
  vm.runInContext(html.slice(html.indexOf(start),html.indexOf(end,html.indexOf(start))),ctx);
}
const click=(attr,index)=>elements['#orgTable'].onclick({target:{closest:selector=>selector==='['+attr+']'?{dataset:{[attr.replace(/^data-/,'').replace(/-([a-z])/g,(_,c)=>c.toUpperCase())]:String(index)}}:null}});
(async()=>{
  ctx.renderOrgTable();ctx.bindOrgInputs();
  assert(elements['#orgTable'].innerHTML.includes('data-del="1"'));
  assert(!elements['#orgTable'].innerHTML.includes('data-del="0"'));
  elements['#orgTable'].oninput({target:{classList:{contains:()=>true},dataset:{i:'2',f:'phone'},value:'123'}});
  assert.equal(ctx.orgDraft.people[2].phone,'123');
  await click('data-del',1);await click('data-confirm-del',1);
  assert.equal(ctx.saved.length,1);assert.equal(ctx.orgDraft.people[1].status,'已删除');
  assert.equal(ctx.orgDraft.people[2].status,'在职');
  assert(elements['#orgTable'].innerHTML.includes('data-del="2"'));
  await click('data-del',2);await click('data-confirm-del',2);
  assert.equal(ctx.saved.length,2);assert(ctx.orgDraft.people.every(p=>p.status==='已删除'));
  ctx.orgDraft.people.push({pid:'4',name:'丙',status:'在职'},{pid:'5',name:'丁',status:'停用'});
  ctx.confirm=()=>false;await ctx.clearOrgPeople();assert.equal(ctx.saved.length,2);
  ctx.confirm=()=>true;await ctx.clearOrgPeople();assert.equal(ctx.saved.length,3);
  assert(ctx.orgDraft.people.every(p=>p.status==='已删除'));
  await ctx.clearOrgPeople();assert.equal(ctx.saved.length,3);
  ctx.orgDraft._showDel=true;ctx.renderOrgTable();assert(elements['#orgTable'].innerHTML.includes('历史'));
  console.log('PASS filtered editing, consecutive deletions, automatic save, clear cancellation, clear all, empty list, history');
})().catch(e=>{console.error(e);process.exitCode=1});
