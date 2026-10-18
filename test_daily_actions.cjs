const fs=require('fs'),vm=require('vm'),assert=require('assert');
const html=fs.readFileSync(__dirname+'/index.html','utf8');
for(const m of html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g))new vm.Script(m[1]);
new vm.Script(fs.readFileSync(__dirname+'/daily.js','utf8'));
const sys=html.slice(html.indexOf('function renderSystems(){'),html.indexOf('function renderSysLed'));
for(const id of ['subIssueAdd','subTodoAdd','subRiskAdd'])assert(!sys.includes(id));
const els={};const ctx={state:{issues:[],todos:[],risks:[],sys_status:{}},ui:{proj:'bg',newIssues:[]},issueDraft:{},todosNew:[],risksNew:[],curSubs:()=>['光纤','监控'],subName:s=>s,isTitan:()=>ctx.ui.proj==='titan',esc:s=>String(s??''),todayStr:()=> '2026-10-01',saveAll:()=>{},toast:()=>{},location:{},encodeURIComponent,editSubsystemAction:()=>{},deleteSubsystemAction:()=>{},$:id=>els[id]||(els[id]={value:'',querySelectorAll:()=>[],innerHTML:''})};
vm.createContext(ctx);vm.runInContext(html.slice(html.indexOf('function recordNumber('),html.indexOf('async function editSubsystemAction')),ctx);vm.runInContext(html.slice(html.indexOf('function renderDailyActions(){'),html.indexOf('function renderSystems(){')),ctx);
for(const proj of ['bg','titan']){
ctx.ui.proj=proj;ctx.state.issues=['光纤','监控',''].map((s,i)=>{const r=Array(21).fill('');r[0]='ISSUE-'+i;r[1]='历史问题'+i;r[proj==='titan'?7:4]=s;r[20]='历史结果';return r;});
ctx.state.todos=['光纤','监控',''].map((s,i)=>{const r=Array(11).fill('');r[0]='TODO-'+i;r[1]='历史待办'+i;r[9]=s;return r;});ctx.state.risks=['光纤','监控',''].map((s,i)=>{const r=Array(19).fill('');r[0]='RISK-'+i;r[1]=s;r[2]='历史风险'+i;r[18]='2026-09-11';return r;});
ctx.$('#dailyActionSub').value='';ctx.renderDailyActions();for(const kind of ['issue','todo','risk'])for(let i=0;i<3;i++)assert(els['#dailyActions'].innerHTML.includes('data-action-edit="'+kind+'|'+kind.toUpperCase()+'-'+i+'"'));
els['#dailyActionSub'].value='光纤';ctx.renderDailyActions();assert(!els['#dailyActions'].innerHTML.includes('ISSUE-1'));assert(els['#dailyActions'].innerHTML.includes('历史结果'));
ctx.$('#subIssueDesc').value='新增问题';els['#subIssueAdd'].onclick();assert.equal(ctx.ui.newIssues.at(-1).sub,'光纤');
ctx.$('#subTodoItem').value='新增待办';els['#subTodoAdd'].onclick();assert.equal(ctx.todosNew.at(-1).sub,'光纤');
ctx.$('#subRiskDesc').value='新增风险';els['#subRiskAdd'].onclick();assert.equal(ctx.risksNew.at(-1).sub,'光纤');
els['#dailyActionSub'].value='';ctx.renderDailyActions();ctx.$('#subIssueDesc').value='项目问题';els['#subIssueAdd'].onclick();assert.equal(ctx.ui.newIssues.at(-1).sub,'项目级');
}
console.log('PASS: syntax, removed subsystem modules, BG/Titan history, filtering, preserved results, creation');
