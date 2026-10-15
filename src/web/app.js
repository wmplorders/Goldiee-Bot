let state={jobs:[],files:[],settings:{},conversations:[]};
let token='',activeId='',selectedUpload='';
const $=s=>document.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const nowTime=()=>new Intl.DateTimeFormat('en-IN',{timeZone:'Asia/Kolkata',hour:'2-digit',minute:'2-digit'}).format(new Date());
const isoDate=d=>new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Kolkata',year:'numeric',month:'2-digit',day:'2-digit'}).format(d);
const today=()=>isoDate(new Date());
const yesterday=()=>{const d=new Date();d.setDate(d.getDate()-1);return isoDate(d)};
const displayDate=s=>{if(!s)return '';const [y,m,d]=s.split('-');return `${d}-${m}-${y}`};
const active=()=>state.conversations.find(c=>c.id===activeId);
const jobById=id=>state.jobs.find(j=>j.id===id);
const link=n=>`<a class="file-link" href="/api/file?name=${encodeURIComponent(n)}">${esc(n)}</a>`;

function notice(message,error=false){const n=$('#notice');n.textContent=message;n.className=error?'error':'';if(message)setTimeout(()=>{if(n.textContent===message)n.textContent=''},5000)}
async function api(path,data){
  const options=data===undefined?{}:{method:'POST',headers:{'Content-Type':'application/json','X-CSRF-Token':token},body:JSON.stringify(data)};
  const response=await fetch(path,options);const value=await response.json();
  if(!response.ok){if(response.status===401&&!$('#login').open)$('#login').showModal();throw Error(value.error||'Request failed')}
  return value;
}
function message(role,text,extra={}){return {role,type:'text',text,time:nowTime(),...extra}}
function assistant(text,extra={}){return message('bot',text,extra)}
function user(text,extra={}){return message('user',text,extra)}

async function createConversation(flow,title,messages=[],context={}){
  const conversation=await api('/api/conversations',{action:'create',flow,title,messages,context});
  state.conversations.unshift(conversation);activeId=conversation.id;renderAll();return conversation;
}
async function saveActive(){const c=active();if(!c)return;const saved=await api('/api/conversations',{action:'save',conversation:c});Object.assign(c,saved);renderHistory()}
async function append(...messages){const c=active();c.messages.push(...messages);await saveActive();renderChat(true)}
function disableLastPrompt(){const c=active();for(let i=c.messages.length-1;i>=0;i--){if(c.messages[i].controls&&!c.messages[i].disabled){c.messages[i].disabled=true;break}}}

function welcome(){return `<div class="welcome"><span class="avatar">G</span><h1>Goldiee Assistant</h1><p>Choose an action below or type a command. Every operation stays in this conversation, including confirmations, progress and generated files.</p><div class="welcome-actions"><button data-start="purchase">Purchase Summary</button><button data-start="unloading">Unloading Slip</button><button data-start="stock">Stock Status</button></div></div>`}
function statusLabel(status){return `<span class="job-status ${esc(status)}">${esc(status)}</span>`}
function friendlyDetail(job){const lines=String(job?.detail||'').trim().split('\n');if(lines[0]?.startsWith('Traceback'))return lines.at(-1).replace(/^\w+(?:Error|Exception):\s*/,'');return job?.detail||'Waiting in queue'}
function controlsHtml(m,index){
  if(!m.controls)return '';
  const disabled=m.disabled?'disabled':'';
  if(m.type==='choice')return `<div class="controls control-buttons">${m.options.map(o=>`<button ${disabled} data-answer="${esc(o.value)}" data-label="${esc(o.label)}">${esc(o.label)}</button>`).join('')}</div>`;
  if(m.type==='date')return `<div class="controls"><div class="control-buttons"><button ${disabled} data-answer="${today()}" data-label="Today">Today</button><button ${disabled} data-answer="${yesterday()}" data-label="Yesterday">Yesterday</button><button ${disabled} data-show-date>Choose date</button></div><form class="date-control" data-date-form hidden><input type="date" value="${today()}" required ${disabled}><button class="primary" ${disabled}>Continue</button></form></div>`;
  if(m.type==='invoices')return `<div class="controls"><div class="invoice-list">${m.invoices.map((x,i)=>`<label class="invoice-option"><input class="invoice-check" type="checkbox" value="${esc(x.invoice)}" data-cases="${Number(x.cases)||0}" ${disabled} checked><span>${esc(x.invoice)}</span><strong>${Number(x.cases).toFixed(2)} cases</strong></label>`).join('')}</div><div class="control-buttons"><button class="primary" data-invoices ${disabled}>Continue</button></div></div>`;
  if(m.type==='text_input')return `<form class="controls text-control" data-text-form><textarea placeholder="${esc(m.placeholder||'Enter details')}" ${disabled} required>${esc(m.value||'')}</textarea><button class="primary" ${disabled}>Continue</button></form>`;
  if(m.type==='upload')return `<div class="controls control-buttons"><button data-upload="${esc(m.purpose)}" ${disabled}>Choose Excel file</button></div>`;
  if(m.type==='confirm')return `<div class="controls"><div class="summary-card">${m.lines.map(x=>`<div class="summary-line"><span>${esc(x.label)}</span><strong>${esc(x.value)}</strong></div>`).join('')}</div><div class="control-buttons"><button class="primary" data-confirm ${disabled}>${esc(m.confirmLabel||'Create')}</button><button data-changes ${disabled}>Make Changes</button><button class="danger" data-cancel ${disabled}>Cancel</button></div></div>`;
  if(m.type==='files')return `<div class="controls"><input class="inline-input file-search" placeholder="Search generated files"><div class="file-list">${fileRows(state.files)}</div></div>`;
  return '';
}
function jobHtml(m){
  const j=jobById(m.jobId)||{status:'queued',detail:'Waiting in queue',files:[]};
  const retry=j.status==='failed'?`<button data-retry="${esc(j.id)}">Retry</button>`:'';
  const again=['completed','failed'].includes(j.status)?`<button data-start="${esc(m.flow)}">Create another</button>`:'';
  return `<div class="sender">Goldiee ${statusLabel(j.status)}</div><div class="bubble-text">${esc(friendlyDetail(j))}</div><div class="job-files">${(j.files||[]).map(link).join('')}</div>${j.diagnostic?`<details><summary>Technical details</summary><pre>${esc(j.diagnostic)}</pre></details>`:''}<div class="controls control-buttons">${retry}${again}</div>`;
}
function renderMessage(m,index){
  if(m.type==='job')return `<div class="message-row bot"><div class="bubble">${jobHtml(m)}<div class="bubble-time">${esc(m.time||'')}</div></div></div>`;
  const sender=m.role==='bot'?'<div class="sender">Goldiee</div>':'';
  return `<div class="message-row ${m.role==='user'?'user':'bot'}"><div class="bubble">${sender}<div class="bubble-text">${esc(m.text||'')}</div>${controlsHtml(m,index)}<div class="bubble-time">${esc(m.time||'')}</div></div></div>`;
}
function renderChat(scroll=false){const chat=$('#chat'),c=active();chat.innerHTML=c?`<div class="date-separator"><span>${new Date(c.created).toLocaleDateString('en-IN',{day:'numeric',month:'short',year:'numeric'})}</span></div>${c.messages.map(renderMessage).join('')}`:welcome();wireDynamic();if(scroll)chat.scrollTop=chat.scrollHeight}
function groupLabel(value){const d=value.slice(0,10);if(d===today())return 'Today';if(d===yesterday())return 'Yesterday';return new Date(value).toLocaleDateString('en-IN',{day:'numeric',month:'short',year:'numeric'})}
function renderHistory(){let last='';$('#history').innerHTML=state.conversations.map(c=>{const group=groupLabel(c.updated||c.created),heading=group!==last?`<div class="history-group">${esc(group)}</div>`:'';last=group;return `${heading}<button class="history-item ${c.id===activeId?'active':''}" data-conversation="${c.id}"><strong>${esc(c.title)}</strong><small>${esc(c.status||'active')}</small></button>`}).join('')||'<div class="empty" style="padding:14px">No conversations yet</div>';document.querySelectorAll('[data-conversation]').forEach(b=>b.onclick=()=>{activeId=b.dataset.conversation;renderAll();closeSidebar()})}
function renderAll(){renderHistory();renderChat(true);$('#company').textContent=state.settings.company||'Not configured';$('#connection').textContent=state.jobs.some(j=>['queued','running'].includes(j.status))?'Working…':'Local operations'}
function fileRows(files){return files.slice(0,60).map(f=>`<div class="file-row">${link(f.name)}<small>${Math.ceil(f.size/1024)} KB</small></div>`).join('')||'<div class="empty">No generated files found.</div>'}

async function refresh(){
  try{
    const current=active(),oldJobs=JSON.stringify(state.jobs.map(j=>[j.id,j.status,j.detail,j.files]));
    const latest=await api('/api/state');token=latest.token;
    if(current)latest.conversations=latest.conversations.map(c=>c.id===current.id?current:c);
    state=latest;
    if(!activeId&&state.conversations.length)activeId=state.conversations[0].id;
    await syncJobConversations();
    renderHistory();
    $('#company').textContent=state.settings.company||'Not configured';
    $('#connection').textContent=state.jobs.some(j=>['queued','running'].includes(j.status))?'Working…':'Local operations';
    const jobsChanged=oldJobs!==JSON.stringify(state.jobs.map(j=>[j.id,j.status,j.detail,j.files]));
    if(!current||jobsChanged)renderChat(jobsChanged);
  }catch(e){notice(e.message,true)}
}
async function syncJobConversations(){
  for(const c of state.conversations){
    const jm=[...c.messages].reverse().find(m=>m.type==='job');if(!jm)continue;
    const j=jobById(jm.jobId);if(!j)continue;
    if(['completed','failed'].includes(j.status)&&c.status!==j.status){c.status=j.status;c.updated=j.finished||c.updated;try{await api('/api/conversations',{action:'save',conversation:c})}catch{}}
  }
}

async function startFlow(flow){
  selectedUpload='';
  const names={purchase:'Purchase Summary',unloading:'Unloading Slip',stock:'Stock Status',dms:'DMS Stock',challans:'Delivery Challans',files:'Generated Files',connections:'Connection Check'};
  const c=await createConversation(flow,names[flow]||'New conversation',[user(names[flow]||flow)]);
  if(flow==='purchase')return append(assistant('Where should I take the invoice from?',{type:'choice',controls:true,options:[{label:'Gmail',value:'gmail'},{label:'Upload Excel',value:'upload'}]}));
  if(flow==='unloading')return askDate('Which purchase invoice date should I use?');
  if(flow==='challans')return append(assistant('Upload the transport details Excel workbook.',{type:'upload',controls:true,purpose:'challans'}));
  if(flow==='files'){c.status='completed';return append(assistant('Here are the latest generated files.',{type:'files',controls:true}))}
  if(flow==='stock')return showConfirmation('stock',[{label:'Action',value:'Fetch detailed stock status from Tally'},{label:'Company',value:state.settings.company||'Not configured'}],'Fetch Stock');
  if(flow==='dms')return showConfirmation('dms',[{label:'Action',value:'Download current DMS stock'},{label:'Company',value:'GOLDIEE GROUP'}],'Fetch DMS Stock');
  if(flow==='connections')return showConfirmation('connections',[{label:'Action',value:'Check Tally and authorization connections'}],'Check Connections');
}
async function askDate(text){await append(assistant(text,{type:'date',controls:true}))}
async function chooseAnswer(value,label){
  const c=active(),prompt=[...c.messages].reverse().find(m=>m.controls&&!m.disabled);if(!prompt)return;
  disableLastPrompt();await append(user(label));
  if(c.flow==='purchase'&&value==='gmail'){c.context.source='gmail';return askDate('Which email received date should I scan?')}
  if(c.flow==='purchase'&&value==='upload')return append(assistant('Upload the purchase invoice Excel workbook.',{type:'upload',controls:true,purpose:'purchase'}));
  if(c.flow==='purchase'&&prompt.type==='date'){c.context.date=value;return showPurchaseConfirmation()}
  if(c.flow==='unloading'&&prompt.type==='date'){c.context.date=value;return previewUnloading(value)}
  if(c.flow==='unloading'&&value==='none'){c.context.dispatch='none';c.context.bills='';return showUnloadingConfirmation()}
  if(c.flow==='unloading'&&value==='bills'){c.context.dispatch='bills';return append(assistant('Enter all undispatched sales bill numbers, separated by commas or new lines.',{type:'text_input',controls:true,placeholder:'26-27/SG-TI-340, 26-27/SG-EX-57'}))}
  if(c.flow==='unloading'&&value.startsWith('change:'))return changeUnloading(value.slice(7));
}
async function previewUnloading(date){
  try{
    await append(assistant('Checking purchase summaries for that date…'));
    const p=await api('/api/preview/unloading',{date});const c=active();c.context.preview=p;
    if(!p.invoices.length)return append(assistant(`I could not find any purchase summaries dated ${displayDate(date)}. Create them first or choose another date.`,{type:'choice',controls:true,options:[{label:'Choose another date',value:'change:date'}]}));
    return append(assistant(`I found ${p.invoices.length} purchase invoice${p.invoices.length===1?'':'s'} totalling ${Number(p.total_cases).toFixed(2)} cases. Select the invoices to include.`,{type:'invoices',controls:true,invoices:p.invoices}));
  }catch(e){notice(e.message,true);await append(assistant(`I could not read the purchase summaries: ${e.message}`))}
}
async function selectInvoices(button){
  const bubble=button.closest('.bubble'),checks=[...bubble.querySelectorAll('.invoice-check:checked')];if(!checks.length)return notice('Select at least one invoice.',true);
  const c=active();c.context.invoices=checks.map(x=>x.value);c.context.cases=checks.reduce((n,x)=>n+Number(x.dataset.cases),0);disableLastPrompt();
  await append(user(`${checks.length} invoice${checks.length===1?'':'s'} selected`),assistant('Are any billed goods still in the godown?',{type:'choice',controls:true,options:[{label:'None',value:'none'},{label:'Enter bill numbers',value:'bills'}]}));
}
async function submitText(form){const value=form.querySelector('textarea').value.trim();if(!value)return;const c=active();disableLastPrompt();if(c.flow==='unloading'){c.context.bills=value;await append(user(value));return showUnloadingConfirmation()}}
async function showPurchaseConfirmation(){const c=active();return showConfirmation('mail',[{label:'Source',value:'Gmail'},{label:'Email received date',value:displayDate(c.context.date)}],'Create Summaries')}
async function showUnloadingConfirmation(){
  const c=active(),x=c.context,lines=[{label:'Purchase date',value:displayDate(x.date)},{label:'Invoices',value:(x.invoices||[]).join(', ')},{label:'Incoming quantity',value:`${Number(x.cases||0).toFixed(2)} cases`},{label:'Undispatched sales bills',value:x.bills||'None'},{label:'Stock source',value:'Fresh Tally stock'},{label:'Company',value:state.settings.company||'Not configured'}];
  return showConfirmation('unloading',lines,'Create Slip');
}
async function showConfirmation(kind,lines,label){const c=active();c.context.pendingKind=kind;return append(assistant('Please review before I start.',{type:'confirm',controls:true,lines,confirmLabel:label}))}
async function changeUnloading(section){
  if(section==='date')return askDate('Choose the purchase invoice date.');
  if(section==='invoices'){const p=active().context.preview;return append(assistant('Select the purchase invoices to include.',{type:'invoices',controls:true,invoices:p?.invoices||[]}))}
  if(section==='bills')return append(assistant('Are any billed goods still in the godown?',{type:'choice',controls:true,options:[{label:'None',value:'none'},{label:'Enter bill numbers',value:'bills'}]}));
}
async function makeChanges(){const c=active();disableLastPrompt();await append(user('Make changes'));if(c.flow==='unloading')return append(assistant('What would you like to change?',{type:'choice',controls:true,options:[{label:'Date',value:'change:date'},{label:'Purchase invoices',value:'change:invoices'},{label:'Undispatched bills',value:'change:bills'}]}));if(c.flow==='purchase')return startFlow('purchase');return append(assistant('Cancel this operation and choose the action again below.'))}
async function cancelFlow(){disableLastPrompt();const c=active();c.status='cancelled';await append(user('Cancel'),assistant('Cancelled. No operation was started.'))}

async function confirmJob(){
  const c=active(),kind=c.context.pendingKind;disableLastPrompt();await append(user(c.messages.findLast(m=>m.type==='confirm')?.confirmLabel||'Create'));
  let request={kind,conversation_id:c.id};
  if(kind==='mail')request.date=c.context.date;
  if(kind==='purchase'||kind==='challans')request.file=c.context.file;
  if(kind==='unloading')request={...request,date:c.context.date,invoices:(c.context.invoices||[]).join(', '),dispatch:c.context.dispatch||'none',bills:c.context.bills||''};
  await enqueue(request,c.flow);
}
async function enqueue(request,flow){
  try{const j=await api('/api/jobs',request),c=active();c.status='processing';c.context.lastRequest=request;await append({role:'bot',type:'job',jobId:j.id,flow,time:nowTime()})}
  catch(e){notice(e.message,true);await append(assistant(`I could not start the operation: ${e.message}`))}
}
async function retryJob(id){const old=jobById(id),c=active();if(!old)return;await append(user('Retry'));await enqueue({...old.request,conversation_id:c.id},c.flow)}

async function uploadFile(purpose){selectedUpload=purpose;$('#hiddenfile').value='';$('#hiddenfile').click()}
async function handleUpload(file){
  if(!file||!selectedUpload)return;
  try{
    notice('Uploading workbook…');const response=await fetch('/api/upload',{method:'POST',headers:{'X-CSRF-Token':token,'X-Filename':encodeURIComponent(file.name)},body:file});const result=await response.json();if(!response.ok)throw Error(result.error);
    const c=active();disableLastPrompt();c.context.file=result.name;c.context.originalFile=file.name;await append(user(`Attached: ${file.name}`));
    const label=selectedUpload==='purchase'?'Create Summary':'Create Challans';await showConfirmation(selectedUpload,[{label:'Workbook',value:file.name},{label:'Action',value:selectedUpload==='purchase'?'Create purchase summary':'Create delivery challans'}],label);
  }catch(e){notice(e.message,true)}finally{selectedUpload=''}
}

function wireDynamic(){
  document.querySelectorAll('[data-start]').forEach(b=>b.onclick=()=>startFlow(b.dataset.start));
  document.querySelectorAll('[data-answer]').forEach(b=>b.onclick=()=>chooseAnswer(b.dataset.answer,b.dataset.label));
  document.querySelectorAll('[data-show-date]').forEach(b=>b.onclick=()=>{b.closest('.controls').querySelector('[data-date-form]').hidden=false;b.hidden=true});
  document.querySelectorAll('[data-date-form]').forEach(f=>f.onsubmit=e=>{e.preventDefault();chooseAnswer(f.querySelector('input').value,displayDate(f.querySelector('input').value))});
  document.querySelectorAll('[data-invoices]').forEach(b=>b.onclick=()=>selectInvoices(b));
  document.querySelectorAll('[data-text-form]').forEach(f=>f.onsubmit=e=>{e.preventDefault();submitText(f)});
  document.querySelectorAll('[data-upload]').forEach(b=>b.onclick=()=>uploadFile(b.dataset.upload));
  document.querySelectorAll('[data-confirm]').forEach(b=>b.onclick=confirmJob);
  document.querySelectorAll('[data-changes]').forEach(b=>b.onclick=makeChanges);
  document.querySelectorAll('[data-cancel]').forEach(b=>b.onclick=cancelFlow);
  document.querySelectorAll('[data-retry]').forEach(b=>b.onclick=()=>retryJob(b.dataset.retry));
  document.querySelectorAll('.file-search').forEach(input=>input.oninput=()=>{const list=input.nextElementSibling,q=input.value.toLowerCase();list.innerHTML=fileRows(state.files.filter(f=>f.name.toLowerCase().includes(q)))})
}

function parseCommand(value){const v=value.trim().toLowerCase();if(v.startsWith('purchase'))return 'purchase';if(v.startsWith('unloading'))return 'unloading';if(v==='stock status'||v.startsWith('tally stock'))return 'stock';if(v.startsWith('dms stock'))return 'dms';if(v.includes('challan'))return 'challans';if(v.includes('file'))return 'files';return ''}
function closeSidebar(){$('#sidebar').classList.remove('open');$('#sidebarshade').classList.remove('open')}

$('#actions').onclick=e=>{const b=e.target.closest('[data-flow]');if(b)startFlow(b.dataset.flow)};
$('#newchat').onclick=()=>{activeId='';renderAll();closeSidebar()};
$('#composer').onsubmit=e=>{e.preventDefault();const input=$('#message'),value=input.value.trim();if(!value)return;input.value='';const flow=parseCommand(value);if(flow)startFlow(flow);else createConversation('chat','Question',[user(value),assistant('I can help with Purchase Summary, Unloading Slip, Stock Status, DMS Stock, Delivery Challans, or generated Files. Choose a button below to begin.')],{}).then(c=>{c.status='completed';saveActive()})};
$('#attach').onclick=()=>{const c=active(),prompt=c&&[...c.messages].reverse().find(m=>m.type==='upload'&&!m.disabled);if(prompt)uploadFile(prompt.purpose);else notice('Choose Purchase Summary or Delivery Challans before attaching an Excel file.',true)};
$('#hiddenfile').onchange=e=>handleUpload(e.target.files[0]);
$('#historytoggle').onclick=()=>{$('#sidebar').classList.add('open');$('#sidebarshade').classList.add('open')};$('#sidebarshade').onclick=closeSidebar;
$('#settingsbutton').onclick=()=>{const d=$('#settingsdialog');d.querySelector('[name=company]').value=state.settings.company||'';d.showModal()};
document.querySelectorAll('[data-close]').forEach(b=>b.onclick=()=>$('#'+b.dataset.close).close());
$('#settingsform').onsubmit=async e=>{e.preventDefault();try{await api('/api/settings',Object.fromEntries(new FormData(e.target)));$('#settingsdialog').close();await refresh();notice('Working company saved.')}catch(x){notice(x.message,true)}};
$('#loadcompanies').onclick=async()=>{try{const r=await api('/api/companies');$('#companies').innerHTML=r.companies.map(c=>`<option value="${esc(c)}">`).join('');if(r.companies.length===1)$('#settingsform [name=company]').value=r.companies[0];notice(`${r.companies.length} loaded companies found.`)}catch(e){notice(e.message,true)}};
$('#checkconnections').onclick=()=>{$('#settingsdialog').close();startFlow('connections')};
$('#loginform').onsubmit=async e=>{e.preventDefault();try{await api('/api/login',Object.fromEntries(new FormData(e.target)));$('#login').close();await refresh()}catch(x){$('#loginerror').textContent=x.message}};

renderAll();refresh();setInterval(refresh,4000);
