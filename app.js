(function(){
"use strict";
/* =================== ρυθμίσεις =================== */
function lsGet(k){ try{ return localStorage.getItem(k)||""; }catch(e){ return ""; } }
function lsSet(k,v){ try{ localStorage.setItem(k,v); }catch(e){} }
const CFG = Object.assign({prospectsList:"CRM_Prospects", callsList:"CRM_Calls", auditsList:"CRM_Audits"}, window.CRM_CONFIG||{});
CFG.tenantId = CFG.tenantId || lsGet("crm_tenant");
CFG.clientId = CFG.clientId || lsGet("crm_client");
CFG.siteUrl  = CFG.siteUrl  || lsGet("crm_site");
const SCOPES = ["User.Read","Sites.ReadWrite.All"];
const MANAGE_SCOPES = ["Sites.Manage.All"]; // ζητείται μόνο για τη δημιουργία λιστών/στηλών
const CAL_SCOPES = ["Calendars.ReadWrite"]; // υπενθυμίσεις στο ημερολόγιο Outlook του χρήστη
const APP_URL = location.origin+location.pathname;
const GRAPH = "https://graph.microsoft.com/v1.0";

const STATUSES = [
  {k:"new",            l:"Νέο",                    key:"",  queue:true,  final:false, color:"var(--muted)"},
  {k:"no_answer",      l:"Δεν απάντησε",           key:"1", queue:true,  final:false, color:"var(--warn)"},
  {k:"busy",           l:"Κατειλημμένο",           key:"2", queue:true,  final:false, color:"var(--warn)"},
  {k:"callback",       l:"Να ξανακαλέσουμε",       key:"3", queue:true,  final:false, color:"var(--info)"},
  {k:"interested",     l:"Ενδιαφέρεται",           key:"4", queue:false, final:false, color:"var(--good)"},
  {k:"send_info",      l:"Θέλει email/πληροφορίες",key:"5", queue:false, final:false, color:"var(--good)"},
  {k:"appointment",    l:"Κλείστηκε ραντεβού",     key:"6", queue:false, final:true,  color:"var(--good)"},
  {k:"not_interested", l:"Δεν ενδιαφέρεται",       key:"7", queue:false, final:true,  color:"var(--bad)"},
  {k:"wrong_number",   l:"Λάθος/ανενεργό νούμερο", key:"8", queue:false, final:true,  color:"var(--bad)"},
  {k:"do_not_call",    l:"Να μην ξανακληθεί",      key:"9", queue:false, final:true,  color:"var(--bad)"},
];
const S = Object.fromEntries(STATUSES.map(s=>[s.k,s]));
const ANSWERED = new Set(["callback","interested","send_info","appointment","not_interested","do_not_call"]);
const LEADS = new Set(["interested","send_info","appointment"]);
const DEFAULT_CATS = ["Λογιστικό γραφείο","Δικηγορικό γραφείο","Συμβολαιογραφείο","Διαγνωστικό κέντρο / Ιατρείο","Οδοντιατρείο","Ασφαλιστικό γραφείο","Μεσιτικό γραφείο","Ξενοδοχείο","Ναυτιλιακή / Μεταφορική","Αρχιτεκτονικό / Τεχνικό γραφείο","Φροντιστήριο","Χονδρεμπόριο / Αποθήκη","Άλλο"];
const RETRY_MS = 2*3600e3, MAX_ATTEMPTS = 5, POLL_MS = 30000, LOCK_MS = 10*60e3;

/* SharePoint στήλες: κλειδί εφαρμογής -> εσωτερικό όνομα στήλης */
const PF = {n:"Title",cat:"Category",a:"Area",ad:"Address",f:"LegalForm",p:"ContactPerson",ph:"Phone",phk:"PhoneKey",em:"Email",w:"Website",src:"Source",
  status:"Status",attempts:"Attempts",lastBy:"LastBy",notes:"Notes",hidden:"Hidden",afm:"AFM",gemi:"GEMI",turnover:"Turnover",employees:"Employees",
  finYear:"FinYear",finNotes:"FinNotes",lockedBy:"LockedBy",lastCallAt:"LastCallAt",callbackAt:"CallbackAt",lockedUntil:"LockedUntil",calEvent:"CalEventId"};
const PDATES = new Set(["lastCallAt","callbackAt","lockedUntil"]);
const PNUM = new Set(["attempts","turnover","employees","finYear"]);
const CF = {name:"Title",pid:"ProspectId",s:"Status",n:"Note",cb:"CallbackAt",by:"CallerName",byEmail:"CallerEmail",t:"CallAt",manual:"Manual"};

const PROSPECT_COLUMNS = [
  ["Category","text"],["Area","text"],["Address","text"],["LegalForm","text"],["ContactPerson","text"],["Phone","text"],["PhoneKey","text",true],
  ["Email","text"],["Website","text"],["Source","text"],["Status","text",true],["Attempts","number"],["LastCallAt","dateTime"],["LastBy","text"],
  ["CallbackAt","dateTime"],["Notes","note"],["Hidden","boolean"],["AFM","text"],["GEMI","text"],["Turnover","number"],["Employees","number"],
  ["FinYear","number"],["FinNotes","note"],["LockedBy","text"],["LockedUntil","dateTime"],["CalEventId","text"]];
const CALL_COLUMNS = [["ProspectId","number",true],["Status","text"],["Note","note"],["CallbackAt","dateTime"],["CallerName","text"],["CallerEmail","text"],["CallAt","dateTime",true],["Manual","boolean"]];

/* =================== helpers =================== */
const $ = id => document.getElementById(id);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmtDT = t => t ? new Date(t).toLocaleString("el-GR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"}) : "—";
const fmtD = t => new Date(t).toLocaleDateString("el-GR",{day:"2-digit",month:"2-digit"});
const pill = k => `<span class="pill s-${esc(k)}">${esc((S[k]||S.new).l)}</span>`;
const dayKey = t => { const d=new Date(t); return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0"); };
const norm = s => String(s||"").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"");
const normKey = s => norm(s).replace(/[^0-9a-zα-ω]/g,"");
function phoneKey(p){ let d=String(p||"").replace(/\D/g,""); if(d.startsWith("0030")) d=d.slice(4); else if(d.startsWith("30")&&d.length===12) d=d.slice(2); return d.slice(-10); }
function dedupKey(o){ const p=phoneKey(o.ph); return p.length>=10 ? "ph:"+p : "n:"+normKey(o.n)+"|"+normKey(o.a); }
const toMs = v => v ? new Date(v).getTime() : null;
const toIso = v => v ? new Date(v).toISOString() : null;
const fmtMoney = v => v==null||v==="" ? "—" : Number(v).toLocaleString("el-GR",{maximumFractionDigits:0})+" €";
function toast(msg){ const t=$("toast"); t.textContent=msg; t.hidden=false; clearTimeout(toast._t); toast._t=setTimeout(()=>t.hidden=true,2800); }
function banner(msg){ const b=$("banner"); b.textContent=msg||""; b.hidden=!msg; }
const sleep = ms => new Promise(r=>setTimeout(r,ms));

/* =================== auth =================== */
let pca=null, account=null, me={name:"",email:""};
async function getToken(extra){
  const req={scopes:extra||SCOPES, account};
  try{ return (await pca.acquireTokenSilent(req)).accessToken; }
  catch(e){ await pca.acquireTokenRedirect(req); throw e; }
}
async function tokenSilent(scopes){ try{ return (await pca.acquireTokenSilent({scopes,account})).accessToken; }catch(e){ return null; } }
$("btn-login").onclick=()=>pca.loginRedirect({scopes:SCOPES, prompt:"select_account"});
$("btn-logout").onclick=()=>pca.logoutRedirect({account});
$("setupform").addEventListener("submit",e=>{ e.preventDefault();
  lsSet("crm_tenant",$("cf-tenant").value.trim()); lsSet("crm_client",$("cf-client").value.trim()); lsSet("crm_site",$("cf-site").value.trim().replace(/\/+$/,"")); location.reload(); });

/* =================== Graph =================== */
async function g(path, opt={}){
  const url = path.startsWith("http") ? path : GRAPH+path;
  for(let attempt=0; attempt<5; attempt++){
    const tok = opt.silent ? await tokenSilent(opt.scopes) : await getToken(opt.scopes);
    if(!tok){ const e=new Error("Χωρίς δικαίωμα"); e.status=401; throw e; }
    const headers = Object.assign({"Authorization":"Bearer "+tok}, opt.headers||{});
    if(opt.body!==undefined && !(opt.body instanceof Blob)) headers["Content-Type"]="application/json";
    let r;
    try{ r = await fetch(url,{method:opt.method||"GET", headers, body: opt.body!==undefined ? JSON.stringify(opt.body) : undefined}); }
    catch(e){ banner("Δεν υπάρχει σύνδεση με το Microsoft 365. Ελέγξτε το δίκτυο."); throw e; }
    if(r.status===429 || r.status===503){ const ra=Number(r.headers.get("Retry-After"))||2; await sleep(ra*1000); continue; }
    if(r.status===204) return null;
    const data = await r.json().catch(()=>({}));
    if(!r.ok){ const e=new Error((data.error&&data.error.message)||("Σφάλμα "+r.status)); e.status=r.status; e.data=data; throw e; }
    return data;
  }
  throw new Error("Το Microsoft 365 δεν απαντά (πολλά αιτήματα). Δοκίμασε σε λίγο.");
}
async function gAll(path){ let out=[], url=path; while(url){ const d=await g(url, {headers:{"Prefer":"HonorNonIndexedQueriesWarningMayFailRandomly"}}); out=out.concat(d.value||[]); url=d["@odata.nextLink"]; } return out; }
async function gBatch(reqs){ // έως 20 ανά αίτημα
  const results=[];
  for(let i=0;i<reqs.length;i+=20){
    const chunk=reqs.slice(i,i+20).map((r,j)=>({id:String(j),method:r.method,url:r.url,headers:r.body?{"Content-Type":"application/json"}:undefined,body:r.body}));
    let pending=chunk;
    for(let tries=0; pending.length && tries<5; tries++){
      const d=await g("/$batch",{method:"POST",body:{requests:pending}});
      const retry=[]; let wait=0;
      for(const resp of d.responses){
        const idx=i+Number(resp.id);
        if(resp.status===429||resp.status===503){ retry.push(pending.find(p=>p.id===resp.id)); wait=Math.max(wait,Number((resp.headers||{})["Retry-After"])||3); }
        else results[idx]=resp;
      }
      pending=retry; if(retry.length) await sleep(wait*1000);
    }
  }
  return results;
}

let site=null, lists={p:null,c:null};
async function resolveSite(){
  const u=new URL(CFG.siteUrl);
  const path=u.pathname.replace(/\/+$/,"");
  site = await g(path && path!=="/" ? `/sites/${u.hostname}:${path}` : `/sites/${u.hostname}`);
}
async function findList(name){ const d=await g(`/sites/${site.id}/lists?$filter=displayName eq '${name}'&$select=id,displayName`); return (d.value||[])[0]||null; }
function colDef([name,type,indexed]){
  const c={name, indexed:!!indexed};
  if(type==="text") c.text={}; else if(type==="note") c.text={allowMultipleLines:true,linesForEditing:6};
  else if(type==="number") c.number={}; else if(type==="dateTime") c.dateTime={format:"dateTime"}; else if(type==="boolean") c.boolean={};
  return c;
}
async function createLists(){
  const msg=$("setupmsg");
  if(!lists.p){ msg.textContent="Δημιουργία "+CFG.prospectsList+"…"; lists.p=await g(`/sites/${site.id}/lists`,{method:"POST",scopes:MANAGE_SCOPES,body:{displayName:CFG.prospectsList,list:{template:"genericList"},columns:PROSPECT_COLUMNS.map(colDef)}}); }
  if(!lists.c){ msg.textContent="Δημιουργία "+CFG.callsList+"…"; lists.c=await g(`/sites/${site.id}/lists`,{method:"POST",scopes:MANAGE_SCOPES,body:{displayName:CFG.callsList,list:{template:"genericList"},columns:CALL_COLUMNS.map(colDef)}}); }
  msg.textContent="Έτοιμο.";
}

async function ensureColumns(){ // προσθήκη νέων στηλών σε λίστες που φτιάχτηκαν με παλιότερη έκδοση
  try{
    const d=await g(`/sites/${site.id}/lists/${lists.p.id}/columns?$select=name`); const have=new Set((d.value||[]).map(c=>c.name));
    for(const col of PROSPECT_COLUMNS){ if(!have.has(col[0])) await g(`/sites/${site.id}/lists/${lists.p.id}/columns`,{method:"POST",scopes:MANAGE_SCOPES,silent:true,body:colDef(col)}); }
  }catch(e){}
}

/* μετατροπές */
function fromP(it){
  const f=it.fields||{}, o={id:String(it.id), etag:it.eTag};
  for(const [k,col] of Object.entries(PF)){ let v=f[col]; if(PDATES.has(k)) v=toMs(v); else if(PNUM.has(k)) v=(v==null||v==="")?null:Number(v); else if(k==="hidden") v=!!v; else v=v??""; o[k]=v; }
  o.status=o.status||"new"; o.attempts=o.attempts||0; o.history=[];
  return o;
}
function toPFields(o){
  const f={};
  for(const [k,col] of Object.entries(PF)){ if(!(k in o)) continue; let v=o[k];
    if(PDATES.has(k)) v=toIso(v); else if(PNUM.has(k)) v=(v===""||v==null||isNaN(Number(v)))?null:Number(v); else if(k==="hidden") v=!!v; else v=String(v??"").slice(0, k==="notes"||k==="finNotes"?60000:255);
    f[col]=v; }
  if("ph" in o) f.PhoneKey=phoneKey(o.ph);
  return f;
}
function fromC(it){ const f=it.fields||{}; return {id:String(it.id), pid:String(f.ProspectId??""), s:f.Status||"", n:f.Note||"", cb:toMs(f.CallbackAt), by:f.CallerName||"", byEmail:f.CallerEmail||"", t:toMs(f.CallAt), manual:!!f.Manual}; }

const offices=new Map(); let calls=[];
async function load(){
  const [pi, ci] = await Promise.all([
    gAll(`/sites/${site.id}/lists/${lists.p.id}/items?$expand=fields&$top=999`),
    gAll(`/sites/${site.id}/lists/${lists.c.id}/items?$expand=fields&$top=999`)]);
  offices.clear(); pi.forEach(it=>{ const o=fromP(it); offices.set(o.id,o); });
  calls=ci.map(fromC).sort((a,b)=>(a.t||0)-(b.t||0));
  for(const c of calls){ const o=offices.get(c.pid); if(o) o.history.push(c); }
}
async function patchP(id, data){ await g(`/sites/${site.id}/lists/${lists.p.id}/items/${id}/fields`,{method:"PATCH",body:toPFields(data)}); const o=offices.get(id); if(o) Object.assign(o,data); return o; }
async function createP(data){ const it=await g(`/sites/${site.id}/lists/${lists.p.id}/items`,{method:"POST",body:{fields:toPFields(data)}}); const o=fromP(it); Object.assign(o,data,{id:String(it.id)}); o.history=[]; offices.set(o.id,o); return o; }
async function addCall(o, c){
  const fields={Title:o.n.slice(0,255),ProspectId:Number(o.id),Status:c.s,Note:c.n||"",CallbackAt:toIso(c.cb),CallerName:me.name,CallerEmail:me.email,CallAt:toIso(c.t||Date.now()),Manual:!!c.manual};
  const it=await g(`/sites/${site.id}/lists/${lists.c.id}/items`,{method:"POST",body:{fields}});
  const cc=fromC(it); calls.push(cc); o.history.push(cc); return cc;
}

/* =================== προβολές =================== */
const VIEWS=["calls","calendar","list","audit","reports","import"];
function show(v){
  if(!VIEWS.includes(v)) v="calls";
  VIEWS.forEach(x=>{ $("view-"+x).hidden=x!==v; $("tab-"+x).setAttribute("aria-selected",x===v); });
  try{ history.replaceState(null,"","#"+v); }catch(e){}
  if(v==="list") renderList(); if(v==="reports") renderReports(); if(v==="calendar") renderCalendar(); if(v==="audit") renderAuditView();
  if(v==="calls" && !currentId && pca && site && lists.p) nextCall();
}
document.querySelectorAll("nav.tabs button").forEach(b=>b.addEventListener("click",()=>show(b.dataset.view)));

/* ουρά */
let currentId=null, pickedOutcome=null, skipped=new Set();
function inQueue(o,now){
  if(o.hidden) return false; const s=S[o.status]||S.new; if(!s.queue) return false;
  if(o.lockedBy && o.lockedBy!==me.email && (o.lockedUntil||0)>now) return false;
  if(o.status==="callback") return (o.callbackAt||0)<=now;
  if(o.status==="no_answer"||o.status==="busy"){ if((o.attempts||0)>=MAX_ATTEMPTS) return false; return (o.lastCallAt||0)+RETRY_MS<=now; }
  return true;
}
function queueList(){
  const now=Date.now(), area=$("qarea").value, cat=$("qcat").value;
  const arr=[...offices.values()].filter(o=>inQueue(o,now)&&(!area||o.a===area)&&(!cat||o.cat===cat)&&o.ph);
  const rank=o=>o.status==="callback"?0:o.status==="new"?1:2;
  arr.sort((a,b)=>{const r=rank(a)-rank(b); if(r) return r;
    if(a.status==="callback") return (a.callbackAt||0)-(b.callbackAt||0);
    if(a.status==="new") return Number(a.id)-Number(b.id);
    return (a.lastCallAt||0)-(b.lastCallAt||0);});
  return arr;
}
async function release(id){ const o=offices.get(id); if(o && o.lockedBy===me.email){ try{ await patchP(id,{lockedBy:"",lockedUntil:null}); }catch(e){} } }
async function claim(o){
  try{
    const it=await g(`/sites/${site.id}/lists/${lists.p.id}/items/${o.id}?$expand=fields($select=LockedBy,LockedUntil)`);
    const lb=it.fields.LockedBy||"", lu=toMs(it.fields.LockedUntil)||0;
    if(lb && lb!==me.email && lu>Date.now()) { o.lockedBy=lb; o.lockedUntil=lu; return false; }
    await g(`/sites/${site.id}/lists/${lists.p.id}/items/${o.id}/fields`,{method:"PATCH",headers:{"If-Match":it.eTag},body:{LockedBy:me.email,LockedUntil:toIso(Date.now()+LOCK_MS)}});
    o.lockedBy=me.email; o.lockedUntil=Date.now()+LOCK_MS; return true;
  }catch(e){ return e.status===412 ? false : true; }
}
async function nextCall(){
  const prev=currentId; if(prev) release(prev);
  const q=queueList().filter(o=>!skipped.has(o.id)&&o.id!==prev);
  for(const o of q.slice(0,10)){ if(await claim(o)){ openCall(o.id); return; } }
  currentId=null; renderCallCard();
}
$("btn-next").onclick=()=>{ if(currentId) skipped.add(currentId); nextCall(); };
$("qarea").onchange=$("qcat").onchange=()=>{ skipped.clear(); renderQueueCount(); nextCall(); };
function renderQueueCount(){
  const q=queueList(), now=Date.now();
  const due=q.filter(o=>o.status==="callback").length;
  const later=[...offices.values()].filter(o=>!o.hidden&&o.status==="callback"&&(o.callbackAt||0)>now).length;
  $("qcount").innerHTML=`Στην ουρά: <b class="mono">${q.length}</b>${due?` · <b class="mono">${due}</b> επανακλήσεις τώρα`:""}${later?` · ${later} προγραμματισμένες`:""}`;
}
function openCall(id){ currentId=id; pickedOutcome=null; renderCallCard(); }

function finFacts(o){
  if(!o.afm&&!o.gemi&&o.turnover==null&&o.employees==null&&!o.finNotes) return "";
  return `<dt>ΑΦΜ / ΓΕΜΗ</dt><dd>${esc(o.afm)||"—"} / ${esc(o.gemi)||"—"}</dd>
    <dt>Κύκλος εργασιών</dt><dd>${fmtMoney(o.turnover)}${o.finYear?` (${esc(o.finYear)})`:""}</dd>
    <dt>Εργαζόμενοι</dt><dd>${o.employees??"—"}</dd>${o.finNotes?`<dt>Οικονομικά</dt><dd>${esc(o.finNotes)}</dd>`:""}`;
}
function contactFacts(o){
  return `<dl class="facts">
    <dt>Κατηγορία</dt><dd>${esc(o.cat)||"—"}</dd>
    <dt>Υπεύθυνος</dt><dd>${esc(o.p)||'<span class="muted">—</span>'}</dd>
    <dt>Email</dt><dd>${o.em?`<a href="mailto:${esc(o.em)}">${esc(o.em)}</a>`:'<span class="muted">—</span>'}</dd>
    <dt>Διεύθυνση</dt><dd>${esc(o.ad)||"—"}</dd>
    <dt>Νομική μορφή</dt><dd>${esc(o.f)||'<span class="muted">Δεν αναφέρεται</span>'}</dd>
    <dt>Website</dt><dd>${o.w?`<a href="https://${esc(o.w.replace(/^https?:\/\//,""))}" target="_blank" rel="noopener">${esc(o.w)}</a>`:'<span class="muted">—</span>'}</dd>
    ${finFacts(o)}
    ${o.src?`<dt>Πηγή</dt><dd class="muted">${esc(o.src)}</dd>`:""}
  </dl>`;
}
function renderCallCard(){
  const card=$("callcard"); renderQueueCount();
  const o=currentId&&offices.get(currentId);
  if(!o){
    card.innerHTML=offices.size?`<div class="empty"><b>Δεν υπάρχουν κλήσεις στην ουρά</b>Όλοι οι υποψήφιοι με αυτά τα φίλτρα έχουν καλεστεί ή περιμένουν επανάκληση. Οι αναπάντητες ξαναμπαίνουν 2 ώρες μετά.</div>`
      :`<div class="empty"><b>Η λίστα είναι άδεια</b>Φόρτωσε υποψηφίους από την καρτέλα «Εισαγωγή».</div>`;
    $("callhist").innerHTML=`<li class="muted" style="border:0">—</li>`; return;
  }
  const tel=(o.ph||"").replace(/\s/g,"");
  card.innerHTML=`
    <div class="callhead"><div style="min-width:0"><h1>${esc(o.n)}</h1><div class="meta">${esc(o.cat)}${o.cat&&o.a?" · ":""}${esc(o.a)} · κλήσεις: <span class="mono">${o.attempts||0}</span>${o.lastCallAt?` · τελευταία ${fmtDT(o.lastCallAt)}`:""}</div></div>${pill(o.status)}</div>
    <div class="phone"><span class="num">${esc(o.ph)||"—"}</span><button class="btn small" id="cp-ph">Αντιγραφή</button>${tel?`<a class="btn small" href="tel:${esc(tel)}">Κλήση</a>`:""}</div>
    ${o.status==="callback"&&o.callbackAt?`<div class="pill s-callback" style="margin-top:4px">Επανάκληση: ${fmtDT(o.callbackAt)}</div>`:""}
    ${o.notes?`<p style="margin:12px 0 0;white-space:pre-wrap"><span class="muted">Σημειώσεις:</span> ${esc(o.notes)}</p>`:""}
    ${contactFacts(o)}
    <h2>Αποτέλεσμα κλήσης</h2>
    <div class="outcomes" id="outs">${STATUSES.filter(s=>s.key).map(s=>`<button data-k="${s.k}" aria-pressed="false"><span class="sw" style="background:${s.color}"></span>${esc(s.l)}<span class="k" style="margin-left:auto">${s.key}</span></button>`).join("")}</div>
    <div id="cbwrap" class="cbwrap" hidden>
      <label class="f"><span id="cblabel">Ημερομηνία & ώρα επανάκλησης</span><input type="datetime-local" id="cbat"></label>
      <div class="quick" id="cbquick">${QUICK.map((q,i)=>`<button type="button" class="btn small" data-q="${i}">${esc(q.l)}</button>`).join("")}</div>
      <div class="cbwhen muted" id="cbwhen"></div>
      <label class="chk"><input type="checkbox" id="cbcal"${calOn()?" checked":""}> Υπενθύμιση στο ημερολόγιο Outlook μου</label>
    </div>
    <div class="row" style="margin-top:12px">
      <label class="f">Υπεύθυνος (αν τον μάθατε)<input id="ed-p" value="${esc(o.p)}"></label>
      <label class="f">Email (αν το δώσανε)<input id="ed-em" type="email" value="${esc(o.em)}"></label>
    </div>
    <label class="f" style="margin-top:12px">Σημείωση κλήσης<textarea id="note" placeholder="Τι ειπώθηκε, με ποιον μιλήσατε, τι ζήτησαν…"></textarea></label>
    <div class="actions">
      <button class="btn primary" id="save-next" disabled>Αποθήκευση & επόμενο</button>
      <button class="btn" id="save-stay" disabled>Αποθήκευση</button>
      <button class="btn" id="edit-one">Στοιχεία / οικονομικά</button>
      <button class="btn" id="hide-one">Απόκρυψη</button>
    </div>
    <p class="muted savehint" id="savehint">Διάλεξε πρώτα αποτέλεσμα κλήσης (ή πάτα 1–9) για να ενεργοποιηθεί η αποθήκευση.</p>`;
  $("cp-ph").onclick=()=>copy(o.ph);
  card.querySelectorAll("#outs button").forEach(b=>b.onclick=()=>pick(b.dataset.k));
  card.querySelectorAll("#cbquick [data-q]").forEach(b=>b.onclick=()=>{ $("cbat").value=toLocalInput(QUICK[Number(b.dataset.q)].f().getTime()); cbWhen(); });
  $("cbat").addEventListener("input",cbWhen);
  $("cbcal").onchange=e=>{ if(e.target.checked&&!calReady){ e.target.checked=false; connectCalendar(); } else lsSet("crm_cal",e.target.checked?"1":"0"); };
  $("save-next").onclick=()=>saveCall(true); $("save-stay").onclick=()=>saveCall(false);
  $("edit-one").onclick=()=>openDrawer(o.id);
  $("hide-one").onclick=async()=>{ try{ await patchP(o.id,{hidden:true,lockedBy:"",lockedUntil:null}); toast("Κρύφτηκε"); nextCall(); }catch(e){ toast(e.message); } };
  renderHistory(o);
}
const DATED = new Set(["callback","appointment"]);
const atHour=(d,h)=>{ d.setHours(h,0,0,0); return d; };
const QUICK=[
  {l:"Σε 1 ώρα", f:()=>{ const d=new Date(Date.now()+3600e3); d.setMinutes(Math.ceil(d.getMinutes()/15)*15,0,0); return d; }},
  {l:"Αύριο 10:00", f:()=>{ const d=new Date(); d.setDate(d.getDate()+1); return atHour(d,10); }},
  {l:"Αύριο 13:00", f:()=>{ const d=new Date(); d.setDate(d.getDate()+1); return atHour(d,13); }},
  {l:"Σε 2 ημέρες", f:()=>{ const d=new Date(); d.setDate(d.getDate()+2); return atHour(d,10); }},
  {l:"Επόμενη Δευτέρα", f:()=>{ const d=new Date(); d.setDate(d.getDate()+((8-d.getDay())%7||7)); return atHour(d,10); }},
  {l:"Σε 1 εβδομάδα", f:()=>{ const d=new Date(); d.setDate(d.getDate()+7); return atHour(d,10); }},
  {l:"Σε 1 μήνα", f:()=>{ const d=new Date(); d.setMonth(d.getMonth()+1); return atHour(d,10); }},
];
function skipWeekend(d){ while(d.getDay()===0||d.getDay()===6) d.setDate(d.getDate()+1); return d; }
const toLocalInput = t => new Date(t-new Date(t).getTimezoneOffset()*6e4).toISOString().slice(0,16);
function cbWhen(){
  const v=$("cbat").value; const el=$("cbwhen"); if(!el) return;
  if(!v){ el.textContent=""; return; }
  const t=new Date(v); el.textContent=t.toLocaleString("el-GR",{weekday:"long",day:"numeric",month:"long",hour:"2-digit",minute:"2-digit"})+(t.getTime()<Date.now()?" — ήδη πέρασε":"");
}
function pick(k){
  pickedOutcome=k;
  document.querySelectorAll("#outs button").forEach(b=>b.setAttribute("aria-pressed",b.dataset.k===k));
  $("cbwrap").hidden=!DATED.has(k);
  if(DATED.has(k)){
    $("cblabel").textContent=k==="appointment"?"Ημερομηνία & ώρα ραντεβού":"Ημερομηνία & ώρα επανάκλησης";
    if(!$("cbat").value) $("cbat").value=toLocalInput(skipWeekend(QUICK[1].f()).getTime());
    cbWhen();
  }
  $("save-next").disabled=false; $("save-stay").disabled=false; $("savehint").hidden=true;
}
document.addEventListener("keydown",e=>{
  if($("app").hidden||$("view-calls").hidden||!currentId||!$("drawer").hidden) return;
  if(/INPUT|TEXTAREA|SELECT/.test(document.activeElement?.tagName||"")) return;
  const s=STATUSES.find(x=>x.key===e.key); if(s){ pick(s.k); e.preventDefault(); }
});
async function saveCall(goNext){
  const id=currentId, o=offices.get(id); if(!o||!pickedOutcome) return;
  const now=Date.now(); const note=$("note").value.trim();
  let cb=null; if(DATED.has(pickedOutcome)){ const v=$("cbat").value; cb=v?new Date(v).getTime():now+864e5; }
  const wantCal=DATED.has(pickedOutcome)&&$("cbcal")&&$("cbcal").checked;
  $("save-next").disabled=true; $("save-stay").disabled=true;
  try{
    await addCall(o,{s:pickedOutcome,n:note,cb,t:now});
    const patch={status:pickedOutcome,attempts:(o.attempts||0)+1,lastCallAt:now,lastBy:me.name,callbackAt:cb,p:$("ed-p").value.trim(),em:$("ed-em").value.trim(),lockedBy:"",lockedUntil:null};
    if(note) patch.notes = (o.notes?o.notes+"\n":"")+`[${fmtDT(now)} ${me.name}] ${note}`;
    if(S[pickedOutcome].final) patch.hidden=true;
    await patchP(id,patch);
    let calMsg="";
    if(o.calEvent && !(wantCal)) await removeEvent(o);
    if(wantCal){ calMsg=await upsertEvent(o,pickedOutcome,cb,note) ? " · μπήκε στο Outlook" : " · δεν μπήκε στο Outlook"; }
    else if(DATED.has(pickedOutcome)) askNotify();
    toast("Αποθηκεύτηκε: "+S[pickedOutcome].l+(cb?" για "+fmtDT(cb):"")+calMsg);
    if(goNext){ skipped.add(id); currentId=null; await nextCall(); skipped.delete(id); } else { pickedOutcome=null; renderCallCard(); }
    renderTodayCallbacks();
  }catch(e){ toast(e.message||"Δεν αποθηκεύτηκε"); $("save-next").disabled=false; $("save-stay").disabled=false; }
}
function renderHistory(o){
  const h=(o.history||[]).slice().reverse();
  $("callhist").innerHTML=h.length?h.map(x=>`<li><div class="when">${fmtDT(x.t)} · ${esc(x.by||"—")}</div>${pill(x.s)}${x.cb?` <span class="muted" style="font-size:13px">→ ${fmtDT(x.cb)}</span>`:""}${x.n?`<div>${esc(x.n)}</div>`:""}</li>`).join("")
    :`<li class="muted" style="border:0">Δεν έχει κληθεί ακόμα.</li>`;
}
function renderTodayCallbacks(){
  const end=new Date(); end.setHours(23,59,59,999);
  const start=new Date(); start.setHours(0,0,0,0);
  const list=[...offices.values()].filter(o=>(o.status==="callback"&&!o.hidden&&(o.callbackAt||0)<=end.getTime())||(o.status==="appointment"&&o.callbackAt>=start.getTime()&&o.callbackAt<=end.getTime())).sort((a,b)=>(a.callbackAt||0)-(b.callbackAt||0));
  const now=Date.now();
  $("todaycb").innerHTML=list.length?list.slice(0,25).map(o=>{ const late=o.status==="callback"&&o.callbackAt&&o.callbackAt<start.getTime();
    return `<div class="cbrow"><a href="#calls" data-open="${o.id}">${o.status==="appointment"?"Ραντεβού: ":""}${esc(o.n)}</a><span class="mono${late||(o.status==="callback"&&o.callbackAt<=now)?" due":""}">${o.callbackAt?(late?fmtD(o.callbackAt)+" ":"")+new Date(o.callbackAt).toLocaleTimeString("el-GR",{hour:"2-digit",minute:"2-digit"}):"χωρίς ώρα"}</span></div>`; }).join(""):"Καμία για σήμερα.";
  $("todaycb").querySelectorAll("[data-open]").forEach(a=>a.onclick=e=>{e.preventDefault(); openItem(a.dataset.open);});
}

/* =================== Outlook ημερολόγιο =================== */
let calReady=false;
const calOn = () => calReady && lsGet("crm_cal")!=="0";
async function checkCalendar(){ calReady=!!(await tokenSilent(CAL_SCOPES)); renderCalStatus(); }
function connectCalendar(){
  if(!confirm("Θα ανοίξει η σελίδα της Microsoft για να επιτρέψεις στην εφαρμογή να βάζει υπενθυμίσεις στο ημερολόγιο Outlook σου. Ό,τι δεν έχεις αποθηκεύσει στην τρέχουσα κλήση θα χαθεί. Συνέχεια;")) return;
  lsSet("crm_cal","1"); release(currentId); pca.acquireTokenRedirect({scopes:CAL_SCOPES,account,redirectStartPage:APP_URL+"#calendar"});
}
function renderCalStatus(){
  const el=$("calstatus"); if(!el) return;
  el.innerHTML=calReady?`<span class="pill s-interested">Outlook συνδεδεμένο</span> <span class="muted">Οι επανακλήσεις και τα ραντεβού που αποθηκεύεις μπαίνουν και στο ημερολόγιό σου με υπενθύμιση 10′ πριν.</span>`
    :`<button class="btn small" id="btn-cal">Σύνδεση με ημερολόγιο Outlook</button> <span class="muted">για να μπαίνουν οι επανακλήσεις και στο Outlook (και στο κινητό) με υπενθύμιση.</span>`;
  if($("btn-cal")) $("btn-cal").onclick=connectCalendar;
}
function eventBody(o,kind,note){
  const tel=(o.ph||"").replace(/\s/g,"");
  return `<p><b>${esc(o.n)}</b>${o.cat?` — ${esc(o.cat)}`:""}${o.a?`, ${esc(o.a)}`:""}</p>
<p>Τηλέφωνο: <a href="tel:${esc(tel)}">${esc(o.ph)}</a>${o.p?`<br>Υπεύθυνος: ${esc(o.p)}`:""}${o.em?`<br>Email: ${esc(o.em)}`:""}${o.ad?`<br>Διεύθυνση: ${esc(o.ad)}`:""}</p>
${note?`<p>Σημείωση: ${esc(note)}</p>`:""}
<p><a href="${APP_URL}#open=${o.id}">Άνοιγμα στο Praxis CRM</a></p>`;
}
const utc = t => new Date(t).toISOString().slice(0,19);
async function upsertEvent(o,kind,t,note){
  try{
    if(o.calEvent) await removeEvent(o,true);
    const appt=kind==="appointment";
    const ev={subject:(appt?"Ραντεβού: ":"Επανάκληση: ")+o.n+(o.ph?" — "+o.ph:""),
      body:{contentType:"HTML",content:eventBody(o,kind,note)},
      start:{dateTime:utc(t),timeZone:"UTC"},end:{dateTime:utc(t+(appt?60:15)*60e3),timeZone:"UTC"},
      isReminderOn:true,reminderMinutesBeforeStart:appt?30:10,showAs:appt?"busy":"free",categories:["Praxis CRM"]};
    if(appt&&o.ad) ev.location={displayName:o.ad};
    const r=await g("/me/events",{method:"POST",scopes:CAL_SCOPES,silent:true,body:ev});
    const ref=me.email+"|"+r.id;
    try{ await patchP(o.id,{calEvent:ref}); }catch(e){ o.calEvent=ref; }
    return true;
  }catch(e){ if(e.status===401){ calReady=false; renderCalStatus(); } return false; }
}
async function removeEvent(o,keepField){
  const [owner,id]=String(o.calEvent||"").split("|");
  if(id&&owner===me.email){ try{ await g("/me/events/"+encodeURIComponent(id),{method:"DELETE",scopes:CAL_SCOPES,silent:true}); }catch(e){} }
  if(!keepField){ try{ await patchP(o.id,{calEvent:""}); }catch(e){} }
}

/* =================== ειδοποιήσεις =================== */
const notified=new Set();
let swReg=null;
async function registerSW(){
  if(!("serviceWorker" in navigator)) return;
  try{ swReg=await navigator.serviceWorker.register("sw.js"); }catch(e){}
  navigator.serviceWorker.addEventListener("message",e=>{ if(e.data&&e.data.type==="open"){ const m=new URL(e.data.url).hash.match(/^#open=(\d+)$/); if(m&&offices.has(m[1])) openItem(m[1]); } });
}
registerSW();
function askNotify(){ try{ if("Notification" in window && Notification.permission==="default") Notification.requestPermission(); }catch(e){} }
function checkDue(){
  const now=Date.now();
  const due=[...offices.values()].filter(o=>o.status==="callback"&&!o.hidden&&o.callbackAt&&o.callbackAt<=now&&o.callbackAt>now-12*3600e3&&(!o.lastBy||o.lastBy===me.name)&&!notified.has(o.id+":"+o.callbackAt));
  for(const o of due){
    notified.add(o.id+":"+o.callbackAt);
    const b=$("duebar"); b.hidden=false;
    b.innerHTML=`<span>⏰ Ώρα για επανάκληση: <b>${esc(o.n)}</b> ${esc(o.ph)}</span><span style="display:flex;gap:6px"><button class="btn small primary" id="due-open">Άνοιγμα</button><button class="btn small" id="due-x">Αργότερα</button></span>`;
    $("due-open").onclick=()=>{ b.hidden=true; openItem(o.id); };
    $("due-x").onclick=()=>{ b.hidden=true; };
    try{ if("Notification" in window && Notification.permission==="granted"){
      const title="Επανάκληση: "+o.n, opt={body:(o.ph||"")+(o.p?" · "+o.p:""),tag:"crm-"+o.id,icon:"icon-192.png",badge:"icon-192.png",data:{url:APP_URL+"#open="+o.id}};
      if(swReg) swReg.showNotification(title,opt).catch(()=>{});
      else { const n=new Notification(title,opt); n.onclick=()=>{ window.focus(); openItem(o.id); n.close(); }; }
    } }catch(e){}
  }
}

/* =================== προβολή ημερολογίου =================== */
let calMonth=(()=>{ const d=new Date(); d.setDate(1); d.setHours(0,0,0,0); return d; })();
function calItems(){ return [...offices.values()].filter(o=>o.callbackAt&&(o.status==="callback"&&!o.hidden||o.status==="appointment")); }
async function openItem(id){
  const o=offices.get(id); if(!o) return;
  if((S[o.status]||S.new).queue){ show("calls"); if(currentId&&currentId!==id) release(currentId); await claim(o); openCall(id); window.scrollTo({top:0,behavior:"smooth"}); }
  else { show("list"); openDrawer(id); }
}
function renderCalendar(){
  renderCalStatus();
  const items=calItems(), now=Date.now(), todayK=dayKey(now);
  const byDay={}; for(const o of items){ const k=dayKey(o.callbackAt); (byDay[k]=byDay[k]||[]).push(o); }
  Object.values(byDay).forEach(a=>a.sort((x,y)=>x.callbackAt-y.callbackAt));
  $("calttl").textContent=calMonth.toLocaleDateString("el-GR",{month:"long",year:"numeric"});
  const first=new Date(calMonth); const offset=(first.getDay()+6)%7; first.setDate(1-offset);
  const days=["Δευ","Τρι","Τετ","Πεμ","Παρ","Σαβ","Κυρ"];
  let html=days.map(d=>`<div class="cal-h">${d}</div>`).join("");
  for(let i=0;i<42;i++){
    const d=new Date(first); d.setDate(first.getDate()+i); const k=dayKey(d.getTime());
    if(i>=35&&d.getMonth()!==calMonth.getMonth()) break;
    const list=byDay[k]||[];
    html+=`<div class="cal-d${d.getMonth()!==calMonth.getMonth()?" out":""}${k===todayK?" today":""}"><div class="n">${d.getDate()}</div>${list.slice(0,4).map(o=>{
      const cls=o.status==="appointment"?"appt":(o.callbackAt<=now?"late":"cb");
      return `<button class="cal-i ${cls}" data-id="${o.id}" title="${esc(o.n)} — ${esc(o.ph)}"><span class="mono">${new Date(o.callbackAt).toLocaleTimeString("el-GR",{hour:"2-digit",minute:"2-digit"})}</span> ${esc(o.n)}</button>`; }).join("")}${list.length>4?`<div class="more">+${list.length-4} ακόμα</div>`:""}</div>`;
  }
  $("calgrid").innerHTML=html;
  $("calgrid").querySelectorAll("[data-id]").forEach(b=>b.onclick=()=>openItem(b.dataset.id));
  // ατζέντα
  const start=new Date(); start.setHours(0,0,0,0);
  const overdue=items.filter(o=>o.status==="callback"&&o.callbackAt<start.getTime()).sort((a,b)=>a.callbackAt-b.callbackAt);
  const next=items.filter(o=>o.callbackAt>=start.getTime()&&o.callbackAt<start.getTime()+14*864e5).sort((a,b)=>a.callbackAt-b.callbackAt);
  const row=o=>`<tr data-id="${o.id}"><td class="num">${new Date(o.callbackAt).toLocaleString("el-GR",{weekday:"short",day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"})}</td><td>${pill(o.status)}</td><td>${esc(o.n)}</td><td class="num">${esc(o.ph)}</td><td>${esc(o.lastBy)}</td><td>${esc((o.notes||"").split("\n").pop().slice(0,90))}</td></tr>`;
  const undated=[...offices.values()].filter(o=>o.status==="callback"&&!o.hidden&&!o.callbackAt);
  const urow=o=>`<tr data-id="${o.id}"><td class="num muted">χωρίς ημ/νία</td><td>${pill(o.status)}</td><td>${esc(o.n)}</td><td class="num">${esc(o.ph)}</td><td>${esc(o.lastBy)}</td><td>${esc((o.notes||"").split("\n").pop().slice(0,90))}</td></tr>`;
  $("calagenda").innerHTML=(undated.length?`<tr><td colspan="6" class="sect">Επανακλήσεις χωρίς ημερομηνία (${undated.length}) — άνοιξέ τες και βάλε ημερομηνία</td></tr>`+undated.map(urow).join(""):"")+(overdue.length?`<tr><td colspan="6" class="sect late">Εκπρόθεσμες (${overdue.length})</td></tr>`+overdue.map(row).join(""):"")
    +`<tr><td colspan="6" class="sect">Επόμενες 14 ημέρες (${next.length})</td></tr>`+(next.length?next.map(row).join(""):`<tr><td colspan="6" class="muted">Δεν υπάρχει κάτι προγραμματισμένο.</td></tr>`);
  $("calagenda").querySelectorAll("tr[data-id]").forEach(tr=>tr.onclick=()=>openItem(tr.dataset.id));
}
$("calprev").onclick=()=>{ calMonth.setMonth(calMonth.getMonth()-1); renderCalendar(); };
$("calnext").onclick=()=>{ calMonth.setMonth(calMonth.getMonth()+1); renderCalendar(); };
$("caltoday").onclick=()=>{ calMonth=new Date(); calMonth.setDate(1); calMonth.setHours(0,0,0,0); renderCalendar(); };

/* λίστα */
function cats(){ return [...new Set([...DEFAULT_CATS, ...[...offices.values()].map(o=>o.cat).filter(Boolean)])].sort((a,b)=>a.localeCompare(b,"el")); }
function fillSelects(){
  const areas=[...new Set([...offices.values()].map(o=>o.a).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"el"));
  const used=[...new Set([...offices.values()].map(o=>o.cat).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"el"));
  for(const sel of [$("qarea"),$("farea")]){ const cur=sel.value; sel.innerHTML=`<option value="">Όλες οι περιοχές</option>`+areas.map(a=>`<option${a===cur?" selected":""}>${esc(a)}</option>`).join(""); }
  for(const sel of [$("qcat"),$("fcat"),$("rcat")]){ const cur=sel.value; sel.innerHTML=`<option value="">${sel.id==="rcat"?"Όλες":"Όλες οι κατηγορίες"}</option>`+used.map(a=>`<option${a===cur?" selected":""}>${esc(a)}</option>`).join(""); }
  $("catlist").innerHTML=cats().map(c=>`<option value="${esc(c)}">`).join("");
  if($("fstatus").options.length<2) $("fstatus").innerHTML+=STATUSES.map(s=>`<option value="${s.k}">${esc(s.l)}</option>`).join("");
}
function renderList(){
  const q=norm($("q").value), area=$("farea").value, cat=$("fcat").value, st=$("fstatus").value, showH=$("fhidden").checked;
  const rows=[...offices.values()].filter(o=>{
    if(!showH&&o.hidden) return false; if(area&&o.a!==area) return false; if(cat&&o.cat!==cat) return false; if(st&&o.status!==st) return false;
    if(q){ const hay=norm([o.n,o.ph,o.em,o.p,o.ad,o.a,o.afm,o.cat].join(" ")); if(!hay.includes(q)&&!hay.replace(/\s/g,"").includes(q.replace(/\s/g,""))) return false; }
    return true;
  }).sort((a,b)=>(a.cat||"").localeCompare(b.cat||"","el")||(a.a||"").localeCompare(b.a||"","el")||(a.n||"").localeCompare(b.n||"","el"));
  const hc=[...offices.values()].filter(o=>o.hidden).length;
  $("listcount").textContent=`${rows.length} υποψήφιοι${!showH&&hc?` · ${hc} κρυμμένοι`:""}`;
  $("tbody").innerHTML=rows.length?rows.slice(0,1500).map(o=>`<tr data-id="${o.id}" class="${o.hidden?"hiddenrow":""}"><td>${esc(o.n)}</td><td>${esc(o.cat)}</td><td>${esc(o.a)}</td><td class="num">${esc(o.ph)}</td><td>${pill(o.status)}</td><td class="num">${o.attempts||0}</td><td class="num">${o.lastCallAt?fmtDT(o.lastCallAt):"—"}</td><td>${esc((o.notes||"").slice(-80))}</td></tr>`).join("")
    :`<tr><td colspan="8" class="empty">Κανένας υποψήφιος με αυτά τα φίλτρα.</td></tr>`;
  $("tbody").querySelectorAll("tr[data-id]").forEach(tr=>tr.onclick=()=>openDrawer(tr.dataset.id));
}
["q","farea","fcat","fstatus","fhidden"].forEach(id=>$(id).addEventListener("input",renderList));

function openDrawer(id){
  const o=id?offices.get(id):{n:"",cat:$("fcat").value||"",a:"",ad:"",ph:"",em:"",p:"",f:"",w:"",status:"new",afm:"",gemi:"",turnover:null,employees:null,finYear:null,finNotes:""};
  if(!o) return; const isNew=!id;
  const v=x=>esc(x??"");
  const gq=encodeURIComponent(o.n||"");
  $("drawerpanel").innerHTML=`
    <div class="callhead"><h1 style="font-size:20px;margin:0">${isNew?"Νέος υποψήφιος":esc(o.n)}</h1><button class="btn small" id="dclose">Κλείσιμο</button></div>
    ${isNew?"":`<div style="margin-top:8px">${pill(o.status)} <span class="muted" style="font-size:13px">κλήσεις: ${o.attempts||0}</span></div>
    <div class="dauditbox"><div class="dahead"><h2 style="margin:0">Έλεγχοι IT</h2><button class="btn small" id="daudit">+ Νέος έλεγχος IT</button></div><div id="daudits" class="daudits muted">Φόρτωση…</div></div>`}
    <div style="display:flex;flex-direction:column;gap:10px;margin-top:16px">
      <label class="f">Επωνυμία<input id="d-n" value="${v(o.n)}"></label>
      <div class="row"><label class="f">Κατηγορία<input id="d-cat" list="catlist" value="${v(o.cat)}"></label><label class="f">Περιοχή<input id="d-a" value="${v(o.a)}"></label></div>
      <div class="row"><label class="f">Τηλέφωνο<input id="d-ph" value="${v(o.ph)}"></label><label class="f">Email<input id="d-em" value="${v(o.em)}"></label></div>
      <label class="f">Διεύθυνση<input id="d-ad" value="${v(o.ad)}"></label>
      <div class="row"><label class="f">Υπεύθυνος<input id="d-p" value="${v(o.p)}"></label><label class="f">Νομική μορφή<input id="d-f" value="${v(o.f)}"></label></div>
      <label class="f">Website<input id="d-w" value="${v(o.w)}"></label>
      <h2 style="margin:8px 0 0">Οικονομικά στοιχεία</h2>
      <div class="row"><label class="f">ΑΦΜ<input id="d-afm" value="${v(o.afm)}"></label><label class="f">Αρ. ΓΕΜΗ<input id="d-gemi" value="${v(o.gemi)}"></label></div>
      <div class="row"><label class="f">Κύκλος εργασιών (€)<input id="d-turnover" type="number" step="1" value="${v(o.turnover)}"></label><label class="f">Εργαζόμενοι<input id="d-employees" type="number" step="1" value="${v(o.employees)}"></label><label class="f">Έτος<input id="d-finYear" type="number" step="1" value="${v(o.finYear)}"></label></div>
      <label class="f">Σημειώσεις οικονομικών<textarea id="d-finNotes">${v(o.finNotes)}</textarea></label>
      <div class="actions" style="margin:0"><a class="btn small" target="_blank" rel="noopener" href="https://www.google.com/search?q=${gq}+%CE%91%CE%A6%CE%9C+%CE%93%CE%95%CE%9C%CE%97">Αναζήτηση ΑΦΜ/ΓΕΜΗ</a><a class="btn small" target="_blank" rel="noopener" href="https://publicity.businessportal.gr/">ΓΕΜΗ (δημοσιότητα)</a><a class="btn small" target="_blank" rel="noopener" href="https://www.google.com/search?q=${gq}+%CE%B9%CF%83%CE%BF%CE%BB%CE%BF%CE%B3%CE%B9%CF%83%CE%BC%CF%8C%CF%82">Ισολογισμοί</a></div>
      ${isNew?"":`<h2 style="margin:8px 0 0">Κατάσταση</h2><label class="f">Κατάσταση<select id="d-status">${STATUSES.map(s=>`<option value="${s.k}"${s.k===o.status?" selected":""}>${esc(s.l)}</option>`).join("")}</select></label>
      <label class="f">Σημειώσεις<textarea id="d-notes" style="min-height:110px">${v(o.notes)}</textarea></label>`}
    </div>
    <div class="actions">
      <button class="btn primary" id="dsave">${isNew?"Προσθήκη":"Αποθήκευση"}</button>
      ${isNew?"":`<button class="btn" id="dcall">Κάλεσε τώρα</button><button class="btn" id="dhide">${o.hidden?"Επαναφορά":"Απόκρυψη"}</button>`}
    </div>
    ${isNew?"":`<h2>Ιστορικό</h2><ul class="hist">${(o.history||[]).slice().reverse().map(x=>`<li><div class="when">${fmtDT(x.t)} · ${esc(x.by||"—")}</div>${pill(x.s)}${x.n?`<div>${esc(x.n)}</div>`:""}</li>`).join("")||'<li class="muted" style="border:0">Δεν έχει κληθεί ακόμα.</li>'}</ul>`}`;
  $("drawer").hidden=false;
  $("dclose").onclick=closeDrawer;
  $("dsave").onclick=async()=>{
    const g_=k=>$("d-"+k).value.trim();
    const data={n:g_("n"),cat:g_("cat"),a:g_("a"),ad:g_("ad"),ph:g_("ph"),p:g_("p"),em:g_("em"),f:g_("f"),w:g_("w"),afm:g_("afm"),gemi:g_("gemi"),turnover:g_("turnover"),employees:g_("employees"),finYear:g_("finYear"),finNotes:g_("finNotes")};
    ["turnover","employees","finYear"].forEach(k=>data[k]=data[k]===""?null:Number(data[k]));
    if(!data.n){ toast("Γράψε την επωνυμία"); return; }
    try{
      if(isNew){
        const k=dedupKey(data); const dup=[...offices.values()].find(x=>dedupKey(x)===k);
        if(dup){ toast("Υπάρχει ήδη: "+dup.n); return; }
        await createP({...data,status:"new",attempts:0,hidden:false,notes:"",src:"Προστέθηκε χειροκίνητα"});
      } else {
        data.notes=$("d-notes").value; const ns=$("d-status").value;
        if(ns!==o.status){ data.status=ns; if(!DATED.has(ns)){ data.callbackAt=null; if(o.calEvent) await removeEvent(o); } await addCall(o,{s:ns,n:"Αλλαγή κατάστασης χειροκίνητα",manual:true,t:Date.now()}); }
        await patchP(id,data);
      }
      toast(isNew?"Προστέθηκε":"Αποθηκεύτηκε"); closeDrawer(); refreshAll(); if(currentId===id) renderCallCard();
    }catch(e){ toast(e.message); }
  };
  if(!isNew){
    $("dcall").onclick=()=>{ closeDrawer(); show("calls"); openCall(id); };
    $("daudit").onclick=()=>{ closeDrawer(); show("audit"); auditOpen(id); };
    (async()=>{
      const box=$("daudits"); let rows=[];
      try{ rows=await fetchAudits(id); }catch(e){ if(box) box.textContent="Δεν φορτώθηκαν οι έλεγχοι: "+e.message; return; }
      if(!box||!box.isConnected||$("drawer").hidden) return;
      box.classList.toggle("muted",!rows.length);
      box.innerHTML=rows.length?rows.map(r=>`<button class="arow" data-rid="${r.id}"><span>${auditLine(r)}</span><span class="go">Άνοιγμα ›</span></button>`).join("")
        :"Δεν έχει γίνει έλεγχος IT. Πάτα «+ Νέος έλεγχος IT».";
      box.querySelectorAll("[data-rid]").forEach(b=>b.onclick=()=>{ closeDrawer(); show("audit"); auditOpen(id,null,b.dataset.rid); });
    })();
    $("dhide").onclick=async()=>{ try{ await patchP(id,{hidden:!o.hidden}); toast(o.hidden?"Εμφανίζεται":"Κρύφτηκε"); closeDrawer(); refreshAll(); }catch(e){ toast(e.message); } };
  }
}
function closeDrawer(){ $("drawer").hidden=true; }
$("drawer").addEventListener("click",e=>{ if(e.target.id==="drawer") closeDrawer(); });
document.addEventListener("keydown",e=>{ if(e.key==="Escape") closeDrawer(); });
$("btn-add").onclick=()=>openDrawer(null);

/* αναφορές */
function range(){ const f=$("rfrom").value,t=$("rto").value; return [f?new Date(f+"T00:00").getTime():0, t?new Date(t+"T23:59:59").getTime():Infinity]; }
function renderReports(){
  const cat=$("rcat").value;
  const os=[...offices.values()].filter(o=>!cat||o.cat===cat);
  const ids=new Set(os.map(o=>o.id));
  const all=calls.filter(c=>!c.manual&&ids.has(c.pid)), [from,to]=range();
  const inR=all.filter(c=>c.t>=from&&c.t<=to), today=dayKey(Date.now());
  const leads=os.filter(o=>LEADS.has(o.status));
  const kp=[["Υποψήφιοι",os.length],["Καλέστηκαν τουλάχιστον 1 φορά",os.filter(o=>o.status!=="new").length],["Κλήσεις στο διάστημα",inR.length],
    ["Κλήσεις σήμερα",all.filter(c=>dayKey(c.t)===today).length],["Ενδιαφέρον / ραντεβού",leads.length],["Εκκρεμείς επανακλήσεις",os.filter(o=>!o.hidden&&o.status==="callback").length]];
  $("kpis").innerHTML=kp.map(([l,v])=>`<div class="kpi"><div class="v">${v}</div><div class="l">${l}</div></div>`).join("");
  const max=Math.max(1,...STATUSES.map(s=>os.filter(o=>o.status===s.k).length));
  $("bystatus").innerHTML=STATUSES.map(s=>{const n=os.filter(o=>o.status===s.k).length;return `<div class="bar"><span>${esc(s.l)}</span><div class="track"><div class="fill" style="width:${n/max*100}%;background:${s.color}"></div></div><span class="n">${n}</span></div>`}).join("");
  const days=[]; for(let i=13;i>=0;i--){ const d=new Date(); d.setHours(12,0,0,0); d.setDate(d.getDate()-i); days.push(d); }
  const counts=days.map(d=>all.filter(c=>dayKey(c.t)===dayKey(d.getTime())).length), mx=Math.max(1,...counts);
  $("bydays").innerHTML=days.map((d,i)=>`<div class="d" title="${fmtD(d)}: ${counts[i]}"><span class="mono">${counts[i]||""}</span><div class="b" style="height:${counts[i]/mx*80}px"></div><span>${fmtD(d)}</span></div>`).join("");
  const by={}; for(const c of inR){ const k=c.by||"—"; by[k]=by[k]||{n:0,a:0,l:0}; by[k].n++; if(ANSWERED.has(c.s)) by[k].a++; if(LEADS.has(c.s)) by[k].l++; }
  $("bycaller").innerHTML=Object.keys(by).length?Object.keys(by).sort((a,b)=>by[b].n-by[a].n).map(k=>`<tr><td>${esc(k)}</td><td class="num">${by[k].n}</td><td class="num">${by[k].a}</td><td class="num">${by[k].l}</td></tr>`).join(""):`<tr><td colspan="4" class="muted">Καμία κλήση στο διάστημα.</td></tr>`;
  const grp=o=>cat?(o.a||"—"):(o.cat||"—");
  const ar={}; for(const o of os){ const k=grp(o); ar[k]=ar[k]||{t:0,c:0,l:0}; ar[k].t++; if(o.status!=="new") ar[k].c++; if(LEADS.has(o.status)) ar[k].l++; }
  $("byarea").innerHTML=Object.keys(ar).sort((a,b)=>a.localeCompare(b,"el")).map(k=>`<tr><td>${esc(k)}</td><td class="num">${ar[k].t}</td><td class="num">${ar[k].c}</td><td class="num">${ar[k].l}</td></tr>`).join("");
  $("leads").innerHTML=leads.length?leads.map(o=>`<tr><td>${esc(o.n)}</td><td>${esc(o.cat)}</td><td>${pill(o.status)}</td><td>${esc(o.p)}</td><td class="num">${esc(o.ph)}</td><td>${esc(o.em)}</td><td>${esc((o.notes||"").slice(-120))}</td></tr>`).join(""):`<tr><td colspan="7" class="muted">Δεν υπάρχουν ακόμα ενδιαφερόμενοι.</td></tr>`;
}
["rfrom","rto","rcat"].forEach(id=>$(id).addEventListener("input",renderReports));
function downloadCsv(name, rows){
  const q=v=>'"'+String(v??"").replace(/"/g,'""')+'"';
  const blob=new Blob(["﻿"+rows.map(r=>r.map(q).join(";")).join("\r\n")],{type:"text/csv;charset=utf-8"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=name; document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(a.href); a.remove();},500);
}
$("btn-csv").onclick=()=>{
  const rows=[["Επωνυμία","Κατηγορία","Περιοχή","Διεύθυνση","Νομική μορφή","Υπεύθυνος","Τηλέφωνο","Email","Website","ΑΦΜ","ΓΕΜΗ","Κύκλος εργασιών","Εργαζόμενοι","Έτος","Κατάσταση","Κλήσεις","Τελευταία κλήση","Από","Επανάκληση","Κρυμμένο","Σημειώσεις"]];
  for(const o of offices.values()) rows.push([o.n,o.cat,o.a,o.ad,o.f,o.p,o.ph,o.em,o.w,o.afm,o.gemi,o.turnover,o.employees,o.finYear,(S[o.status]||S.new).l,o.attempts,o.lastCallAt?fmtDT(o.lastCallAt):"",o.lastBy,o.callbackAt?fmtDT(o.callbackAt):"",o.hidden?"Ναι":"",o.notes]);
  downloadCsv(`ypopsifioi-${dayKey(Date.now())}.csv`,rows);
};
$("btn-csv-calls").onclick=()=>{
  const rows=[["Ημερομηνία","Υποψήφιος","Κατηγορία","Τηλέφωνο","Αποτέλεσμα","Σημείωση","Επανάκληση","Συνεργάτης"]];
  for(const c of calls){ const o=offices.get(c.pid)||{}; rows.push([fmtDT(c.t),o.n,o.cat,o.ph,(S[c.s]||{}).l||c.s,c.n,c.cb?fmtDT(c.cb):"",c.by]); }
  downloadCsv(`istoriko-kliseon-${dayKey(Date.now())}.csv`,rows);
};

/* =================== έλεγχος IT =================== */
const AUDIT_COLUMNS=[["ProspectId","number",true],["AuditAt","dateTime"],["AuditorName","text"],["Proposal","text"],["Critical","number"],["Issues","number"],["Data","note"]];
const AQ=[
  {k:"breaks", q:"Τι σας χαλάει πιο συχνά;", o:["Internet","Εκτυπωτές","Αργοί Η/Υ","Email","Πρόγραμμα λογιστικής"]},
  {k:"who",    q:"Ποιος το φτιάχνει σήμερα;", o:["Κανείς","Γνωστός","Freelancer","Εταιρεία IT"], one:true},
  {k:"speed",  q:"Πόσο γρήγορα έρχεται;", o:["Ίδια μέρα","1–2 μέρες","Όποτε μπορεί"], one:true},
  {k:"down",   q:"Ακινησία στην τελευταία βλάβη", o:["Κάτω από 1 ώρα","Μισή μέρα","1+ μέρα"], one:true},
  {k:"loss",   q:"Αν χαλάσει ο Η/Υ με τα αρχεία, τι γίνεται;", o:["Υπάρχει backup","Δεν ξέρω","Χάνονται"], one:true},
  {k:"remote", q:"Δουλεύετε από έξω (σπίτι, κινητό);", o:["Ναι","Όχι"], one:true},
  {k:"clients",q:"Οι πελάτες σας ζητούν βοήθεια για IT ή τιμολόγηση;", o:["Συχνά","Σπάνια","Ποτέ"], one:true},
  {k:"plans",  q:"Σχέδια επόμενων 12 μηνών", o:["Νέα άτομα","Μετακόμιση","Νέο λογισμικό","Τίποτα"]}];
const AT=["Backup υπάρχει, εκτός γραφείου, δοκιμασμένο","Router / firewall επαγγελματικό, ενημερωμένο","WiFi με ξεχωριστό δίκτυο επισκεπτών",
  "Windows 10/11 ενημερωμένα, χωρίς παλιές εκδόσεις","Antivirus ενεργό σε όλους τους Η/Υ","Ξεχωριστοί κωδικοί ανά χρήστη, όχι κοινοί",
  "Email σε δικό τους domain, με 2FA","Απομακρυσμένη πρόσβαση ασφαλής (όχι ανοιχτό AnyDesk/RDP)","UPS στον server / κύριο Η/Υ",
  "Εκτυπωτές και σαρωτές λειτουργούν σωστά","Λογισμικό λογιστικής / τιμολόγησης σταθερό"];
const ALEVELS=[["ok","ΟΚ"],["issue","Θέμα"],["crit","Κρίσιμο"]];
const APLANS=["Starter","Business","Professional","Project"];
const ADRAFT="crm_audit_draft";
let auditList=null, auditPid=null, auditItemId=null, auditRows=[], auditReady=false;

const auditLabel=o=>`${o.n} · ${o.ph||o.a||"#"+o.id}`;
const todayInput=()=>dayKey(Date.now());
function abox(g,val,checked,lbl,one,aria){
  return `<label class="ab"><input type="checkbox" data-g="${esc(g)}"${one?' data-one="1"':""} value="${esc(val)}"${checked?" checked":""}${aria?` aria-label="${esc(aria)}"`:""}>${lbl?" "+esc(lbl):""}</label>`;
}
function renderAuditSheet(o,d){
  d=d||{}; const info=d.info||{}, q=d.q||{}, t=d.t||[], f=d.f||[], nx=d.next||{};
  const v=(x,def)=>esc(x!=null&&x!==""?x:(def??""));
  const sh=$("auditsheet");
  sh.innerHTML=`
    <div class="shead"><h1>Checklist Ελέγχου IT</h1><span class="co"><img class="slogo" src="praxis-logo.svg" alt="PRAXIS iT — Technology. Put to work."></span></div>
    <div class="ainfo">
      <label>Επιχείρηση<input type="text" id="ai-n" value="${v(info.n,o&&o.n)}"></label>
      <label>Υπεύθυνος<input type="text" id="ai-p" value="${v(info.p,o&&o.p)}"></label>
      <label>Ημερομηνία<input type="date" id="ai-date" value="${v(info.date,o?todayInput():"")}"></label>
      <label>Άτομα / Η/Υ<input type="text" id="ai-staff" value="${v(info.staff,o&&o.employees!=null?String(o.employees):"")}"></label>
    </div>
    <h3>Ερωτήσεις προς τον πελάτη</h3>
    <div class="tablewrap"><table class="qtab"><thead><tr><th style="width:42%">Ερώτηση</th><th>Απάντηση</th></tr></thead><tbody>
      ${AQ.map(x=>`<tr><td>${esc(x.q)}</td><td>${x.o.map(opt=>abox("q_"+x.k,opt,(q[x.k]||[]).includes(opt),opt,x.one)).join("")}</td></tr>`).join("")}
    </tbody></table></div>
    <h3>Τεχνικός έλεγχος</h3>
    <div class="tablewrap"><table class="ttab"><thead><tr><th>Σημείο</th>${ALEVELS.map(([,l])=>`<th class="c">${l}</th>`).join("")}<th style="width:26%">Σημείωση</th></tr></thead><tbody>
      ${AT.map((p,i)=>{ const cur=(t[i]||{}).s||""; return `<tr><td>${esc(p)}</td>${ALEVELS.map(([k,l])=>`<td class="c c-${k}">${abox("t"+i,k,cur===k,"",true,p+": "+l).replace("</label>",`<span class="ml">${l}</span></label>`)}</td>`).join("")}<td><input type="text" id="at-n${i}" placeholder="Σημείωση" value="${v((t[i]||{}).n)}" aria-label="Σημείωση: ${esc(p)}"></td></tr>`; }).join("")}
    </tbody></table></div>
    <h3>Ευρήματα και επόμενο βήμα</h3>
    <div class="agrid">
      <div class="afind">${[0,1,2].map(i=>`<label><span class="mono">${i+1}.</span><input type="text" id="af${i}" value="${v(f[i])}" aria-label="Εύρημα ${i+1}"></label>`).join("")}</div>
      <div class="anext">
        ${abox("nx_report","1",!!nx.report,"Αποστολή αναφοράς σε 24 ώρες")}
        <label class="ameet">Δεύτερη συνάντηση<input type="datetime-local" id="an-meet" value="${v(nx.meet)}"></label>
        <div><span class="muted">Πρόταση:</span> ${APLANS.map(p=>abox("nx_plan",p,nx.plan===p,p,true)).join("")}</div>
      </div>
    </div>
    <div class="afoot"><span>PRAXIS IT SOLUTIONS Ε.Ε. · 210 444 0218 · info@praxisit.gr · praxisit.gr</span><span>${o?esc(o.ph||""):""}</span></div>`;
}
function auditCollect(){
  const sh=$("auditsheet"); const val=id=>{ const el=sh.querySelector("#"+id); return el?el.value.trim():""; };
  const gv=g=>[...sh.querySelectorAll(`input[data-g="${g}"]:checked`)].map(x=>x.value);
  return {v:1, info:{n:val("ai-n"),p:val("ai-p"),date:val("ai-date"),staff:val("ai-staff")},
    q:Object.fromEntries(AQ.map(x=>[x.k,gv("q_"+x.k)])),
    t:AT.map((_,i)=>({s:gv("t"+i)[0]||"",n:val("at-n"+i)})),
    f:[0,1,2].map(i=>val("af"+i)),
    next:{report:gv("nx_report").length>0, meet:$("an-meet")?$("an-meet").value:"", plan:gv("nx_plan")[0]||""}};
}
function auditSaveDraft(){ if(!$("auditsheet").innerHTML) return; lsSet(ADRAFT,JSON.stringify({pid:auditPid,item:auditItemId,data:auditCollect()})); }
function auditDraft(){ try{ return JSON.parse(lsGet(ADRAFT)||"null"); }catch(e){ return null; } }
function auditStatus(){
  const o=auditPid&&offices.get(auditPid);
  const r=auditItemId&&auditRows.find(x=>x.id===auditItemId);
  $("astatus").textContent=!o?"Κενό φύλλο. Για αποθήκευση διάλεξε πρώτα πελάτη από τη λίστα.":
    r?`Έλεγχος της ${r.at?new Date(r.at).toLocaleDateString("el-GR"):"—"} από ${r.by||"—"}. Η αποθήκευση τον ενημερώνει.`:
    `Νέος έλεγχος για ${o.n}. Τα στοιχεία κρατιούνται πρόχειρα σε αυτή τη συσκευή μέχρι την αποθήκευση.`;
}
async function fetchAudits(pid){
  if(!auditList||!pid) return [];
  const items=await gAll(`/sites/${site.id}/lists/${auditList.id}/items?$expand=fields&$filter=fields/ProspectId eq ${Number(pid)}&$top=200`);
  return items.filter(it=>Number((it.fields||{}).ProspectId)===Number(pid)).map(it=>{ const fl=it.fields||{}; let data=null; try{ data=JSON.parse(fl.Data||"null"); }catch(e){}
    return {id:String(it.id),at:toMs(fl.AuditAt),by:fl.AuditorName||"",plan:fl.Proposal||"",crit:fl.Critical||0,iss:fl.Issues||0,data}; }).sort((a,b)=>(b.at||0)-(a.at||0));
}
const auditLine=r=>`${r.at?new Date(r.at).toLocaleDateString("el-GR"):"—"} · ${esc(r.by||"—")}${r.plan?" · "+esc(r.plan):""} · ${r.crit} κρίσιμα, ${r.iss} θέματα`;
async function auditLoadRows(pid){
  auditRows=[];
  try{ auditRows=await fetchAudits(pid); }catch(e){ toast("Δεν φορτώθηκαν οι προηγούμενοι έλεγχοι: "+e.message); }
  $("aprev").innerHTML=`<option value="">Νέος έλεγχος</option>`+auditRows.map(r=>`<option value="${r.id}"${r.id===auditItemId?" selected":""}>${r.at?new Date(r.at).toLocaleDateString("el-GR"):"—"} · ${esc(r.by)}${r.plan?" · "+esc(r.plan):""}${r.crit?` · ${r.crit} κρίσιμα`:""}</option>`).join("");
}
function auditUnsaved(){ const d=auditDraft(); return !!(d&&d.pid&&d.pid===auditPid&&!d.item); }
async function auditOpen(id,draft,itemId){
  const o=offices.get(id); if(!o) return;
  if(auditPid&&auditPid!==id&&auditUnsaved()&&!confirm("Ο έλεγχος που συμπληρώνεις δεν έχει αποθηκευτεί. Να ανοίξει άλλος πελάτης; (Ό,τι δεν αποθήκευσες θα χαθεί)")){ $("aq").value=auditLabel(offices.get(auditPid)); return; }
  auditPid=id; auditItemId=itemId||(draft&&draft.item)||null;
  $("aq").value=auditLabel(o);
  renderAuditSheet(o,draft&&draft.data);
  auditStatus();
  await auditLoadRows(id);
  if(itemId){ const r=auditRows.find(x=>x.id===itemId); if(r){ renderAuditSheet(o,r.data); $("aprev").value=r.id; } else auditItemId=null; }
  auditStatus();
}
/* επιλογή πελάτη από τη λίστα του CRM */
let aqItems=[], aqIdx=-1;
function aqMatches(q){
  const now=Date.now(), all=[...offices.values()];
  if(!q){
    const appt=all.filter(o=>o.status==="appointment").sort((a,b)=>Math.abs((a.callbackAt||0)-now)-Math.abs((b.callbackAt||0)-now));
    const leads=all.filter(o=>o.status==="interested"||o.status==="send_info").sort((a,b)=>(b.lastCallAt||0)-(a.lastCallAt||0));
    const rest=all.filter(o=>!LEADS.has(o.status)).sort((a,b)=>(a.n||"").localeCompare(b.n||"","el"));
    return [...appt,...leads,...rest].slice(0,40);
  }
  const nq=norm(q), dq=q.replace(/\D/g,"");
  return all.map(o=>{ const n=norm(o.n), hay=norm([o.n,o.p,o.a,o.cat,o.afm,o.em].join(" "));
      let s=n.startsWith(nq)?0:n.includes(nq)?1:hay.includes(nq)?2:(dq.length>=3&&String(o.ph||"").replace(/\D/g,"").includes(dq))?2:9;
      return [s,o]; }).filter(x=>x[0]<9).sort((a,b)=>a[0]-b[0]||(a[1].n||"").localeCompare(b[1].n||"","el")).slice(0,40).map(x=>x[1]);
}
function aqRender(){
  const box=$("aqlist"), q=$("aq").value.trim();
  const cur=auditPid&&offices.get(auditPid); const query=(cur&&q===auditLabel(cur))?"":q;
  aqItems=aqMatches(query); aqIdx=aqItems.length?0:-1;
  box.innerHTML=aqItems.length?(query?"":`<div class="aqhead">Ραντεβού και ενδιαφερόμενοι πρώτα · γράψε για αναζήτηση</div>`)+aqItems.map((o,i)=>`<div class="aqopt" role="option" id="aqo${i}" data-i="${i}" aria-selected="${i===aqIdx}"><div class="aqmain"><b>${esc(o.n)}</b><span class="muted">${[o.cat,o.a,o.ph].filter(Boolean).map(esc).join(" · ")}</span></div>${pill(o.status)}</div>`).join("")
    :`<div class="aqhead">Δεν βρέθηκε. Πρόσθεσέ τον από τη «Λίστα» → «+ Νέος υποψήφιος».</div>`;
  box.hidden=false; $("aq").setAttribute("aria-expanded","true");
  box.querySelectorAll(".aqopt").forEach(el=>el.addEventListener("mousedown",e=>{ e.preventDefault(); aqPick(Number(el.dataset.i)); }));
}
function aqMove(d){ if(!aqItems.length) return; aqIdx=(aqIdx+d+aqItems.length)%aqItems.length; $("aqlist").querySelectorAll(".aqopt").forEach((el,i)=>el.setAttribute("aria-selected",i===aqIdx)); const el=$("aqo"+aqIdx); if(el) el.scrollIntoView({block:"nearest"}); }
function aqClose(){ $("aqlist").hidden=true; $("aq").setAttribute("aria-expanded","false"); }
function aqPick(i){ const o=aqItems[i]; aqClose(); if(o){ $("aq").blur(); auditOpen(o.id); } }
$("aq").addEventListener("focus",()=>{ $("aq").select(); aqRender(); });
$("aq").addEventListener("input",aqRender);
$("aq").addEventListener("keydown",e=>{
  if(e.key==="ArrowDown"){ e.preventDefault(); if($("aqlist").hidden) aqRender(); else aqMove(1); }
  else if(e.key==="ArrowUp"){ e.preventDefault(); aqMove(-1); }
  else if(e.key==="Enter"){ e.preventDefault(); if(aqIdx>=0) aqPick(aqIdx); }
  else if(e.key==="Escape"){ aqClose(); }
});
$("aq").addEventListener("blur",()=>setTimeout(()=>{ aqClose(); const cur=auditPid&&offices.get(auditPid); $("aq").value=cur?auditLabel(cur):""; },120));
function renderAuditView(){
  const now=Date.now();
  const appts=[...offices.values()].filter(o=>o.status==="appointment"&&o.callbackAt&&o.callbackAt>now-30*864e5).sort((a,b)=>Math.abs(a.callbackAt-now)-Math.abs(b.callbackAt-now)).slice(0,8);
  $("achips").innerHTML=appts.length?`<span class="muted">Ραντεβού:</span>`+appts.map(o=>`<button class="btn small" data-aid="${o.id}">${esc(o.n)} · ${fmtDT(o.callbackAt)}</button>`).join(""):"";
  $("achips").querySelectorAll("[data-aid]").forEach(b=>b.onclick=()=>auditOpen(b.dataset.aid));
  if(!auditReady){
    auditReady=true;
    const d=auditDraft();
    if(d&&d.pid&&offices.has(d.pid)){ auditOpen(d.pid,d); return; }
    if(d&&!d.pid){ renderAuditSheet(null,d.data); auditStatus(); return; }
    renderAuditSheet(null,{}); auditStatus();
  }
}
async function ensureAuditList(){
  if(auditList) return auditList;
  auditList=await findList(CFG.auditsList);
  if(!auditList) auditList=await g(`/sites/${site.id}/lists`,{method:"POST",scopes:MANAGE_SCOPES,body:{displayName:CFG.auditsList,list:{template:"genericList"},columns:AUDIT_COLUMNS.map(colDef)}});
  return auditList;
}
async function auditSave(){
  const o=auditPid&&offices.get(auditPid);
  if(!o){ toast("Διάλεξε πρώτα πελάτη από τη λίστα"); $("aq").focus(); return; }
  auditSaveDraft();
  const d=auditCollect();
  const crit=d.t.filter(x=>x.s==="crit").length, iss=d.t.filter(x=>x.s==="issue").length;
  const at=d.info.date?new Date(d.info.date+"T12:00").getTime():Date.now();
  const fields={Title:(d.info.n||o.n).slice(0,255),ProspectId:Number(o.id),AuditAt:toIso(at),AuditorName:me.name,Proposal:d.next.plan,Critical:crit,Issues:iss,Data:JSON.stringify(d)};
  const btn=$("asave"); btn.disabled=true;
  try{
    await ensureAuditList();
    const isNew=!auditItemId;
    if(isNew){ const it=await g(`/sites/${site.id}/lists/${auditList.id}/items`,{method:"POST",body:{fields}}); auditItemId=String(it.id); }
    else await g(`/sites/${site.id}/lists/${auditList.id}/items/${auditItemId}/fields`,{method:"PATCH",body:fields});
    if(isNew){
      const line=`[${fmtDT(Date.now())} ${me.name}] Έλεγχος IT: ${crit} κρίσιμα, ${iss} θέματα${d.next.plan?" · πρόταση "+d.next.plan:""}`;
      try{ await patchP(o.id,{notes:(o.notes?o.notes+"\n":"")+line}); }catch(e){}
    }
    clearTimeout(auditSaveDraft._t); try{ localStorage.removeItem(ADRAFT); }catch(e){}
    await auditLoadRows(o.id); auditStatus();
    toast(isNew?"Ο έλεγχος αποθηκεύτηκε":"Ο έλεγχος ενημερώθηκε");
  }catch(e){ toast("Δεν αποθηκεύτηκε: "+(e.message||"σφάλμα")); }
  finally{ btn.disabled=false; }
}
$("auditsheet").addEventListener("change",e=>{
  const t=e.target;
  if(t.dataset&&t.dataset.one&&t.checked) $("auditsheet").querySelectorAll(`input[data-g="${t.dataset.g}"]`).forEach(x=>{ if(x!==t) x.checked=false; });
  auditSaveDraft();
});
$("auditsheet").addEventListener("input",()=>{ clearTimeout(auditSaveDraft._t); auditSaveDraft._t=setTimeout(auditSaveDraft,400); });
$("aprev").onchange=()=>{
  const o=auditPid&&offices.get(auditPid); if(!o) return;
  const r=auditRows.find(x=>x.id===$("aprev").value);
  auditItemId=r?r.id:null; renderAuditSheet(o,r?r.data:null); auditSaveDraft(); auditStatus();
};
$("asave").onclick=auditSave;
$("aprint").onclick=()=>window.print();
$("ablank").onclick=()=>{
  if(auditPid&&!confirm("Θα ανοίξει κενό φύλλο. Ό,τι δεν αποθήκευσες μένει στο πρόχειρο μόνο μέχρι να γράψεις στο κενό. Συνέχεια;")) return;
  auditPid=null; auditItemId=null; $("aq").value=""; $("aprev").innerHTML=`<option value="">Νέος έλεγχος</option>`;
  renderAuditSheet(null,{}); auditStatus(); setTimeout(()=>window.print(),50);
};

/* =================== εισαγωγή =================== */
const HEAD={"επωνυμια":"n","name":"n","τιτλος":"n","κατηγορια":"cat","category":"cat","περιοχη":"a","area":"a","διευθυνση":"ad","address":"ad","νομικημορφη":"f",
  "υπευθυνος":"p","υπευθυνοςιδιοκτητης":"p","τηλεφωνο":"ph","phone":"ph","email":"em","e-mail":"em","website":"w","site":"w","σημειωση":"src","πηγη":"src",
  "αφμ":"afm","vat":"afm","γεμη":"gemi","αργεμη":"gemi","κυκλοςεργασιων":"turnover","turnover":"turnover","τζιρος":"turnover","εργαζομενοι":"employees","employees":"employees","ετος":"finYear"};
function parseCSV(text){
  text=text.replace(/^﻿/,""); const first=(text.split(/\r?\n/)[0]||"");
  const d=(first.split(";").length>=first.split(",").length)?";":",";
  const rows=[]; let row=[], cell="", q=false;
  for(let i=0;i<text.length;i++){ const ch=text[i];
    if(q){ if(ch==='"'){ if(text[i+1]==='"'){cell+='"';i++;} else q=false; } else cell+=ch; }
    else if(ch==='"') q=true; else if(ch===d){ row.push(cell); cell=""; }
    else if(ch==="\n"||ch==="\r"){ if(ch==="\r"&&text[i+1]==="\n") i++; row.push(cell); rows.push(row); row=[]; cell=""; }
    else cell+=ch; }
  if(cell||row.length){ row.push(cell); rows.push(row); }
  const head=(rows.shift()||[]).map(h=>HEAD[normKey(h)]||HEAD[norm(h).replace(/[^a-zα-ω-]/g,"")]||null);
  return rows.map(r=>{ const o={}; r.forEach((v,i)=>{ v=(v||"").trim(); if(head[i]&&v&&v!=="—"&&v!=="Δεν αναφέρεται") o[head[i]]=v; }); return o; }).filter(o=>o.n);
}
let pendingImport=null;
$("imppreview").onclick=async()=>{
  const f=$("impfile").files[0]; if(!f){ toast("Διάλεξε αρχείο"); return; }
  const text=await f.text(); let rows=[], callsIn=[];
  try{
    if(/\.json$/i.test(f.name)){ const d=JSON.parse(text); rows=Array.isArray(d)?d:(d.prospects||d.rows||[]); callsIn=d.calls||[]; }
    else rows=parseCSV(text);
  }catch(e){ $("impres").textContent="Το αρχείο δεν διαβάζεται: "+e.message; return; }
  const defCat=$("impcat").value.trim();
  rows.forEach(r=>{ if(!r.cat&&defCat) r.cat=defCat; ["turnover","employees","finYear"].forEach(k=>{ if(r[k]!=null&&r[k]!==""){ const n=Number(String(r[k]).replace(/\./g,"").replace(",",".").replace(/[^\d.-]/g,"")); r[k]=isNaN(n)?null:n; } }); });
  const exist=new Map(); for(const o of offices.values()) exist.set(dedupKey(o),o);
  let nNew=0,nFill=0,nSame=0; const seen=new Set();
  for(const r of rows){ const k=r.key||dedupKey(r); r._k=dedupKey(r); if(seen.has(r._k)){ nSame++; continue; } seen.add(r._k); exist.has(r._k)?nFill++:nNew++; }
  pendingImport={rows,callsIn,source:f.name};
  $("impres").innerHTML=`Βρέθηκαν <b>${rows.length}</b> γραμμές: <b>${nNew}</b> νέοι υποψήφιοι, <b>${nFill}</b> υπάρχουν ήδη (θα συμπληρωθούν μόνο κενά πεδία)${nSame?`, ${nSame} διπλές μέσα στο αρχείο`:""}${callsIn.length?`, <b>${callsIn.length}</b> κλήσεις από το ιστορικό`:""}.`;
  $("imprun").disabled=!rows.length;
};
const FILLABLE=["cat","a","ad","f","p","em","w","afm","gemi","turnover","employees","finYear"];
$("imprun").onclick=async()=>{
  if(!pendingImport) return; $("imprun").disabled=true;
  const {rows,callsIn,source}=pendingImport; const res=$("impres");
  const exist=new Map(); for(const o of offices.values()) exist.set(dedupKey(o),o);
  const keyToId=new Map(); const creates=[], patches=[]; const seen=new Set();
  for(const r of rows){
    const k=r._k; if(seen.has(k)) { continue; } seen.add(k);
    const ex=exist.get(k);
    if(ex){ keyToId.set(r.key||k, ex.id); const fill={}; for(const f of FILLABLE){ const v=r[f]; if(v!=null&&v!==""&&(ex[f]==null||ex[f]==="")) fill[f]=v; } if(Object.keys(fill).length) patches.push([ex.id,fill]); continue; }
    const data={n:r.n,cat:r.cat||"",a:r.a||"",ad:r.ad||"",f:r.f||"",p:r.p||"",ph:r.ph||"",em:r.em||"",w:r.w||"",src:r.src||("Εισαγωγή από "+source),
      afm:r.afm||"",gemi:r.gemi||"",turnover:r.turnover??null,employees:r.employees??null,finYear:r.finYear??null,finNotes:r.finNotes||"",
      status:r.status&&S[r.status]?r.status:"new",attempts:Number(r.attempts)||0,lastCallAt:toMs(r.lastCallAt),lastBy:r.lastBy||"",callbackAt:toMs(r.callbackAt),
      notes:r.notes||"",hidden:!!r.hidden};
    creates.push([r.key||k,data]);
  }
  try{
    res.textContent=`Δημιουργία ${creates.length} υποψηφίων…`;
    const out=await gBatch(creates.map(([,d])=>({method:"POST",url:`/sites/${site.id}/lists/${lists.p.id}/items`,body:{fields:toPFields(d)}})));
    let ok=0, fail=0; const createdKeys=new Set();
    out.forEach((resp,i)=>{ if(resp&&resp.status<300){ ok++; keyToId.set(creates[i][0], String(resp.body.id)); createdKeys.add(creates[i][0]); } else fail++; });
    res.textContent=`Ενημέρωση ${patches.length} υπαρχόντων…`;
    await gBatch(patches.map(([id,fill])=>({method:"PATCH",url:`/sites/${site.id}/lists/${lists.p.id}/items/${id}/fields`,body:toPFields(fill)})));
    let cOk=0;
    if(callsIn.length){
      res.textContent=`Μεταφορά ${callsIn.length} κλήσεων…`;
      const creq=[]; const nameOf=k=>{ const r=rows.find(x=>(x.key||x._k)===k); return r?r.n:""; };
      for(const c of callsIn){ if(!createdKeys.has(c.key)) continue; const pid=keyToId.get(c.key); if(!pid) continue;
        creq.push({method:"POST",url:`/sites/${site.id}/lists/${lists.c.id}/items`,body:{fields:{Title:nameOf(c.key).slice(0,255),ProspectId:Number(pid),Status:c.s,Note:c.n||"",CallbackAt:toIso(c.cb),CallerName:c.by||"",CallerEmail:"",CallAt:toIso(c.t),Manual:!!c.manual}}}); }
      const co=await gBatch(creq); cOk=co.filter(r=>r&&r.status<300).length;
    }
    await load(); refreshAll(); if(!currentId) nextCall();
    res.innerHTML=`<b>Ολοκληρώθηκε.</b> Νέοι υποψήφιοι: ${ok}${fail?` (απέτυχαν ${fail})`:""} · συμπληρώθηκαν στοιχεία σε ${patches.length} υπάρχοντες${callsIn.length?` · μεταφέρθηκαν ${cOk} κλήσεις`:""}.`;
    pendingImport=null;
  }catch(e){ res.textContent="Σφάλμα: "+e.message; $("imprun").disabled=false; }
};

/* =================== boot =================== */
async function copy(t){ try{ await navigator.clipboard.writeText(t); toast("Αντιγράφηκε: "+t); }catch(e){ toast(t); } }
function refreshAll(){ fillSelects(); renderQueueCount(); renderTodayCallbacks(); if(!$("view-list").hidden) renderList(); if(!$("view-reports").hidden) renderReports(); if(!$("view-calendar").hidden) renderCalendar(); checkDue(); }
async function poll(){ if(document.hidden||!$("drawer").hidden||pendingImport) return; try{ await load(); refreshAll(); if(currentId&&offices.has(currentId)) renderHistory(offices.get(currentId)); else if(!currentId&&!$("view-calls").hidden) nextCall(); }catch(e){} }

async function boot(){
  if(!CFG.tenantId||!CFG.clientId||!CFG.siteUrl){ $("setup").hidden=false; $("cf-tenant").value=CFG.tenantId; $("cf-client").value=CFG.clientId; $("cf-site").value=CFG.siteUrl; return; }
  pca=new msal.PublicClientApplication({auth:{clientId:CFG.clientId,authority:`https://login.microsoftonline.com/${CFG.tenantId}`,redirectUri:location.origin+location.pathname},cache:{cacheLocation:"localStorage"}});
  await pca.initialize();
  try{ const r=await pca.handleRedirectPromise(); if(r&&r.account) pca.setActiveAccount(r.account); }catch(e){ $("lerr").textContent=e.message; }
  account=pca.getActiveAccount()||pca.getAllAccounts()[0];
  if(!account){ $("login").hidden=false; return; }
  pca.setActiveAccount(account);
  me={name:account.name||account.username, email:(account.username||"").toLowerCase()};
  $("app").hidden=false; $("who").textContent=me.name;
  try{
    await resolveSite();
    lists.p=await findList(CFG.prospectsList); lists.c=await findList(CFG.callsList);
    if(!lists.p||!lists.c){
      VIEWS.forEach(x=>$("view-"+x).hidden=true); $("view-setup-lists").hidden=false;
      $("ln1").textContent=CFG.prospectsList; $("ln2").textContent=CFG.callsList; $("sitename").textContent=site.webUrl;
      $("btn-create-lists").onclick=async()=>{ try{ await createLists(); location.reload(); }catch(e){ $("setupmsg").textContent="Δεν δημιουργήθηκαν: "+e.message; } };
      return;
    }
    await load();
    try{ auditList=await findList(CFG.auditsList); }catch(e){ auditList=null; }
  }catch(e){
    banner(e.status===403||e.status===401?"Δεν έχεις πρόσβαση στο SharePoint site ή δεν δόθηκε συναίνεση στην εφαρμογή. Δες το README (βήμα 1-2).":(e.status===404?"Δεν βρέθηκε το SharePoint site: "+CFG.siteUrl:"Σφάλμα σύνδεσης: "+e.message));
    return;
  }
  ensureColumns();
  fillSelects();
  const h=(location.hash||"").slice(1), m=h.match(/^open=(\d+)$/);
  await checkCalendar();
  if(m&&offices.has(m[1])){ await openItem(m[1]); }
  else { show(h||"calls"); if(!$("view-calls").hidden&&!currentId) await nextCall(); }
  renderTodayCallbacks(); checkDue();
  setInterval(checkDue,30000);
  window.addEventListener("hashchange",()=>{ const mm=(location.hash||"").match(/^#open=(\d+)$/); if(mm&&offices.has(mm[1])) openItem(mm[1]); });
  setInterval(poll,POLL_MS);
}
boot();
})();
