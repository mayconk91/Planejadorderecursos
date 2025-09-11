// ===== Utilidades de data =====
function toYMD(d){const y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,'0'),day=String(d.getDate()).padStart(2,'0');return `${y}-${m}-${day}`}
function fromYMD(s){return new Date(`${s}T00:00:00`)}
function addDays(d,n){const nd=new Date(d);nd.setDate(nd.getDate()+n);return nd}
function diffDays(a,b){const A=new Date(a.getFullYear(),a.getMonth(),a.getDate());const B=new Date(b.getFullYear(),b.getMonth(),b.getDate());return Math.round((A-B)/(1000*60*60*24))}
function clampDate(d){return new Date(d.getFullYear(),d.getMonth(),d.getDate())}
function uuid(){if (crypto && crypto.randomUUID) return crypto.randomUUID(); const s=()=>Math.floor((1+Math.random())*0x10000).toString(16).substring(1); return `${s()}${s()}-${s()}-${s()}-${s()}-${s()}${s()}${s()}`}

// ===== Domínio =====
const TIPOS=["Interno","Externo"];
const SENIORIDADES=["Jr","Pl","Sr","NA"];
const STATUS=["Planejada","Em Execução","Bloqueada","Concluída","Cancelada"];

// ===== Persistência =====
const LS={res:"rp_resources_v2",act:"rp_activities_v2",trail:"rp_trail_v1",user:"rp_user_v1"};
function loadLS(k,fallback){try{const raw=localStorage.getItem(k);return raw?JSON.parse(raw):fallback}catch{return fallback}}
function saveLS(k,v){localStorage.setItem(k,JSON.stringify(v))}

// ===== IndexedDB helpers (tiny) =====
function idbProm(req){return new Promise((res,rej)=>{req.onsuccess=()=>res(req.result); req.onerror=()=>rej(req.error);})}
function idbOpen(name, store){
  return new Promise((resolve,reject)=>{
    const req=indexedDB.open(name,1);
    req.onupgradeneeded=()=>{const db=req.result; if(!db.objectStoreNames.contains(store)) db.createObjectStore(store);};
    req.onsuccess=()=>resolve(req.result); req.onerror=()=>reject(req.error);
  });
}
async function idbSet(dbname, store, key, value){
  const db=await idbOpen(dbname, store);
  const tx=db.transaction(store,"readwrite"); const st=tx.objectStore(store);
  const req=st.put(value, key); await idbProm(req); await idbProm(tx.done||tx); db.close();
}
async function idbGet(dbname, store, key){
  const db=await idbOpen(dbname, store);
  const tx=db.transaction(store,"readonly"); const st=tx.objectStore(store);
  const req=st.get(key); const val=await idbProm(req); await idbProm(tx.done||tx); db.close(); return val;
}

// ===== File System Access (Salvar na pasta) =====
const DATAFILES = {
  [LS.res]: "resources.json",
  [LS.act]: "activities.json",
  [LS.trail]: "trails.json"
};
const FSA_DB="planner_fs";
const FSA_STORE="handles";
let dirHandle=null;

async function hasFSA(){ return 'showDirectoryPicker' in window; }

async function fsaLoadHandle(){
  try{
    const h=await idbGet(FSA_DB,FSA_STORE,'dir');
    if(!h) return null;
    const perm = await h.queryPermission({mode:'readwrite'});
    if(perm==='granted') return h;
    const perm2 = await h.requestPermission({mode:'readwrite'});
    return perm2==='granted'?h:null;
  }catch(e){ console.warn('fsaLoadHandle error',e); return null; }
}

async function fsaPickFolder(){
  try{
    const h=await window.showDirectoryPicker();
    await idbSet(FSA_DB,FSA_STORE,'dir',h);
    dirHandle=h;
    updateFolderStatus();
    alert('Pasta definida: '+(h.name||'(sem nome)'));
  }catch(e){ if(e&&e.name!=='AbortError') alert('Não foi possível selecionar a pasta: '+e.message); }
}

async function writeFile(handle, name, content){
  const fhandle = await handle.getFileHandle(name, {create:true});
  const ws = await fhandle.createWritable();
  await ws.write(content);
  await ws.close();
}

async function readFile(handle, name){
  const fhandle = await handle.getFileHandle(name, {create:false});
  const file = await fhandle.getFile();
  const text = await file.text();
  return text;
}

async function saveAllToFolder(){
  if(!dirHandle) return false;
  try{
    await writeFile(dirHandle, DATAFILES[LS.res], JSON.stringify(resources,null,2));
    await writeFile(dirHandle, DATAFILES[LS.act], JSON.stringify(activities,null,2));
    await writeFile(dirHandle, DATAFILES[LS.trail], JSON.stringify(trails,null,2));
    updateFolderStatus('Salvo em '+new Date().toLocaleTimeString());
    return true;
  }catch(e){ console.error(e); alert('Falha ao salvar na pasta: '+e.message); return false; }
}

async function loadAllFromFolder(){
  if(!dirHandle) return false;
  try{
    const rtxt = await readFile(dirHandle, DATAFILES[LS.res]).catch(e=>{ if(e && e.name==='NotFoundError') return '[]'; else throw e; });
    const atxt = await readFile(dirHandle, DATAFILES[LS.act]).catch(e=>{ if(e && e.name==='NotFoundError') return '[]'; else throw e; });
    const ttxt = await readFile(dirHandle, DATAFILES[LS.trail]).catch(e=>{ if(e && e.name==='NotFoundError') return '{}'; else throw e; });
    const r = JSON.parse(rtxt); const a = JSON.parse(atxt); const t = JSON.parse(ttxt);
    if(Array.isArray(r)&&Array.isArray(a)&&t&&typeof t==='object'){
      resources=r; activities=a; trails=t;
      saveLS(LS.res,resources); saveLS(LS.act,activities); saveLS(LS.trail,trails);
      renderAll();
      updateFolderStatus('Carregado da pasta');
      return true;
    } else { alert('Arquivos inválidos na pasta.'); return false; }
  }catch(e){ console.error(e); alert('Falha ao carregar da pasta: '+e.message); return false; }
}

function updateFolderStatus(extra){
  const el=document.getElementById('folderStatus');
  if(!el) return;
  if(!dirHandle){ el.textContent='(nenhuma pasta definida)'; return; }
  el.textContent='Pasta: '+(dirHandle.name||'(sem nome)') + (extra? ' — '+extra : '');
}

// Hook nos botões
const btnPickFolder=document.getElementById('btnPickFolder');
const btnSaveNow=document.getElementById('btnSaveNow');
const btnReloadFromFolder=document.getElementById('btnReloadFromFolder');
if(btnPickFolder) btnPickFolder.onclick=()=>fsaPickFolder();
if(btnSaveNow) btnSaveNow.onclick=()=>saveAllToFolder();
if(btnReloadFromFolder) btnReloadFromFolder.onclick=()=>loadAllFromFolder();

// Na inicialização, tenta recuperar a pasta e carregar (sem bloquear)
(async()=>{
  if(await hasFSA()){
    dirHandle = await fsaLoadHandle();
    updateFolderStatus();
    if(dirHandle){
      try{ await loadAllFromFolder(); }catch{ /* ignore */ }
    }
  } else {
    const el=document.getElementById('folderStatus'); if(el) el.textContent='(navegador sem suporte de salvar em pasta — usando armazenamento do navegador)';
  }
})();

// Redirecionar persistência: salva em LS e também espelha para a pasta (best-effort)
const _saveLS_orig = saveLS;
saveLS = function(k,v){
  _saveLS_orig(k,v);
  if(dirHandle) { try{ saveAllToFolder(); }catch{} }
};


// ===== Dados iniciais =====
const today=clampDate(new Date());
const sampleResources=[
  {id:uuid(),nome:"Rafael",tipo:"Interno",senioridade:"Sr",ativo:true,capacidade:100,inicioAtivo:"",fimAtivo:""},
  {id:uuid(),nome:"Anna",tipo:"Externo",senioridade:"Pl",ativo:true,capacidade:100,inicioAtivo:"",fimAtivo:""},
  {id:uuid(),nome:"Maycon",tipo:"Interno",senioridade:"Sr",ativo:true,capacidade:50,inicioAtivo:"",fimAtivo:""} // exemplo part-time
];
const sampleActivities=[
  {id:uuid(),titulo:"Implantação Bula Digital",resourceId:sampleResources[2].id,inicio:toYMD(addDays(today,-5)),fim:toYMD(addDays(today,6)),status:"Em Execução",alocacao:100},
  {id:uuid(),titulo:"Quarentena — Ajustes",resourceId:sampleResources[0].id,inicio:toYMD(addDays(today,1)),fim:toYMD(addDays(today,10)),status:"Planejada",alocacao:50},
  {id:uuid(),titulo:"Auditoria Interna",resourceId:sampleResources[1].id,inicio:toYMD(addDays(today,8)),fim:toYMD(addDays(today,15)),status:"Planejada",alocacao:100},
];

let resources=loadLS(LS.res,sampleResources);
let activities=loadLS(LS.act,sampleActivities);
let trails=loadLS(LS.trail,{}); // { [atividadeId]: [{ts, oldInicio, oldFim, newInicio, newFim, justificativa, user}] }
function saveTrails(){ saveLS(LS.trail, trails); }
function addTrail(atividadeId, entry){
  if(!trails[atividadeId]) trails[atividadeId]=[];
  trails[atividadeId].push(entry);
  saveTrails();
}

// ===== Estado dos filtros =====
const selectedStatus=new Set(STATUS); // por padrão, todos ativos
let filtroTipo="";
let filtroSenioridade="";
let buscaTitulo="";
let buscaRecurso="";

let rangeStart=toYMD(addDays(today,-14));
let rangeEnd=toYMD(addDays(today,60));

// ===== UI refs =====
const statusChips=document.getElementById("statusChips");
const tipoSel=document.getElementById("tipoSel");
const senioridadeSel=document.getElementById("senioridadeSel");
const buscaTituloInput=document.getElementById("buscaTitulo");
const buscaRecursoInput=document.getElementById("buscaRecurso");
const inicioVisao=document.getElementById("inicioVisao");
const fimVisao=document.getElementById("fimVisao");

const tblRecursos=document.querySelector("#tblRecursos tbody");
const tblAtividades=document.querySelector("#tblAtividades tbody");
const gantt=document.getElementById("gantt");

const dlgRecurso=document.getElementById("dlgRecurso");
const formRecurso=document.getElementById("formRecurso");
const dlgRecursoTitulo=document.getElementById("dlgRecursoTitulo");

const dlgAtividade=document.getElementById("dlgAtividade");
const formAtividade=document.getElementById("formAtividade");
const dlgAtividadeTitulo=document.getElementById("dlgAtividadeTitulo");

const dlgJust=document.getElementById("dlgJust");
const formJust=document.getElementById("formJust");
const justResumo=document.getElementById("justResumo");
const btnJustConfirm=document.getElementById("btnJustConfirm");

const dlgHist=document.getElementById("dlgHist");
const histList=document.getElementById("histList");
const btnHistExport=document.getElementById("btnHistExport");
let histCurrentId=null;

const currentUserInput=document.getElementById("currentUser");
const btnHistAll=document.getElementById("btnHistAll");
const btnBackup=document.getElementById("btnBackup");
const fileRestore=document.getElementById("fileRestore");
const tooltip=document.getElementById("tooltip");
const aggGran=document.getElementById("aggGran");
const aggCharts=document.getElementById("aggCharts");

// Usuário atual (opcional) para trilha
let currentUser = loadLS(LS.user, "");
if(currentUserInput){ currentUserInput.value=currentUser; currentUserInput.oninput=()=>{ currentUser=currentUserInput.value.trim(); saveLS(LS.user,currentUser); }; }


// ===== Abas (tabs) =====
function activateTab(name){
  document.querySelectorAll('.tab').forEach(b=>b.classList.toggle('active', b.dataset.tab===name));
  document.querySelectorAll('.tabpanel').forEach(p=>p.classList.toggle('active', p.id==='tab-'+name));
}
document.addEventListener('click', (ev)=>{
  const b=ev.target.closest('.tab'); if(!b) return;
  activateTab(b.dataset.tab);
});

// ===== Chips de status =====
function renderStatusChips(){
  statusChips.innerHTML="";
  STATUS.forEach(s=>{
    const span=document.createElement("span");
    span.className="chip"+(selectedStatus.has(s)?" active":"");
    span.textContent=s;
    span.onclick=()=>{ if(selectedStatus.has(s)) selectedStatus.delete(s); else selectedStatus.add(s); renderStatusChips(); renderAll(); };
    statusChips.appendChild(span);
  });
}

// ===== Filtros =====
tipoSel.onchange=()=>{filtroTipo=tipoSel.value; renderAll();};
senioridadeSel.onchange=()=>{filtroSenioridade=senioridadeSel.value; renderAll();};
buscaTituloInput.oninput=()=>{buscaTitulo=buscaTituloInput.value.toLowerCase(); renderAll();};
if(buscaRecursoInput){
 buscaRecursoInput.oninput=()=>{
    buscaRecurso=buscaRecursoInput.value.toLowerCase().trim();
    renderAll();
  };
}

// ===== Range =====
inicioVisao.value=rangeStart; fimVisao.value=rangeEnd;
inicioVisao.onchange=()=>{rangeStart=inicioVisao.value; renderAll();};
fimVisao.onchange=()=>{rangeEnd=fimVisao.value; renderAll();};

// ===== CRUD Recurso =====
document.getElementById("btnNovoRecurso").onclick=()=>{
  dlgRecursoTitulo.textContent="Novo Recurso";
  formRecurso.reset();
  formRecurso.elements["id"].value="";
  formRecurso.elements["capacidade"].value=100;
  dlgRecurso.showModal();
};
/* submit default allowed (dialog will close on cancel) */
document.getElementById("btnSalvarRecurso").onclick=()=>{
  const f=formRecurso.elements;
  const rec={
    id:f["id"].value||uuid(),
    nome:f["nome"].value.trim(),
    tipo:f["tipo"].value,
    senioridade:f["senioridade"].value,
    ativo:f["ativo"].checked,
    capacidade:Math.max(1,Number(f["capacidade"].value||100)),
    inicioAtivo:f["inicioAtivo"].value||"",
    fimAtivo:f["fimAtivo"].value||""
  };
  if(!rec.nome){alert("Informe o nome.");return}
  const idx=resources.findIndex(r=>r.id===rec.id);
  if(idx>=0) resources[idx]=rec; else resources.push(rec);
  saveLS(LS.res,resources);
  dlgRecurso.close();
  renderAll();
};

// ===== CRUD Atividade =====
document.getElementById("btnNovaAtividade").onclick=()=>{
  dlgAtividadeTitulo.textContent="Nova Atividade";
  formAtividade.reset();
  fillRecursoOptions();
  formAtividade.elements["id"].value="";
  formAtividade.elements["inicio"].value=toYMD(today);
  formAtividade.elements["fim"].value=toYMD(addDays(today,5));
  formAtividade.elements["alocacao"].value=100;
  dlgAtividade.showModal();
};
function fillRecursoOptions(){
  const sel=formAtividade.elements["resourceId"];
  sel.innerHTML="";
  resources.forEach(r=>{
    if(!r.ativo) return;
    const opt=document.createElement("option");
    opt.value=r.id; opt.textContent=r.nome;
    sel.appendChild(opt);
  });
}
/* submit default allowed (dialog will close on cancel) */
document.getElementById("btnSalvarAtividade").onclick=()=>{
  const f=formAtividade.elements;
  const at={
    id:f["id"].value||uuid(),
    titulo:f["titulo"].value.trim(),
    resourceId:f["resourceId"].value,
    inicio:f["inicio"].value,
    fim:f["fim"].value,
    status:f["status"].value,
    alocacao:Math.max(1,Number(f["alocacao"].value||100))
  };
  if(!at.titulo) return alert("Informe o título.");
  if(!at.resourceId) return alert("Selecione o recurso.");
  if(fromYMD(at.fim)<fromYMD(at.inicio)) return alert("Fim não pode ser menor que início.");

  // valida janela ativa do recurso, se houver
  const rec=resources.find(x=>x.id===at.resourceId);
  if(rec){
    if(rec.inicioAtivo && fromYMD(at.inicio) < fromYMD(rec.inicioAtivo)){
      return alert(`Início da atividade (${at.inicio}) menor que início ativo do recurso (${rec.inicioAtivo}).`);
    }
    if(rec.fimAtivo && fromYMD(at.fim) > fromYMD(rec.fimAtivo)){
      return alert(`Fim da atividade (${at.fim}) maior que fim ativo do recurso (${rec.fimAtivo}).`);
    }
  }
  // alerta de sobrealocação >100%
  let over=false;
  const start=fromYMD(at.inicio), end=fromYMD(at.fim);
  for(let d=new Date(start); d<=end; d=addDays(d,1)){
    const sum = activities.filter(x=>x.id!==at.id && x.resourceId===at.resourceId && fromYMD(x.inicio)<=d && d<=fromYMD(x.fim))
                          .reduce((acc,x)=>acc+(x.alocacao||100),0) + (at.alocacao||100);
    const cap = rec? (rec.capacidade||100) : 100;
    if(sum>cap){ over=true; break; }
  }
  if(over && !confirm("Aviso: esta alteração resultará em sobrealocação (>100%) em pelo menos um dia. Deseja continuar?")) return;

  const idx=activities.findIndex(a=>a.id===at.id);
  if(idx>=0){
    const prev=activities[idx];
    const mudouDatas = prev.inicio!==at.inicio || prev.fim!==at.fim;
    if(mudouDatas){
      // abrir justificativa
      window.__pendingAt=at;
      window.__pendingIdx=idx;
      justResumo.textContent=`${prev.titulo} — Início: ${prev.inicio} → ${at.inicio} | Fim: ${prev.fim} → ${at.fim}`;
      formJust.elements["just"].value="";
      dlgJust.showModal();
      return; // continua após confirmar justificativa
    } else {
      activities[idx]=at;
      saveLS(LS.act,activities);
      dlgAtividade.close();
      renderAll();
    }
  } else {
    // novo registro (não exige justificativa para criação)
    activities.push(at);
    saveLS(LS.act,activities);
    dlgAtividade.close();
    renderAll();
  }
};

// Justificativa confirm
btnJustConfirm.onclick=(e)=>{
  e.preventDefault();
  const txt=formJust.elements["just"].value.trim();
  if(!txt){ alert("Informe a justificativa."); return; }
  const at=window.__pendingAt;
  const idx=window.__pendingIdx;
  if(at==null || idx==null){ dlgJust.close(); return; }
  const prev=activities[idx];
  addTrail(at.id, {
    ts: new Date().toISOString(),
    oldInicio: prev.inicio, oldFim: prev.fim,
    newInicio: at.inicio, newFim: at.fim,
    justificativa: txt,
    user: currentUser||""
  });
  activities[idx]=at;
  saveLS(LS.act,activities);
  dlgJust.close();
  dlgAtividade.close();
  renderAll();
};

// ===== Tabelas =====
function renderTables(filteredActs){
  // Recursos
  tblRecursos.innerHTML="";
  resources.forEach(r=>{
    const tr=document.createElement("tr");
    tr.innerHTML=`
      <td>${r.nome}</td>
      <td>${r.tipo}</td>
      <td>${r.senioridade}</td>
      <td>${r.ativo?"Sim":"Não"}</td>
      <td>${r.capacidade}</td>
      <td>${r.inicioAtivo||""}</td>
      <td>${r.fimAtivo||""}</td>
      <td class="actions">
        <button class="btn">Editar</button>
        <button class="btn danger">Excluir</button>
      </td>`;
    tr.querySelectorAll("button")[0].onclick=()=>{
      dlgRecursoTitulo.textContent="Editar Recurso";
      formRecurso.elements["id"].value=r.id;
      formRecurso.elements["nome"].value=r.nome;
      formRecurso.elements["tipo"].value=r.tipo;
      formRecurso.elements["senioridade"].value=r.senioridade;
      formRecurso.elements["ativo"].checked=!!r.ativo;
      formRecurso.elements["capacidade"].value=r.capacidade||100;
      formRecurso.elements["inicioAtivo"].value=r.inicioAtivo||"";
      formRecurso.elements["fimAtivo"].value=r.fimAtivo||"";
      dlgRecurso.showModal();
    };
    tr.querySelectorAll("button")[1].onclick=()=>{
      if(!confirm("Remover recurso e suas alocações?")) return;
      resources=resources.filter(x=>x.id!==r.id);
      activities=activities.filter(a=>a.resourceId!==r.id);
      saveLS(LS.res,resources); saveLS(LS.act,activities);
      renderAll();
    };
    tblRecursos.appendChild(tr);
  });

  // Atividades
  tblAtividades.innerHTML="";
  filteredActs.forEach(a=>{
    const r=resources.find(x=>x.id===a.resourceId);
    const tr=document.createElement("tr");
    tr.innerHTML=`
      <td>${a.titulo}</td>
      <td>${r? r.nome:"—"}</td>
      <td>${a.inicio}</td>
      <td>${a.fim}</td>
      <td>${a.status}</td>
      <td>${a.alocacao||100}</td>
      <td class="actions">
        <button class="btn">Editar</button>
        <button class="btn">Histórico</button>
        <button class="btn danger">Excluir</button>
      </td>`;
    tr.querySelectorAll("button")[0].onclick=()=>{
      dlgAtividadeTitulo.textContent="Editar Atividade";
      fillRecursoOptions();
      formAtividade.elements["id"].value=a.id;
      formAtividade.elements["titulo"].value=a.titulo;
      formAtividade.elements["resourceId"].value=a.resourceId;
      formAtividade.elements["inicio"].value=a.inicio;
      formAtividade.elements["fim"].value=a.fim;
      formAtividade.elements["status"].value=a.status;
      formAtividade.elements["alocacao"].value=a.alocacao||100;
      dlgAtividade.showModal();
    };
    tr.querySelectorAll("button")[1].onclick=()=>{
      // abrir histórico
      histCurrentId=a.id;
      const list = trails[a.id]||[];
      if(list.length===0){
        histList.innerHTML='<div class="muted">Sem alterações de datas registradas para esta atividade.</div>';
      }else{
        const s=document.getElementById('histStart').value;
        const e=document.getElementById('histEnd').value;
        const rows=list.slice().reverse().filter(it=>{
          const t=new Date(it.ts);
          return (!s || t>=fromYMD(s)) && (!e || t<=addDays(fromYMD(e),0));
        }).map(it=>{
          return `<div style="padding:6px 8px; background:#fff; border:1px solid #e2e8f0; border-radius:8px; margin:6px 0">
              <div style="font-size:12px;color:#475569">${new Date(it.ts).toLocaleString()}${it.user? ' • ' + it.user : ''}</div>
              <div style="font-weight:600">Início: ${it.oldInicio} → ${it.newInicio} | Fim: ${it.oldFim} → ${it.newFim}</div>
              <div style="margin-top:4px">${it.justificativa? it.justificativa.replace(/</g,'&lt;') : ''}</div>
            </div>`;
        }).join("");
        histList.innerHTML=rows || '<div class="muted">Sem registros no período.</div>';
      }
      dlgHist.showModal();
      const btn=document.getElementById('histApply'); if(btn){ btn.onclick=()=>{ tr.querySelectorAll('button')[1].onclick(); }; }
    };
    tr.querySelectorAll("button")[2].onclick=()=>{
      if(!confirm("Remover atividade?")) return;
      activities=activities.filter(x=>x.id!==a.id);
      saveLS(LS.act,activities);
      renderAll();
    };
    tblAtividades.appendChild(tr);
  });
}

// ===== Interseção/intervalos =====
function mergeIntervals(intervals){
  if(!intervals.length) return [];
  const sorted=[...intervals].sort((a,b)=>fromYMD(a.inicio)-fromYMD(b.inicio));
  const res=[sorted[0]];
  for(let i=1;i<sorted.length;i++){
    const prev=res[res.length-1];
    const cur=sorted[i];
    const prevEnd=fromYMD(prev.fim);
    const curStart=fromYMD(cur.inicio);
    if(curStart<=addDays(prevEnd,1)){
      const newEnd=fromYMD(cur.fim)>prevEnd?cur.fim:prev.fim;
      res[res.length-1]={inicio:prev.inicio,fim:newEnd};
    }else res.push(cur);
  }
  return res;
}
function invertIntervals(intervals,startYMD,endYMD){
  const free=[]; let cursor=startYMD;
  const sDate=fromYMD(startYMD); const eDate=fromYMD(endYMD);
  const merged=mergeIntervals(intervals);
  for(const intv of merged){
    const iStart=fromYMD(intv.inicio);
    if(iStart>sDate && fromYMD(cursor)<iStart){
      free.push({inicio:cursor,fim:toYMD(addDays(iStart,-1))});
    }
    const iEnd=fromYMD(intv.fim);
    cursor=toYMD(addDays(iEnd,1));
  }
  if(fromYMD(cursor)<=eDate) free.push({inicio:cursor,fim:endYMD});
  return free;
}

// ===== Gantt =====
function statusClass(s){
  switch(s){
    case "Planejada": return "planejada";
    case "Em Execução": return "execucao";
    case "Bloqueada": return "bloqueada";
    case "Concluída": return "concluida";
    case "Cancelada": return "cancelada";
    default: return "";
  }
}
function buildDays(){
  const start=fromYMD(rangeStart), end=fromYMD(rangeEnd);
  const out=[]; for(let d=new Date(start); d<=end; d=addDays(d,1)) out.push(new Date(d));
  return out;
}

function renderGantt(filteredActs){
  gantt.innerHTML="";
  const days=buildDays();
  // Header
  const header=document.createElement("div");
  header.className="header";
  const left=document.createElement("div");
  left.className="col-fixed";
  left.innerHTML=`<div class="muted" style="font-size:12px;font-weight:600">RECURSO</div>`;
  const right=document.createElement("div");
  right.className="col-grid";

  const gridDays=document.createElement("div");
  gridDays.className="grid-days";
  gridDays.style.gridTemplateColumns=`repeat(${days.length}, 28px)`;
  // Linha meses
  const rowMonths=document.createElement("div");
  rowMonths.className="row-months";
  rowMonths.style.display="grid";
  rowMonths.style.gridTemplateColumns=`repeat(${days.length}, 28px)`;
  days.forEach((d,i)=>{
    const isFirstOfMonth=d.getDate()===1 || i===0;
    const cell=document.createElement("div");
    cell.className="cell-day";
    cell.style.fontWeight=isFirstOfMonth?"600":"400";
    cell.textContent=isFirstOfMonth? d.toLocaleDateString(undefined,{month:"short",year:"2-digit"}):"";
    rowMonths.appendChild(cell);
  });
  // Linha dias
  const rowDays=document.createElement("div");
  rowDays.style.display="grid";
  rowDays.style.gridTemplateColumns=`repeat(${days.length}, 28px)`;
  days.forEach(d=>{
    const cell=document.createElement("div");
    cell.className="cell-day";
    cell.textContent=String(d.getDate()).padStart(2,"0");
    rowDays.appendChild(cell);
  });
  right.appendChild(rowMonths); right.appendChild(rowDays);
  header.appendChild(left); header.appendChild(right);
  gantt.appendChild(header);

  // Mapa: atividades por recurso (já filtradas)
  const byRes=Object.fromEntries(resources.map(r=>[r.id,[]]));
  filteredActs.forEach(a=>{ if(byRes[a.resourceId]) byRes[a.resourceId].push(a); });
  Object.keys(byRes).forEach(k=>byRes[k].sort((a,b)=>fromYMD(a.inicio)-fromYMD(b.inicio)));

  // Render rows
  resources.forEach(r=>{
    // filtros
    if(filtroTipo && r.tipo!==filtroTipo) return;
    if(filtroSenioridade && r.senioridade!==filtroSenioridade) return;
    if(!r.ativo) return;
    const acts=byRes[r.id]||[];

    const row=document.createElement("div");
    row.className="row";

    const info=document.createElement("div");
    info.className="info";
    info.innerHTML=`<div style="font-weight:600">${r.nome}</div>
      <div class="muted" style="font-size:12px">${r.tipo} • ${r.senioridade} • Cap: ${r.capacidade}%${(r.inicioAtivo||r.fimAtivo)? " • Janela: " + (r.inicioAtivo||"…") + " → " + (r.fimAtivo||"…") : ""}</div>`;

    const bargrid=document.createElement("div");
    bargrid.className="bargrid";

    // Heatmap de capacidade por dia + tooltip + contador concorrência
    const cap=r.capacidade||100;
    days.forEach((d,i)=>{
      const dy=toYMD(d);
      const activeActs = acts.filter(a=>fromYMD(a.inicio)<=d && d<=fromYMD(a.fim));
      const sum=activeActs.reduce((acc,a)=>acc+(a.alocacao||100),0);
      const perc=cap? (sum/cap)*100 : 0;
      const heat=document.createElement("div");
      heat.className="heatcell";
      heat.style.left=`${i*28}px`; heat.style.width="28px";
      if(perc>100) heat.classList.add("heat-over");
      else if(perc>0) heat.classList.add(perc>70?"heat-high":"heat-ok");
      heat.onmouseenter=(ev)=>{
        const rows = activeActs.map(a=>`<div class="t-row"><strong>${a.titulo}</strong> — ${a.alocacao||100}% (${a.status})</div>`).join("");
        tooltip.innerHTML = `<div class="t-title">${r.nome} — ${dy}</div><div class="muted">Ocupação: ${Math.round(perc)}% (cap ${cap}%) • Concorrência: ${activeActs.length}</div>${rows}`;
        tooltip.classList.remove("hidden");
      };
      heat.onmousemove=(ev)=>{ tooltip.style.left = (ev.clientX+12)+"px"; tooltip.style.top = (ev.clientY+12)+"px"; };
      heat.onmouseleave=()=>{ tooltip.classList.add("hidden"); };
      bargrid.appendChild(heat);
      // contador concorrência
      const c=activeActs.length;
      if(c>1){
        const cc=document.createElement("div");
        cc.className="ccell "+(c>=4?"high":(c>=3?"med":"low"));
        cc.style.left=`${i*28}px`; cc.style.width="28px";
        cc.textContent=String(c);
        cc.title=`${c} atividades simultâneas`;
        bargrid.appendChild(cc);
      }
    });

    // Lanes (pistas) para empilhar atividades concorrentes
    const daysLen = days.length;
    const startBase = fromYMD(rangeStart);
    function dayIndex(ymd){
      return Math.max(0, Math.min(daysLen-1, diffDays(fromYMD(ymd), startBase)));
    }
    const intervals = acts.map(a=>{
      const sIdx = dayIndex(a.inicio);
      const eIdx = dayIndex(a.fim);
      return {a, sIdx, eIdx};
    }).filter(iv=>iv.eIdx>=0 && iv.sIdx<=daysLen-1);
    const lanes = []; // greedy
    const placed = intervals.sort((x,y)=>x.sIdx - y.sIdx).map(iv=>{
      let lane = 0;
      while(lane < lanes.length && !(lanes[lane] < iv.sIdx - 0)) lane++;
      if(lane === lanes.length) lanes.push(-Infinity);
      lanes[lane] = iv.eIdx;
      return {...iv, lane};
    });

    // Gaps
    const busy=acts.map(a=>({inicio:toYMD(new Date(Math.max(fromYMD(a.inicio),fromYMD(rangeStart)))),
                             fim:toYMD(new Date(Math.min(fromYMD(a.fim),fromYMD(rangeEnd))))}))
                  .filter(x=>fromYMD(x.inicio)<=fromYMD(x.fim));
    const gaps=invertIntervals(busy,rangeStart,rangeEnd);
    gaps.forEach(g=>{
      const startIdx=Math.max(0,diffDays(fromYMD(g.inicio),fromYMD(rangeStart)));
      const endIdx=Math.min(days.length-1,diffDays(fromYMD(g.fim),fromYMD(rangeStart)));
      const el=document.createElement("div");
      el.className="gapblock";
      el.style.left=`${startIdx*28}px`;
      el.style.width=`${(endIdx-startIdx+1)*28}px`;
      el.title=`Lacuna: ${g.inicio} → ${g.fim}`;
      bargrid.appendChild(el);
    });

    // Atividades empilhadas
    placed.forEach(p=>{
      const a=p.a;
      if(!selectedStatus.has((a.status||"").trim())) return;
      const aStart=Math.max(0,p.sIdx);
      const aEnd=Math.min(days.length-1,p.eIdx);
      const b=document.createElement("div");
      b.className=`activity ${statusClass(a.status)}`;
      b.style.left=`${aStart*28}px`;
      b.style.width=`${(aEnd-aStart+1)*28}px`;
      b.style.top=`${p.lane*22 + 2}px`;
      b.textContent=a.titulo;
      b.title=`${a.titulo} — ${a.inicio} → ${a.fim} • ${a.status} • ${a.alocacao||100}%`;
      b.onclick=()=>{
        dlgAtividadeTitulo.textContent="Editar Atividade";
        fillRecursoOptions();
        formAtividade.elements["id"].value=a.id;
        formAtividade.elements["titulo"].value=a.titulo;
        formAtividade.elements["resourceId"].value=a.resourceId;
        formAtividade.elements["inicio"].value=a.inicio;
        formAtividade.elements["fim"].value=a.fim;
        formAtividade.elements["status"].value=a.status;
        formAtividade.elements["alocacao"].value=a.alocacao||100;
        dlgAtividade.showModal();
      };
      bargrid.appendChild(b);
    });

    // altura da linha conforme número de lanes (mínimo 42px)
    const lanesH = Math.max(42, lanes.length*22 + 6);
    bargrid.style.height = lanesH + "px";

    // linhas de grade verticais
    const gridBg=document.createElement("div");
    gridBg.style.position="absolute"; gridBg.style.top="0"; gridBg.style.bottom="0"; gridBg.style.left="0"; gridBg.style.right="0";
    gridBg.style.display="grid"; gridBg.style.gridTemplateColumns=`repeat(${days.length}, 28px)`; gridBg.style.pointerEvents="none";
    for(let i=0;i<days.length;i++){
      const v=document.createElement("div");
      v.style.borderLeft="1px solid #f1f5f9";
      gridBg.appendChild(v);
    }
    bargrid.appendChild(gridBg);

    row.appendChild(info); row.appendChild(bargrid);
    gantt.appendChild(row);
  });
}

// ===== Filtragem de atividades =====
function getFilteredActivities(){
  return activities.filter(a=>{
    if(!selectedStatus.has((a.status||"").trim())) return false;
    const r=resources.find(x=>x.id===a.resourceId);
    if(!r) return false;
    if(filtroTipo && r.tipo!==filtroTipo) return false;
    if(filtroSenioridade && r.senioridade!==filtroSenioridade) return false;
    if(buscaRecurso && !(r.nome||"").toLowerCase().includes(buscaRecurso)) return false;
    if(buscaTitulo && !a.titulo.toLowerCase().includes(buscaTitulo)) return false;
    const s=fromYMD(a.inicio), e=fromYMD(a.fim);
    if(e<fromYMD(rangeStart) || s>fromYMD(rangeEnd)) return false;
    return true;
  });
}

// ===== Exportações =====
function download(name,content,type="text/plain"){
  const blob=new Blob([content],{type});
  const a=document.createElement("a");
  a.href=URL.createObjectURL(blob);
  a.download=name;
  document.body.appendChild(a);
  a.click();
  setTimeout(()=>{URL.revokeObjectURL(a.href); a.remove();}, 0);
}

function toCSV(rows, headerOrder){
  const esc=v=>{
    if(v===null||v===undefined) return "";
    const s=String(v).replace(/"/g,'""');
    return /[",\n;]/.test(s)?`"${s}"`:s;
  };
  const header=headerOrder||Object.keys(rows[0]||{});
  const lines=[header.join(";")];
  rows.forEach(r=>{lines.push(header.map(h=>esc(r[h])).join(";"))});
  return lines.join("\n");
}

// Export CSV Recursos/Atividades (estado atual, incluindo IDs)
document.getElementById("btnExportCSV").onclick=()=>{
  const rec = resources.map(r=>({id:r.id,nome:r.nome,tipo:r.tipo,senioridade:r.senioridade,ativo:r.ativo,capacidade:r.capacidade||100,inicioAtivo:r.inicioAtivo||"",fimAtivo:r.fimAtivo||""}));
  const atv = activities.map(a=>({id:a.id,titulo:a.titulo,resourceId:a.resourceId,inicio:a.inicio,fim:a.fim,status:a.status,alocacao:a.alocacao||100}));
  download("recursos.csv", toCSV(rec, ["id","nome","tipo","senioridade","ativo","capacidade","inicioAtivo","fimAtivo"]), "text/csv;charset=utf-8");
  download("atividades.csv", toCSV(atv, ["id","titulo","resourceId","inicio","fim","status","alocacao"]), "text/csv;charset=utf-8");
  alert("Exportados: recursos.csv e atividades.csv");
};

// Export Excel (XLS compatível via HTML) — um arquivo para recursos e outro para atividades
function tableHTML(name, rows, cols){
  const header = cols.map(c=>`<th>${c}</th>`).join("");
  const body = rows.map(r=>`<tr>${cols.map(c=>`<td>${(r[c]??"")}</td>`).join("")}</tr>`).join("");
  return `
  <html xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:x="urn:schemas-microsoft-com:office:excel"
        xmlns="http://www.w3.org/TR/REC-html40">
  <head><meta charset="utf-8">
  <!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>
    <x:ExcelWorksheet><x:Name>${name}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet>
  </x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
  </head><body>
    <table border="1"><thead><tr>${header}</tr></thead><tbody>${body}</tbody></table>
  </body></html>`;
}
document.getElementById("btnExportXLS").onclick=()=>{
  const rec = resources.map(r=>({id:r.id,nome:r.nome,tipo:r.tipo,senioridade:r.senioridade,ativo:r.ativo,capacidade:r.capacidade||100,inicioAtivo:r.inicioAtivo||"",fimAtivo:r.fimAtivo||""}));
  const atv = activities.map(a=>({id:a.id,titulo:a.titulo,resourceId:a.resourceId,inicio:a.inicio,fim:a.fim,status:a.status,alocacao:a.alocacao||100}));
  download("recursos.xls", tableHTML("Recursos", rec, ["id","nome","tipo","senioridade","ativo","capacidade","inicioAtivo","fimAtivo"]), "application/vnd.ms-excel");
  download("atividades.xls", tableHTML("Atividades", atv, ["id","titulo","resourceId","inicio","fim","status","alocacao"]), "application/vnd.ms-excel");
  alert("Exportados: recursos.xls e atividades.xls");
};

// Export CSV diário (Power BI)
document.getElementById("btnExportPBI").onclick=()=>{
  const rows=[];
  const start=fromYMD(rangeStart), end=fromYMD(rangeEnd);
  const byId=Object.fromEntries(resources.map(r=>[r.id,r]));
  activities.forEach(a=>{
    const s=fromYMD(a.inicio), e=fromYMD(a.fim);
    const r=byId[a.resourceId];
    if(!r) return;
    for(let d=new Date(Math.max(s,start)); d<=Math.min(e,end); d=addDays(d,1)){
      rows.push({
        data: toYMD(d),
        atividadeId: a.id,
        atividadeTitulo: a.titulo,
        status: a.status,
        alocacao: a.alocacao||100,
        recursoId: r.id,
        recursoNome: r.nome,
        recursoTipo: r.tipo,
        recursoSenioridade: r.senioridade,
        recursoCapacidade: r.capacidade||100
      });
    }
  });
  download("powerbi_atividades_diarias.csv",
    toCSV(rows,["data","atividadeId","atividadeTitulo","status","alocacao","recursoId","recursoNome","recursoTipo","recursoSenioridade","recursoCapacidade"]),
    "text/csv;charset=utf-8");
  alert(`Exportado: powerbi_atividades_diarias.csv (${rows.length} linhas)`);
};

// Histórico consolidado (todas as atividades)
if(btnHistAll) btnHistAll.onclick=()=>{
  const rows=[];
  const byId=Object.fromEntries(resources.map(r=>[r.id,r]));
  Object.keys(trails).forEach(aid=>{
    const a=activities.find(x=>x.id===aid);
    (trails[aid]||[]).forEach(it=>{
      rows.push({
        atividadeId: aid,
        atividadeTitulo: a? a.titulo:"(excluída)",
        recursoId: a? a.resourceId:"",
        recursoNome: a && byId[a.resourceId]? byId[a.resourceId].nome:"",
        ts: it.ts,
        oldInicio: it.oldInicio, oldFim: it.oldFim,
        newInicio: it.newInicio, newFim: it.newFim,
        justificativa: it.justificativa||"",
        user: it.user||""
      });
    });
  });
  if(!rows.length){ alert("Sem registros de histórico."); return; }
  download("historico_consolidado.csv", toCSV(rows, ["atividadeId","atividadeTitulo","recursoId","recursoNome","ts","oldInicio","oldFim","newInicio","newFim","justificativa","user"]), "text/csv;charset=utf-8");
};

// Backup/Restore (JSON)
if(btnBackup) btnBackup.onclick=()=>{
  const dump={resources, activities, trails, meta:{version:"v2", exportedAt:new Date().toISOString()}};
  download("backup_planejador.json", JSON.stringify(dump,null,2), "application/json;charset=utf-8");
};
if(fileRestore) fileRestore.onchange=(ev)=>{
  const f=ev.target.files[0]; if(!f) return;
  const reader=new FileReader();
  reader.onload=()=>{
    try{
      const dump=JSON.parse(reader.result);
      if(!dump.resources || !dump.activities) throw new Error("Arquivo inválido.");
      resources=dump.resources; activities=dump.activities; trails=dump.trails||{};
      saveLS(LS.res,resources); saveLS(LS.act,activities); saveLS(LS.trail,trails||{});
      renderAll(); alert("Restauração concluída.");
    }catch(err){ alert("Falha ao restaurar: "+err.message); }
  };
  reader.readAsText(f,"utf-8");
};

// ===== Importar CSV (recursos.csv e/ou atividades.csv) =====
document.getElementById("fileImport").onchange=(ev)=>{
  const files=[...ev.target.files];
  if(!files.length) return;
  let pending=files.length;
  files.forEach(file=>{
    const reader=new FileReader();
    reader.onload=()=>{
      const text=reader.result;
      const lines=text.split(/\r?\n/).filter(l=>l.trim().length>0);
      if(!lines.length){ if(--pending===0){renderAll(); alert("Importação concluída.");} return; }
      const sep=lines[0].includes(";")?";":",";
      const headers=lines[0].split(sep).map(h=>h.trim());
      const idx=(name)=>headers.indexOf(name);
      if(headers.includes("nome") && headers.includes("capacidade")){
        // recursos.csv
        const arr=[];
        for(let i=1;i<lines.length;i++){
          const cols=lines[i].split(sep);
          if(cols.length!==headers.length) continue;
          const rec={
            id: cols[idx("id")]||uuid(),
            nome: cols[idx("nome")]||"",
            tipo: cols[idx("tipo")]||"Interno",
            senioridade: cols[idx("senioridade")]||"NA",
            ativo: (cols[idx("ativo")]||"true").toLowerCase()!=="false",
            capacidade: Number(cols[idx("capacidade")]||100),
            inicioAtivo: cols[idx("inicioAtivo")]||"",
            fimAtivo: cols[idx("fimAtivo")]||""
          };
          if(rec.nome) arr.push(rec);
        }
        const ids=new Set(resources.map(r=>r.id));
        resources=[...resources, ...arr.filter(r=>!ids.has(r.id))];
        saveLS(LS.res,resources);
      } else if(headers.includes("titulo") && headers.includes("resourceId")){
        // atividades.csv
        const arr=[];
        for(let i=1;i<lines.length;i++){
          const cols=lines[i].split(sep);
          if(cols.length!==headers.length) continue;
          const at={
            id: cols[idx("id")]||uuid(),
            titulo: cols[idx("titulo")]||"",
            resourceId: cols[idx("resourceId")]||"",
            inicio: cols[idx("inicio")]||"",
            fim: cols[idx("fim")]||"",
            status: cols[idx("status")]||"Planejada",
            alocacao: Number(cols[idx("alocacao")]||100)
          };
          if(at.titulo && at.resourceId && at.inicio && at.fim) arr.push(at);
        }
        const ids=new Set(activities.map(a=>a.id));
        activities=[...activities, ...arr.filter(a=>!ids.has(a.id))];
        saveLS(LS.act,activities);
      }
      if(--pending===0){ renderAll(); alert("Importação concluída."); }
    };
    reader.readAsText(file, "utf-8");
  });
};

// ===== Hist dialog export atual =====
btnHistExport.onclick=(e)=>{
  e.preventDefault();
  if(!histCurrentId){ alert("Abra o histórico de uma atividade."); return; }
  const list = trails[histCurrentId]||[];
  if(list.length===0){ alert("Sem registros para exportar."); return; }
  const s=document.getElementById('histStart').value;
  const e2=document.getElementById('histEnd').value;
  const rows = list.filter(it=>{
    const t=new Date(it.ts);
    return (!s || t>=fromYMD(s)) && (!e2 || t<=addDays(fromYMD(e2),0));
  }).map(it=>({ts:it.ts, oldInicio:it.oldInicio, oldFim:it.oldFim, newInicio:it.newInicio, newFim:it.newFim, justificativa:it.justificativa||"", user:it.user||""}));
  download(`historico_${histCurrentId}.csv`, toCSV(rows, ["ts","oldInicio","oldFim","newInicio","newFim","justificativa","user"]), "text/csv;charset=utf-8");
};

// ===== Agregados =====
aggGran.onchange=()=>renderAggregates();
function bucketKey(d, gran){
  if(gran==="weekly"){
    const date=new Date(d); const day=(date.getDay()+6)%7; 
    const monday=new Date(date); monday.setDate(date.getDate()-day);
    const y=monday.getFullYear(); const m=String(monday.getMonth()+1).padStart(2,"0"); const day2=String(monday.getDate()).padStart(2,"0");
    return `W ${y}-${m}-${day2}`;
  } else {
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
  }
}
function renderAggregates(){
  aggCharts.innerHTML="";
  const gran=aggGran.value;
  const days=buildDays();
  const byRes=Object.fromEntries(resources.filter(r=>r.ativo).map(r=>[r.id,{}]));
  activities.forEach(a=>{
    const r=resources.find(x=>x.id===a.resourceId && x.ativo);
    if(!r) return;
    const cap=r.capacidade||100;
    for(let d of days){
      if(fromYMD(a.inicio)<=d && d<=fromYMD(a.fim)){
        const key=bucketKey(d,gran);
        const map=byRes[r.id];
        if(!map[key]) map[key]={sum:0,capDays:0};
        map[key].sum += (a.alocacao||100);
        map[key].capDays += cap;
      }
    }
  });
  resources.filter(r=>r.ativo).forEach(r=>{
    const card=document.createElement("div");
    card.className="card";
    const h=document.createElement("h3");
    h.textContent=`${r.nome} — ${gran==="weekly"?"Semanal":"Mensal"}`;
    const canvas=document.createElement("canvas");
    canvas.width=600; canvas.height=140; canvas.className="chart";
    const ctx=canvas.getContext("2d");
    const entries=Object.entries(byRes[r.id]||{});
    entries.sort((a,b)=>a[0]>b[0]?1:-1);
    const margin=30, W=canvas.width, H=canvas.height;
    ctx.clearRect(0,0,W,H);
    ctx.beginPath(); ctx.moveTo(margin,10); ctx.lineTo(margin,H-margin+10); ctx.lineTo(W-10,H-margin+10); ctx.stroke();
    const barW=Math.max(8, (W - margin - 20) / Math.max(1, entries.length) - 6);
    entries.forEach((kv,idx)=>{
      const key=kv[0]; const v=kv[1];
      const perc = v.capDays? Math.min(100, (v.sum / v.capDays) * 100) : 0;
      const x = margin + 10 + idx*(barW+6);
      const y = (H - margin +10) - (perc/100)*(H - margin - 20);
      ctx.fillRect(x, y, barW, (H - margin +10) - y);
      ctx.save(); ctx.translate(x+barW/2, H - margin + 18); ctx.rotate(-Math.PI/4); ctx.textAlign="right"; ctx.font="10px sans-serif"; ctx.fillText(key, 0, 0); ctx.restore();
      ctx.font="10px sans-serif"; ctx.fillText(Math.round(perc)+"%", x, y-4);
    });
    card.appendChild(h); card.appendChild(canvas);
    aggCharts.appendChild(card);
  });
}

// ===== Render principal =====
function renderAll(){
  const filtered=getFilteredActivities();
  renderTables(filtered);
  renderGantt(filtered);
  renderAggregates();
}

renderStatusChips();

// ===== Disponibilidade (% de capacidade livre) =====
(function(){
  const avBtn = document.getElementById('avBtn');
  const avRes = document.getElementById('avResultado');
  if(!avBtn || !avRes) return;

  function isBusinessDay(d){
    const wd = d.getDay(); // 0=dom..6=sab
    return wd>=1 && wd<=5;
  }

  function sumAllocationOn(resourceId, date) {
    // ignore Concluída/Cancelada para disponibilidade
    const acts = activities.filter(a=>a.resourceId===resourceId && a.status!=='Concluída' && a.status!=='Cancelada' && fromYMD(a.inicio)<=date && date<=fromYMD(a.fim));
    return acts.reduce((acc,a)=>acc+(a.alocacao||100),0);
  }

  function hasWindow(resource, startDate, daysNeeded, businessOnly, requiredPerc){
    const cap = resource.capacidade||100;
    const limit = fromYMD(rangeEnd);
    let d = new Date(startDate);
    const maxSearch = new Date(startDate.getFullYear()+1, startDate.getMonth(), startDate.getDate());
    const hardLimit = limit && limit>startDate ? limit : maxSearch;
    function recIsActiveOn(dt){
      if(resource.inicioAtivo && dt < fromYMD(resource.inicioAtivo)) return false;
      if(resource.fimAtivo && dt > fromYMD(resource.fimAtivo)) return false;
      return true;
    }
    while(d <= hardLimit){
      let cnt = 0;
      let step = new Date(d);
      let actualStart = null;
      let ok = true;
      let guard = 0;
      while(cnt < daysNeeded && guard < 4000){
        guard++;
        if(businessOnly && !isBusinessDay(step)){
          step = new Date(step.getFullYear(), step.getMonth(), step.getDate()+1);
          continue;
        }
        if(!recIsActiveOn(step)) { ok=false; break; }
        const used = sumAllocationOn(resource.id, step);
        const free = (cap - used);
        if(free < requiredPerc){ ok=false; break; }
        if(actualStart===null) actualStart = new Date(step);
        cnt++;
        step = new Date(step.getFullYear(), step.getMonth(), step.getDate()+1);
      }
      if(ok && cnt>=daysNeeded && actualStart){
        return toYMD(actualStart);
      }
      d = new Date(d.getFullYear(), d.getMonth(), d.getDate()+1);
    }
    return null;
  }

  function runAvailability(){
    const dias = Math.max(1, Number(document.getElementById('avDias').value||1));
    const uteis = (document.getElementById('avUteis').value||'1')==='1';
    const percReq = Math.max(1, Number(document.getElementById('avPerc').value||25));
    const tipo = document.getElementById('avTipo').value||'';
    const sen = document.getElementById('avSenioridade').value||'';
    const inicioStr = document.getElementById('avInicio').value || toYMD(today);
    const inicio = fromYMD(inicioStr);

    const out = [];
    resources.forEach(r=>{
      if(!r.ativo) return;
      if(tipo && r.tipo!==tipo) return;
      if(sen && r.senioridade!==sen) return;
      const next = hasWindow(r, inicio, dias, uteis, percReq);
      if(next) out.push({recurso:r, inicio:next});
    });
    out.sort((a,b)=> fromYMD(a.inicio) - fromYMD(b.inicio) || a.recurso.nome.localeCompare(b.recurso.nome));
    if(!out.length){
      avRes.innerHTML = '<div class="muted">Nenhum recurso atende aos critérios dentro do horizonte de busca.</div>';
      return;
    }
    const rows = out.map(it=>`<tr><td>${it.recurso.nome}</td><td>${it.recurso.tipo}</td><td>${it.recurso.senioridade}</td><td>${it.inicio}</td></tr>`).join('');
    avRes.innerHTML = `<table class="tbl"><thead><tr><th>Recurso</th><th>Tipo</th><th>Senioridade</th><th>Data mais próxima</th></tr></thead><tbody>${rows}</tbody></table>`;
  }

  avBtn.addEventListener('click', runAvailability);
  const avInicio = document.getElementById('avInicio');
  if(avInicio && !avInicio.value){ avInicio.value = toYMD(today); }
})();
renderAll();


// ===== KPIs (Visão Executiva) =====
function renderKPIs(){
  try{
    const total=activities.length;
    const concluidas=activities.filter(a=>a.status==="Concluída").length;
    const perc=total? Math.round((concluidas/total)*100):0;
    const el1=document.getElementById("kpiExecucao"); if(el1) el1.textContent=perc+"%";
    const el2=document.getElementById("kpiRecursos"); if(el2) el2.textContent=resources.filter(r=>r.ativo).length;
    let sobre=0;
    resources.filter(r=>r.ativo).forEach(r=>{
      const cap=r.capacidade||100;
      const days=buildDays();
      for(const d of days){
        // Ignore activities that are concluded or cancelled when calculating overload
        const acts=activities.filter(a=>a.resourceId===r.id && a.status !== 'Concluída' && a.status !== 'Cancelada' && fromYMD(a.inicio)<=d && d<=fromYMD(a.fim));
        const sum=acts.reduce((acc,a)=>acc+(a.alocacao||100),0);
        if(sum>cap){sobre++; break;}
      }
    });
    const el3=document.getElementById("kpiSobrecarga"); if(el3) el3.textContent=sobre;
    // Renderizar detalhes de sobrecarga (exec)
    try{
      renderOverloadDetails();
    }catch(err){ console.error(err); }
  }catch(e){ /* ignora falhas de KPI fora da aba */ }
}

// ===== Exportar PDF (fallback para janela de impressão se jsPDF/html2canvas não estiverem disponíveis) =====
(function(){
  const btn=document.getElementById("btnExportPDF");
  if(!btn) return;
  btn.onclick = async () => {
    const hasJsPDF = !!(window.jspdf && window.jspdf.jsPDF);
    const hasHtml2Canvas = !!window.html2canvas;
    if(hasJsPDF){
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF();
      doc.setFontSize(14);
      doc.text("Relatório Planejador de Recursos", 14, 20);
      doc.setFontSize(10);
      doc.text("Gerado em: "+new Date().toLocaleString(), 14, 28);
      // KPIs
      renderKPIs();
      doc.text("KPIs:", 14, 40);
      const k1=document.getElementById("kpiExecucao")?.textContent||"";
      const k2=document.getElementById("kpiRecursos")?.textContent||"";
      const k3=document.getElementById("kpiSobrecarga")?.textContent||"";
      doc.text("% Execução: "+k1, 20, 48);
      doc.text("Recursos Ativos: "+k2, 20, 56);
      doc.text("Recursos Sobrecarregados: "+k3, 20, 64);
      // Listas (resumo textual)
      doc.text("Recursos:", 14, 78);
      let y=86;
      resources.forEach(r=>{ doc.text("- "+r.nome+" ("+r.tipo+", "+r.senioridade+", Cap "+(r.capacidade||100)+"%)", 20, y); y+=6; if(y>270){doc.addPage(); y=20;} });
      y+=6; doc.text("Atividades:", 14, y); y+=8;
      activities.forEach(a=>{
        const rec=resources.find(r=>r.id===a.resourceId);
        doc.text("- "+a.titulo+" ("+(rec?rec.nome:"—")+") ["+a.status+"] "+a.inicio+" → "+a.fim, 20, y);
        y+=6; if(y>270){doc.addPage(); y=20;}
      });
      // Snapshot Gantt (opcional)
      if(hasHtml2Canvas){
        try{
          const canvas = await html2canvas(document.getElementById("gantt"));
          const img = canvas.toDataURL("image/png");
          doc.addPage(); doc.text("Visão Gantt",14,20); doc.addImage(img,"PNG",14,30,180,100);
        }catch(e){ /* ignora */ }
      }
      doc.save("planejador_relatorio.pdf");
    } else {
      // Fallback: abre janela de impressão (usuario pode salvar como PDF)
      const w = window.open("", "_blank");
      const cssCompact = `body{font-family:Arial,sans-serif;padding:16px} h2{margin:16px 0 8px} table{width:100%;border-collapse:collapse;font-size:12px} th,td{border:1px solid #ddd;padding:6px}`;
      w.document.write("<html><head><title>Relatório Planejador</title><style>"+cssCompact+"</style></head><body>");
      w.document.write("<h1>Relatório Planejador de Recursos</h1>");
      w.document.write("<div>Gerado em: "+new Date().toLocaleString()+"</div>");
      renderKPIs();
      w.document.write("<h2>KPIs</h2>");
      w.document.write("<div>% Execução: "+(document.getElementById("kpiExecucao")?.textContent||"0%")+"</div>");
      w.document.write("<div>Recursos Ativos: "+(document.getElementById("kpiRecursos")?.textContent||"0")+"</div>");
      w.document.write("<div>Recursos Sobrecarregados: "+(document.getElementById("kpiSobrecarga")?.textContent||"0")+"</div>");
      // Tabelas simples
      w.document.write("<h2>Recursos</h2><table><tr><th>Nome</th><th>Tipo</th><th>Senioridade</th><th>Capacidade%</th></tr>");
      resources.forEach(r=>{ w.document.write("<tr><td>"+r.nome+"</td><td>"+r.tipo+"</td><td>"+r.senioridade+"</td><td>"+(r.capacidade||100)+"</td></tr>"); });
      w.document.write("</table>");
      w.document.write("<h2>Atividades</h2><table><tr><th>Título</th><th>Recurso</th><th>Status</th><th>Início</th><th>Fim</th><th>Alocação%</th></tr>");
      activities.forEach(a=>{
        const rec=resources.find(r=>r.id===a.resourceId);
        w.document.write("<tr><td>"+a.titulo+"</td><td>"+(rec?rec.nome:"—")+"</td><td>"+a.status+"</td><td>"+a.inicio+"</td><td>"+a.fim+"</td><td>"+(a.alocacao||100)+"</td></tr>");
      });
      w.document.write("</table>");
      try{ w.document.close(); w.focus(); w.print(); }catch(e){}
    }
  };
})();

// ===== renderTables (sobrescrito para incluir 'Duplicar') =====
function renderTables(filteredActs){
  // Recursos
  tblRecursos.innerHTML="";
  resources.forEach(r=>{
    const tr=document.createElement("tr");
    tr.innerHTML=`
      <td>${r.nome}</td>
      <td>${r.tipo}</td>
      <td>${r.senioridade}</td>
      <td>${r.ativo?"Sim":"Não"}</td>
      <td>${r.capacidade}</td>
      <td>${r.inicioAtivo||""}</td>
      <td>${r.fimAtivo||""}</td>
      <td class="actions">
        <button class="btn">Editar</button>
        <button class="btn dup">Duplicar</button>
        <button class="btn danger">Excluir</button>
      </td>`;
    const [btnEdit, btnDup, btnDel] = tr.querySelectorAll("button");
    btnEdit.onclick=()=>{
      dlgRecursoTitulo.textContent="Editar Recurso";
      formRecurso.elements["id"].value=r.id;
      formRecurso.elements["nome"].value=r.nome;
      formRecurso.elements["tipo"].value=r.tipo;
      formRecurso.elements["senioridade"].value=r.senioridade;
      formRecurso.elements["ativo"].checked=!!r.ativo;
      formRecurso.elements["capacidade"].value=r.capacidade||100;
      formRecurso.elements["inicioAtivo"].value=r.inicioAtivo||"";
      formRecurso.elements["fimAtivo"].value=r.fimAtivo||"";
      dlgRecurso.showModal();
    };
    btnDup.onclick=()=>{
      const copy={...r,id:uuid(),nome:"Cópia de "+r.nome};
      resources.push(copy); saveLS(LS.res,resources); renderAll();
    };
    btnDel.onclick=()=>{
      if(!confirm("Remover recurso e suas alocações?")) return;
      resources=resources.filter(x=>x.id!==r.id);
      activities=activities.filter(a=>a.resourceId!==r.id);
      saveLS(LS.res,resources); saveLS(LS.act,activities);
      renderAll();
    };
    tblRecursos.appendChild(tr);
  });

  // Atividades
  tblAtividades.innerHTML="";
  filteredActs.forEach(a=>{
    const r=resources.find(x=>x.id===a.resourceId);
    const tr=document.createElement("tr");
    tr.innerHTML=`
      <td>${a.titulo}</td>
      <td>${r? r.nome:"—"}</td>
      <td>${a.inicio}</td>
      <td>${a.fim}</td>
      <td>${a.status}</td>
      <td>${a.alocacao||100}</td>
      <td class="actions">
        <button class="btn">Editar</button>
        <button class="btn">Histórico</button>
        <button class="btn dup">Duplicar</button>
        <button class="btn danger">Excluir</button>
      </td>`;
    const [btnEdit, btnHist, btnDup, btnDel] = tr.querySelectorAll("button");
    btnEdit.onclick=()=>{
      dlgAtividadeTitulo.textContent="Editar Atividade";
      fillRecursoOptions();
      formAtividade.elements["id"].value=a.id;
      formAtividade.elements["titulo"].value=a.titulo;
      formAtividade.elements["resourceId"].value=a.resourceId;
      formAtividade.elements["inicio"].value=a.inicio;
      formAtividade.elements["fim"].value=a.fim;
      formAtividade.elements["status"].value=a.status;
      formAtividade.elements["alocacao"].value=a.alocacao||100;
      dlgAtividade.showModal();
    };
    btnHist.onclick=()=>{
      histCurrentId=a.id;
      const list = trails[a.id]||[];
      if(list.length===0){
        histList.innerHTML='<div class="muted">Sem alterações de datas registradas para esta atividade.</div>';
      }else{
        const s=document.getElementById('histStart').value;
        const e=document.getElementById('histEnd').value;
        const rows=list.slice().reverse().filter(it=>{
          const t=new Date(it.ts);
          return (!s || t>=fromYMD(s)) && (!e || t<=addDays(fromYMD(e),0));
        }).map(it=>{
          return `<div style="padding:6px 8px; background:#fff; border:1px solid #e2e8f0; border-radius:8px; margin:6px 0">
              <div style="font-size:12px;color:#475569">${new Date(it.ts).toLocaleString()}${it.user? ' • ' + it.user : ''}</div>
              <div style="font-weight:600">Início: ${it.oldInicio} → ${it.newInicio} | Fim: ${it.oldFim} → ${it.newFim}</div>
              <div style="margin-top:4px">${it.justificativa? it.justificativa.replace(/</g,'&lt;') : ''}</div>
            </div>`;
        }).join("");
        histList.innerHTML=rows || '<div class="muted">Sem registros no período.</div>';
      }
      dlgHist.showModal();
      const btn=document.getElementById('histApply'); if(btn){ btn.onclick=()=>{ btnHist.onclick(); }; }
    };
    btnDup.onclick=()=>{
      const copy={...a,id:uuid(),titulo:"Cópia de "+a.titulo};
      activities.push(copy); saveLS(LS.act,activities); renderAll();
    };
    btnDel.onclick=()=>{
      if(!confirm("Remover atividade?")) return;
      activities=activities.filter(x=>x.id!==a.id);
      saveLS(LS.act,activities);
      renderAll();
    };
    tblAtividades.appendChild(tr);
  });
}

// ===== Hook no renderAll para atualizar KPIs sem quebrar fluxo =====
(function(){
  const _renderAll = renderAll;
  renderAll = function(){
    _renderAll();
    renderKPIs();
  };
})();

// === Helpers para detalhamento de sobrecarga na Visão Executiva ===
function computeOverloads(){
  const result=[];
  resources.forEach(r=>{
    const cap=r.capacidade||100;
    const tasks=activities.filter(a=>a.resourceId===r.id && a.status!=='Concluída' && a.status!=='Cancelada');
    if(!tasks.length) return;
    const events=[];
    tasks.forEach(a=>{
      const start=fromYMD(a.inicio);
      const end=fromYMD(a.fim);
      const alloc=a.alocacao||100;
      events.push({date:new Date(start),delta:alloc});
      const after=new Date(end);
      after.setDate(after.getDate()+1);
      events.push({date:after,delta:-alloc});
    });
    events.sort((a,b)=>a.date-b.date);
    let sum=0;
    let openStart=null;
    for(let i=0;i<events.length;i++){
      sum += events[i].delta;
      if(sum>cap && openStart===null){
        openStart=new Date(events[i].date);
      }
      if(sum<=cap && openStart!==null){
        const endDate=new Date(events[i].date);
        endDate.setDate(endDate.getDate()-1);
        const conc=tasks.filter(a=>{
          const s=fromYMD(a.inicio);
          const e=fromYMD(a.fim);
          return s<=endDate && e>=openStart;
        }).map(a=> a.titulo || ("Atividade "+a.id));
        result.push({nome:r.nome, periodo: toYMD(openStart)+" → "+toYMD(endDate), atividades: Array.from(new Set(conc)).join(', ')});
        openStart=null;
      }
    }
    if(openStart!==null){
      let lastEnd=new Date(0);
      tasks.forEach(a=>{
        const e=fromYMD(a.fim);
        if(e>lastEnd) lastEnd=new Date(e);
      });
      const conc=tasks.filter(a=>{
        const s=fromYMD(a.inicio);
        const e=fromYMD(a.fim);
        return s<=lastEnd && e>=openStart;
      }).map(a=> a.titulo || ("Atividade "+a.id));
      result.push({nome:r.nome, periodo: toYMD(openStart)+" → "+toYMD(lastEnd), atividades: Array.from(new Set(conc)).join(', ')});
    }
  });
  return result;
}

function renderOverloadDetails(){
  const tbody=document.querySelector('#overloadDetails tbody');
  const wrap=document.getElementById('overloadDetailsWrap');
  const empty=document.getElementById('overloadEmptyMsg');
  if(!tbody||!wrap||!empty) return;
  const rows=computeOverloads();
  if(!rows.length){
    wrap.style.display='none';
    empty.style.display='';
    return;
  }
  empty.style.display='none';
  wrap.style.display='';
  tbody.innerHTML = rows.map(r=>`<tr><td>${r.nome}</td><td>${r.periodo}</td><td>${r.atividades}</td></tr>`).join('');
}
