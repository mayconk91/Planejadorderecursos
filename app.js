// ===== Utilidades de data =====
function toYMD(d){const y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,'0'),day=String(d.getDate()).padStart(2,'0');return `${y}-${m}-${day}`}
function fromYMD(s){return new Date(`${s}T00:00:00`)}
function addDays(d,n){const nd=new Date(d);nd.setDate(nd.getDate()+n);return nd}
function diffDays(a,b){const A=new Date(a.getFullYear(),a.getMonth(),a.getDate());const B=new Date(b.getFullYear(),b.getMonth(),b.getDate());return Math.round((A-B)/(1000*60*60*24))}
function clampDate(d){return new Date(d.getFullYear(),d.getMonth(),d.getDate())}
function uuid(){if (crypto && crypto.randomUUID) return crypto.randomUUID(); const s=()=>Math.floor((1+Math.random())*0x10000).toString(16).substring(1); return `${s()}${s()}-${s()}-${s()}-${s()}-${s()}${s()}${s()}`}

// ===== Importação em massa =====
// Normaliza nomes para comparação (trim e lower case). Utilizado no merge de recursos e atividades importados.
/**
 * Normaliza um nome para comparação, removendo acentuação, espaços nas extremidades
 * e convertendo para minúsculas. Isso permite que nomes com
 * diacríticos sejam comparados de forma insensível a acentos (ex.:
 * "João" e "Joao" serão tratados como iguais). Se o valor for
 * nulo/undefined, retorna uma string vazia.
 *
 * @param {string} str Texto a normalizar
 * @returns {string} Nome normalizado
 */
function normalizeName(str) {
  if (!str) return '';
  // Converte para string, remove espaços extras, normaliza unicode
  const s = String(str).trim().toLowerCase().normalize('NFD');
  // Substitui alguns caracteres especiais antes de remover diacríticos.
  // Por exemplo: 'ç' -> 'c', 'Ç' -> 'c', 'ñ' -> 'n'.
  let normalized = s
    .replace(/ç/g, 'c')
    .replace(/Ç/g, 'c')
    .replace(/ñ/g, 'n')
    .replace(/Ñ/g, 'n');
  // Remove diacríticos (acentos)
  normalized = normalized.replace(/[\u0300-\u036f]/g, '');
  return normalized;
}

// ===== Utilidade de Tags =====
function normalizeTag(tag) {
  if (!tag) return '';
  const trimmed = tag.trim();
  // Primeira letra maiúscula, resto minúsculo
  return trimmed.charAt(0).toUpperCase() + trimmed.slice(1).toLowerCase();
}

// ===== Domínio =====
const TIPOS=["Interno","Externo"];
const SENIORIDADES=["Jr","Pl","Sr","NA"];
const STATUS=["Planejada","Em Execução","Bloqueada","Concluída","Cancelada"];

// ===== Persistência =====
// Principais chaves para arrays de domínio
const LS={res:"rp_resources_v2",act:"rp_activities_v2",trail:"rp_trail_v1",user:"rp_user_v1"};

// Chaves adicionais para log de eventos e snapshots.  O log de eventos registra cada
// alteração de recurso/atividade (criação, atualização, exclusão) de forma
// append-only.  Snapshots são capturas periódicas do estado completo e
// permitem reconstruir rapidamente o estado sem precisar reprocessar todos
// os eventos.
const LS_LOG="rp_event_log_v1";
const LS_SNAP="rp_snapshot_v1";
function loadLS(k,fallback){try{const raw=localStorage.getItem(k);return raw?JSON.parse(raw):fallback}catch{return fallback}}
function saveLS(k,v){localStorage.setItem(k,JSON.stringify(v))}

// ===== Event Log e Snapshot =====
// Carrega log de eventos e snapshot, se existirem.  O log é uma lista de
// objetos {ts, type, action, id, data}.  O snapshot é um objeto
// {ts, resources, activities, trails}.  Essas estruturas são usadas para
// reconstruir o estado no início da aplicação e reduzir conflitos de
// concorrência.
let eventLog = loadLS(LS_LOG, []);
let snapshot = loadLS(LS_SNAP, null);

/**
 * Reinicia o log de eventos e o snapshot. Deve ser chamado quando o usuário
 * aponta um novo arquivo de Banco de Dados como origem, para evitar que
 * eventos antigos (antes de selecionar o BD) sejam reaplicados e causem
 * "dados fantasmas" ao reconstruir o estado. Zera o eventLog e o
 * snapshot persistidos em localStorage.
 */
function resetEventLogAndSnapshot(){
  try {
    eventLog = [];
    saveLS(LS_LOG, eventLog);
  } catch(e){ console.warn('Erro ao zerar o log de eventos', e); }
  try {
    snapshot = null;
    saveLS(LS_SNAP, snapshot);
  } catch(e){ console.warn('Erro ao zerar snapshot', e); }
}

/**
 * Aplica eventos da lista eventLog sobre os estados base (resources/activities)
 * e substitui os arrays globais.  Caso exista um snapshot válido, ele é
 * utilizado como base; caso contrário, os arrays atuais de resources e
 * activities são usados como base.  Ao final, os estados globais e o
 * localStorage são atualizados para refletir o resultado.
 */
function loadSnapshotAndEvents() {
  try {
    let baseRes, baseActs, baseTrails;
    if (snapshot && Array.isArray(snapshot.resources) && Array.isArray(snapshot.activities)) {
      baseRes = snapshot.resources.map(r => ({ ...r }));
      baseActs = snapshot.activities.map(a => ({ ...a }));
      baseTrails = { ...(snapshot.trails || {}) };
    } else {
      // Caso não haja snapshot, utiliza os arrays atuais como base
      baseRes = (resources || []).map(r => ({ ...r }));
      baseActs = (activities || []).map(a => ({ ...a }));
      baseTrails = { ...(trails || {}) };
    }
    // Mapas para acesso rápido por id
    const resMap = {};
    baseRes.forEach(r => { resMap[r.id] = { ...r }; });
    const actMap = {};
    baseActs.forEach(a => { actMap[a.id] = { ...a }; });

    // Ordena o log por timestamp crescente para garantir reprodutibilidade
    const sortedEvents = (eventLog || []).slice().sort((a, b) => (a.ts || 0) - (b.ts || 0));
    for (const ev of sortedEvents) {
      if (!ev || !ev.type || !ev.action) continue;
      if (ev.type === 'resource') {
        if (ev.action === 'create' || ev.action === 'update') {
          resMap[ev.id] = { ...ev.data };
        } else if (ev.action === 'delete') {
          if (resMap[ev.id]) {
            resMap[ev.id] = { ...resMap[ev.id], deletedAt: ev.data.deletedAt, updatedAt: ev.data.deletedAt, version: (resMap[ev.id].version || 0) + 1 };
          }
          // marcas todas as atividades relacionadas como deletadas
          for (const aid in actMap) {
            if (actMap[aid] && actMap[aid].resourceId === ev.id) {
              actMap[aid] = { ...actMap[aid], deletedAt: ev.data.deletedAt, updatedAt: ev.data.deletedAt, version: (actMap[aid].version || 0) + 1 };
            }
          }
        }
      } else if (ev.type === 'activity') {
        if (ev.action === 'create' || ev.action === 'update') {
          actMap[ev.id] = { ...ev.data };
        } else if (ev.action === 'delete') {
          if (actMap[ev.id]) {
            actMap[ev.id] = { ...actMap[ev.id], deletedAt: ev.data.deletedAt, updatedAt: ev.data.deletedAt, version: (actMap[ev.id].version || 0) + 1 };
          }
        }
      }
    }
    resources = Object.values(resMap);
    activities = Object.values(actMap);
    trails = baseTrails;
    // Persistir o estado reconstruído
    saveLS(LS.res, resources);
    saveLS(LS.act, activities);
    saveLS(LS.trail, trails);
  } catch (e) {
    console.error('Erro ao aplicar snapshot/eventos', e);
  }
}

/**
 * Registra um evento no log.  O evento inclui: timestamp atual,
 * tipo (resource|activity), ação (create|update|delete), id do item
 * e dados (objeto com os campos relevantes ou no caso de delete
 * apenas deletedAt).  Após registrar, salva o log no localStorage e
 * invoca maybeSaveSnapshot() para decidir se um snapshot deve ser gerado.
 * @param {string} type Tipo da entidade ('resource' ou 'activity')
 * @param {string} action Ação realizada ('create','update','delete')
 * @param {string} id ID do item
 * @param {object} data Dados atuais do item (ou {deletedAt:ts} para delete)
 */
function recordEvent(type, action, id, data) {
  try {
    const ev = { ts: Date.now(), type, action, id, data };
    eventLog.push(ev);
    saveLS(LS_LOG, eventLog);
    maybeSaveSnapshot();
  } catch (e) {
    console.error('Erro ao registrar evento', e);
  }
}

/**
 * Gera um snapshot do estado atual se o log atingir um limiar definido.
 * O snapshot captura resources, activities e trails e armazena em LS_SNAP.
 * Após salvar o snapshot, o log é esvaziado para evitar crescimento
 * descontrolado.  Ajuste o threshold conforme necessário.
 */
function maybeSaveSnapshot() {
  const threshold = 30;
  if ((eventLog || []).length >= threshold) {
    try {
      snapshot = {
        ts: Date.now(),
        resources: (resources || []).map(r => ({ ...r })),
        activities: (activities || []).map(a => ({ ...a })),
        trails: { ...(trails || {}) }
      };
      saveLS(LS_SNAP, snapshot);
      // Limpa o log
      eventLog = [];
      saveLS(LS_LOG, eventLog);
    } catch (e) {
      console.error('Erro ao salvar snapshot', e);
    }
  }
}

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

// === Observador (watcher) para sincronização multiusuário ===
let fsaWatchTimer = null;
let fsaLastSeen = {};

function startFsaWatcher() {
  if (fsaWatchTimer) clearInterval(fsaWatchTimer);
  fsaLastSeen = {};
  if (!dirHandle) return;
  fsaWatchTimer = setInterval(async () => {
    try {
      for (const key of [LS.res, LS.act, LS.trail]) {
        const fileName = DATAFILES[key];
        let fh;
        try {
          fh = await dirHandle.getFileHandle(fileName, { create: false });
        } catch (e) {
          continue; 
        }
        const file = await fh.getFile();
        const lm = file.lastModified;
        if (!fsaLastSeen[fileName] || lm > fsaLastSeen[fileName]) {
          fsaLastSeen[fileName] = lm;
          const text = await file.text();
          let changed = false;
          try {
            let data;
            if (key === LS.trail) {
              data = JSON.parse(text || '{}');
            } else {
              data = JSON.parse(text || '[]');
            }
            if (key === LS.res) {
              if (JSON.stringify(resources) !== JSON.stringify(data)) {
                resources = data;
                saveLS(LS.res, resources);
                changed = true;
              }
            } else if (key === LS.act) {
              if (JSON.stringify(activities) !== JSON.stringify(data)) {
                activities = data;
                saveLS(LS.act, activities);
                changed = true;
              }
            } else if (key === LS.trail) {
              if (JSON.stringify(trails) !== JSON.stringify(data)) {
                trails = data;
                saveLS(LS.trail, trails);
                changed = true;
              }
            }
          } catch (e) {
            // JSON inválido: ignorar alteração
          }
          if (changed) {
            // Ao detectar mudanças em arquivos da pasta, recarregue snapshot e eventos antes de renderizar.
            try {
              // Reconstrói o estado a partir do snapshot mais recente e aplica eventos pendentes.
              loadSnapshotAndEvents();
            } catch (e) {
              console.error('Erro ao aplicar snapshot e eventos após mudança no FSA:', e);
            }
            renderAll();
            updateFolderStatus('Atualizado por outra sessão às ' + new Date(lm).toLocaleTimeString());
            try { saveBDDebounced(); } catch (e) {}
          }
        }
      }
    } catch (e) {
      // falhas silenciosas
    }
  }, 3000);
}

// === Banco de Dados (BD) ===
let bdHandle = null;
let bdFileExt = '';
let bdFileName = '';
let _saveBDTimer = null;

// === Rastreio simples da última gravação do BD ===
// Este valor é atualizado sempre que um salvamento ocorre ou quando o watcher detecta
// uma mudança externa. É utilizado para detectar se o arquivo foi modificado por
// outra sessão para evitar sobrescrever alterações recentes.
if (typeof window !== 'undefined') {
  window.__bdLastWrite = 0;
}

// === Observador para o arquivo de Banco de Dados (Excel/CSV) ===
let bdWatchTimer = null;
let bdLastSeen = null;

function startBDWatcher() {
  if (bdWatchTimer) clearInterval(bdWatchTimer);
  bdLastSeen = null;
  if (!bdHandle) return;
  bdWatchTimer = setInterval(async () => {
    try {
      const file = await bdHandle.getFile();
      const lm = file.lastModified;
      if (!bdLastSeen || lm > bdLastSeen) {
        bdLastSeen = lm;
        let parsed;
        if (bdFileExt === 'csv') {
          const txt = await file.text();
          parsed = parseCSVBDUnico(txt);
        } else {
          const txt = await file.text();
          parsed = parseHTMLBDTables(txt);
        }
        const newResources = (parsed.recursos || []).map(coerceResource);
        const newActivities = (parsed.atividades || []).map(coerceActivity);
        const newHoras = parsed.horas || [];
        const newCfg = parsed.cfg || [];
        const newFeriados = parsed.feriados || [];
        const newTrails = {};
        (parsed.historico || []).forEach(h => {
          const id = h.activityId;
          if (!id) return;
          if (!newTrails[id]) newTrails[id] = [];
          newTrails[id].push({
            ts: h.timestamp,
            oldInicio: h.oldInicio,
            oldFim: h.oldFim,
            newInicio: h.newInicio,
            newFim: h.newFim,
            justificativa: h.justificativa,
            user: h.user
          });
        });

        let changed = false;
        if (JSON.stringify(resources) !== JSON.stringify(newResources)) {
          resources = newResources;
          saveLS(LS.res, resources);
          changed = true;
        }
        if (JSON.stringify(activities) !== JSON.stringify(newActivities)) {
          activities = newActivities;
          saveLS(LS.act, activities);
          changed = true;
        }
        if (JSON.stringify(trails) !== JSON.stringify(newTrails)) {
          trails = newTrails;
          saveLS(LS.trail, trails);
          changed = true;
        }

        let horasChanged = false;
        try {
          if (typeof window.getHorasExternosData === 'function' && typeof window.setHorasExternosData === 'function') {
            const curHoras = window.getHorasExternosData() || [];
            if (JSON.stringify(curHoras) !== JSON.stringify(newHoras)) {
              window.setHorasExternosData(newHoras);
              horasChanged = true;
            }
          }
        } catch(e) {}
        try {
          if (typeof window.getHorasExternosConfig === 'function' && typeof window.setHorasExternosConfig === 'function') {
            const curCfg = window.getHorasExternosConfig() || [];
            if (JSON.stringify(curCfg) !== JSON.stringify(newCfg)) {
              window.setHorasExternosConfig(newCfg);
              horasChanged = true;
            }
          }
        } catch(e) {}
        try {
            if (typeof window.getFeriados === 'function' && typeof window.setFeriados === 'function') {
                const curFeriados = window.getFeriados() || [];
                if (JSON.stringify(curFeriados) !== JSON.stringify(newFeriados)) {
                    window.setFeriados(newFeriados);
                    horasChanged = true;
                }
            }
        } catch(e) {}
        if (changed || horasChanged) {
          // Reaplica snapshot e log antes de renderizar, garantindo que deleções e updates do log sejam reconstituídos.
          try {
            loadSnapshotAndEvents();
          } catch (e) {
            console.error('Erro ao aplicar snapshot e eventos após mudança no BD:', e);
          }
          renderAll();
          updateBDStatus('Atualizado por outra sessão às ' + new Date(lm).toLocaleTimeString());
          // Atualiza o rastreio de última modificação quando o arquivo é alterado por outra sessão
          if (typeof window !== 'undefined') {
            window.__bdLastWrite = lm;
          }
        }
      }
    } catch (e) {
      // Erros silenciosos
    }
  }, 3000);
}

// ===== Overlay de espera para gravação concorrente =====
let _bdWaitInterval = null;
function showBDWaitOverlay(seconds) {
  const overlay = document.getElementById('bdWaitOverlay');
  if (!overlay) return;
  // Exibe overlay
  overlay.style.display = 'flex';
  const countdownEl = document.getElementById('bdWaitCountdown');
  let remaining = seconds;
  if (countdownEl) countdownEl.textContent = remaining;
  // Limpa intervalos anteriores
  if (_bdWaitInterval) clearInterval(_bdWaitInterval);
  _bdWaitInterval = setInterval(() => {
    remaining -= 1;
    if (countdownEl) countdownEl.textContent = remaining >= 0 ? remaining : 0;
    if (remaining <= 0) {
      clearInterval(_bdWaitInterval);
      _bdWaitInterval = null;
    }
  }, 1000);
}

function hideBDWaitOverlay() {
  const overlay = document.getElementById('bdWaitOverlay');
  if (!overlay) return;
  overlay.style.display = 'none';
  if (_bdWaitInterval) {
    clearInterval(_bdWaitInterval);
    _bdWaitInterval = null;
  }
}

/**
 * Aguarda se outra sessão estiver salvando recentemente o BD.
 * Se o arquivo foi modificado por outra sessão nos últimos milissegundos definidos,
 * exibe um overlay com contagem regressiva e aguarda até que seja seguro salvar.
 */
async function acquireBDLockIfBusy() {
  // se não houver BD selecionado, nada a fazer
  if (!bdHandle) return true;
  const WAIT_WINDOW_MS = 5000; // Janela de espera de 5s para considerar salvamento recente
  while (true) {
    try {
      const file = await bdHandle.getFile();
      const lm = file.lastModified;
      const now = Date.now();
      // se o arquivo foi modificado por outra sessão (lm > __bdLastWrite) e ainda está dentro da janela, aguarda
      if (typeof window !== 'undefined' && lm > (window.__bdLastWrite || 0) && (now - lm) < WAIT_WINDOW_MS) {
        const remainingMs = WAIT_WINDOW_MS - (now - lm);
        const secondsLeft = Math.ceil(remainingMs / 1000);
        showBDWaitOverlay(secondsLeft);
        // aguarda a duração restante
        await new Promise(resolve => setTimeout(resolve, remainingMs));
        hideBDWaitOverlay();
        // após aguardar, continua o loop para revalidar
        continue;
      }
      return true;
    } catch (e) {
      // se falhar ao ler, simplesmente continua
      return true;
    }
  }
}

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
    try {
      await loadAllFromFolder();
    } catch(e) { /* ignore */ }
    startFsaWatcher();
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

// ===== Indicação de status de sincronização do Banco de Dados =====
/**
 * Atualiza o banner de status do banco de dados no topo da aplicação.
 * Quando `synced` for verdadeiro, exibe mensagem positiva; caso contrário, exibe mensagem de não sincronizado.
 */
function updateDBStatusBanner(synced){
  const el = document.getElementById('dbStatusBanner');
  if(!el) return;
  if(synced){
    el.textContent = 'Banco de Dados: Sincronizado';
    el.classList.add('ok');
    el.classList.remove('warning');
  } else {
    el.textContent = 'Banco de Dados: Não sincronizado';
    el.classList.remove('ok');
    el.classList.add('warning');
  }
}

// ===== Gerenciar caminho de BD em rede (campo de texto) =====
// Permite ao usuário salvar um caminho (string) para o banco de dados em rede.
// Este caminho é persistido em localStorage e pode ser copiado para a área de transferência.
const bdPathInput = document.getElementById('bdPathInput');
const btnSavePath = document.getElementById('btnSavePath');
const btnCopyPath = document.getElementById('btnCopyPath');

// Carrega o caminho salvo no carregamento inicial do aplicativo.
if (bdPathInput) {
  try {
    const saved = window.localStorage.getItem('rp_bd_path') || '';
    bdPathInput.value = saved;
  } catch (e) {
    console.warn('Não foi possível ler rp_bd_path', e);
  }
}

// Salva o caminho informado pelo usuário em localStorage
if (btnSavePath && bdPathInput) {
  btnSavePath.onclick = () => {
    const path = (bdPathInput.value || '').trim();
    try {
      window.localStorage.setItem('rp_bd_path', path);
      alert('Caminho salvo!');
    } catch (e) {
      alert('Não foi possível salvar o caminho no armazenamento local.');
      console.warn('Erro ao salvar rp_bd_path', e);
    }
  };
}

// Copia o caminho salvo/atual para a área de transferência
if (btnCopyPath && bdPathInput) {
  btnCopyPath.onclick = async () => {
    const path = (bdPathInput.value || '').trim();
    if (!path) {
      alert('Nenhum caminho definido.');
      return;
    }
    try {
      await navigator.clipboard.writeText(path);
      alert('Caminho copiado!');
    } catch (e) {
      alert('Não foi possível copiar para a área de transferência. Por favor, copie manualmente.');
      console.warn('Erro ao copiar rp_bd_path', e);
    }
  };
}

const btnPickFolder=document.getElementById('btnPickFolder');
const btnSaveNow=document.getElementById('btnSaveNow');
const btnReloadFromFolder=document.getElementById('btnReloadFromFolder');
if(btnPickFolder) btnPickFolder.onclick=()=>fsaPickFolder();
if(btnSaveNow) btnSaveNow.onclick=()=>saveAllToFolder();
if(btnReloadFromFolder) btnReloadFromFolder.onclick=()=>loadAllFromFolder();

(async()=>{
  if(await hasFSA()){
    dirHandle = await fsaLoadHandle();
    updateFolderStatus();
    // Sempre atualizar status do BD para não sincronizado por padrão
    updateDBStatusBanner(false);
    if(dirHandle){
      try{ await loadAllFromFolder(); }catch{ /* ignore */ }
      startFsaWatcher();
    }
    // Carregar BD padrão previamente configurado (se houver)
    try {
      const savedBD = await idbGet(FSA_DB, FSA_STORE, 'bd');
      if(savedBD){
        bdHandle = savedBD;
        try {
          // Tentar ler o arquivo salvo e carregar dados
          const file = await bdHandle.getFile();
          bdFileName = file.name || '';
          bdFileExt = (bdFileName.split('.').pop() || '').toLowerCase();
          const text = await file.text();
          let parsed;
          if(bdFileExt === 'csv') parsed = parseCSVBDUnico(text);
          else parsed = parseHTMLBDTables(text);
          resources = (parsed.recursos || []).map(coerceResource);
          activities = (parsed.atividades || []).map(coerceActivity);
          // horas/cfg/feriados se disponíveis
          if (parsed.horas && typeof window.setHorasExternosData === 'function') {
            window.setHorasExternosData(parsed.horas);
          }
          if (parsed.cfg && typeof window.setHorasExternosConfig === 'function') {
            window.setHorasExternosConfig(parsed.cfg);
          }
          if (parsed.feriados && typeof window.setFeriados === 'function') {
            window.setFeriados(parsed.feriados);
          }
          const newTrails = {};
          (parsed.historico || []).forEach(h => {
            const id = h.activityId;
            if (!id) return;
            if (!newTrails[id]) newTrails[id] = [];
            newTrails[id].push({
              ts: h.timestamp,
              oldInicio: h.oldInicio,
              oldFim: h.oldFim,
              newInicio: h.newInicio,
              newFim: h.newFim,
              justificativa: h.justificativa,
              user: h.user
            });
          });
          trails = newTrails;
          // Ao carregar BD padrão no início, limpar eventLog e snapshot para evitar
          // reaplicar eventos antigos e trazer dados fora do BD.
          resetEventLogAndSnapshot();
          saveLS(LS.res, resources);
          saveLS(LS.act, activities);
          saveLS(LS.trail, trails);
          renderAll();
          updateBDStatus('BD carregado do padrão: ' + bdFileName);
          // armazena o lastModified atual como referência de última versão carregada
          try {
            const fileLm = file.lastModified;
            if (typeof window !== 'undefined') {
              window.__bdLastWrite = fileLm;
            }
          } catch(e){}
          startBDWatcher();
          updateDBStatusBanner(true);
        } catch (e) {
          console.warn('Falha ao carregar BD padrão', e);
          updateBDStatus('Falha ao carregar BD padrão');
          updateDBStatusBanner(false);
        }
      }
    } catch(e){
      console.warn('Erro ao recuperar BD padrão', e);
    }
  } else {
    const el=document.getElementById('folderStatus'); if(el) el.textContent='(navegador sem suporte de salvar em pasta — usando armazenamento do navegador)';
    // Também informar status de BD como não sincronizado
    updateDBStatusBanner(false);
  }
})();

const _saveLS_orig = saveLS;
saveLS = function(k,v){
  _saveLS_orig(k,v);
  if(dirHandle) { try{ saveAllToFolder(); }catch{} }
};

// ===== Dados iniciais =====
const today=clampDate(new Date());
const sampleResources = [];
const sampleActivities = [];

try {
  const VERSION_KEY = 'rv-version';
  const CUR_VERSION = '3';
  const storedV = localStorage.getItem(VERSION_KEY);
  if (!storedV || storedV < CUR_VERSION) {
    localStorage.removeItem(LS.res);
    localStorage.removeItem(LS.act);
    localStorage.removeItem(LS.trail);
    localStorage.removeItem('rv-enhancer-v1');
    localStorage.setItem(VERSION_KEY, CUR_VERSION);
  }
} catch (e) {
  // Ignorar erros de armazenamento
}

let resources = loadLS(LS.res, sampleResources);
// Garante que todos os recursos tenham campos de versionamento e exclusão
resources = (resources || []).map(r => {
  const nowTs = Date.now();
  return {
    ...r,
    version: typeof r.version === 'number' ? r.version : 1,
    updatedAt: typeof r.updatedAt === 'number' ? r.updatedAt : nowTs,
    deletedAt: r.deletedAt || null
  };
});
let activities = loadLS(LS.act, sampleActivities);
// Garante que todas as atividades tenham campos de versionamento e exclusão
activities = (activities || []).map(a => {
  const nowTs = Date.now();
  return {
    ...a,
    version: typeof a.version === 'number' ? a.version : 1,
    updatedAt: typeof a.updatedAt === 'number' ? a.updatedAt : nowTs,
    deletedAt: a.deletedAt || null
  };
});
let trails=loadLS(LS.trail,{});
function saveTrails(){ saveLS(LS.trail, trails); }
function addTrail(atividadeId, entry){
  if(!trails[atividadeId]) trails[atividadeId]=[];
  trails[atividadeId].push(entry);
  saveTrails();
}

// Após carregar arrays de recursos/atividades e garantir campos de versionamento,
// aplicamos snapshot e eventos do log para reconstruir o estado.  Isso
// preserva compatibilidade: se não houver snapshot ou log, nada é alterado.
loadSnapshotAndEvents();

// ===== Estado dos filtros =====
const selectedStatus=new Set(STATUS);
let filtroTipo="";
let filtroSenioridade="";
let buscaTitulo="";
let buscaTag = "";
let buscaRecurso="";

// Ao iniciar a aplicação, definir o início da visão para a data atual (hoje) ao invés de duas semanas antes.
let rangeStart=toYMD(today);
let rangeEnd=toYMD(addDays(today,60));

// ===== UI refs =====
const statusChips=document.getElementById("statusChips");
const tipoSel=document.getElementById("tipoSel");
const senioridadeSel=document.getElementById("senioridadeSel");
const buscaTituloInput=document.getElementById("buscaTitulo");
const buscaTagInput = document.getElementById("buscaTag");
const buscaRecursoInput=document.getElementById("buscaRecurso");
const inicioVisao=document.getElementById("inicioVisao");
const fimVisao=document.getElementById("fimVisao");

// As referências às tabelas são mantidas para compatibilidade com outros scripts (enhancer.js)
// mas a função renderTables agora usa os contêineres de cards.
const tblRecursos=document.querySelector("#recursos-container"); // Anteriormente '#tblRecursos tbody'
const tblAtividades=document.querySelector("#atividades-container"); // Anteriormente '#tblAtividades tbody'
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
buscaTagInput.oninput = () => { buscaTag = buscaTagInput.value.toLowerCase(); renderAll(); };
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
document.getElementById("btnSalvarRecurso").onclick=()=>{
  const f=formRecurso.elements;
  // Build resource object with versioning and timestamps
  const idValue = f["id"].value || uuid();
  const existingIndex = resources.findIndex(r => r.id === idValue);
  const nowTs = Date.now();
  let version = 1;
  let deletedAt = null;
  if (existingIndex >= 0) {
    const existing = resources[existingIndex];
    version = (existing.version || 0) + 1;
    deletedAt = existing.deletedAt || null;
  }
  const rec={
    id: idValue,
    nome:f["nome"].value.trim(),
    tipo:f["tipo"].value,
    senioridade:f["senioridade"].value,
    ativo:f["ativo"].checked,
    capacidade:Math.max(1,Number(f["capacidade"].value||100)),
    inicioAtivo:f["inicioAtivo"].value||"",
    fimAtivo:f["fimAtivo"].value||"",
    version: version,
    updatedAt: nowTs,
    deletedAt: deletedAt
  };
  if(!rec.nome){alert("Informe o nome.");return}
  if(existingIndex>=0) resources[existingIndex]=rec; else resources.push(rec);
  saveLS(LS.res,resources);
  // Registra evento no log: criação ou atualização de recurso
  recordEvent('resource', existingIndex>=0 ? 'update' : 'create', rec.id, rec);
  dlgRecurso.close();
  renderAll();
  saveBDDebounced();
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
/**
 * Preenche a lista de opções de recursos (datalist) para o campo de seleção de recurso na criação/edição de atividade.
 * Em vez de um <select>, utilizamos um <input> com datalist para permitir digitação e sugestão automática.
 */
function fillRecursoOptions(){
  const list = document.getElementById('resourceList');
  if(!list) return;
  list.innerHTML = '';
  resources.forEach(r=>{
    if(!r.ativo) return;
    const opt = document.createElement('option');
    opt.value = r.nome;
    list.appendChild(opt);
  });
}
document.getElementById("btnSalvarAtividade").onclick=()=>{
  const f=formAtividade.elements;

  const tagsInput = f["tags"].value || '';
  const tags = tagsInput.split(',')
                       .map(normalizeTag)
                       .filter(t => t);

  // Obtém o nome do recurso digitado e valida sua existência
  const resourceName = (f["resourceName"].value || '').trim();
  const recByName = resources.find(r => (r.nome || '').toLowerCase() === resourceName.toLowerCase());

  // Prepare id and versioning
  const atId = f["id"].value || uuid();
  const existingAtIndex = activities.findIndex(a => a.id === atId);
  const nowAtTs = Date.now();
  let atVersion = 1;
  let atDeletedAt = null;
  if (existingAtIndex >= 0) {
    const existingA = activities[existingAtIndex];
    atVersion = (existingA.version || 0) + 1;
    atDeletedAt = existingA.deletedAt || null;
  }
  const at={
    id: atId,
    titulo:f["titulo"].value.trim(),
    // ResourceId será atribuído a partir do recurso encontrado pelo nome
    resourceId: recByName ? recByName.id : '',
    inicio:f["inicio"].value,
    fim:f["fim"].value,
    status:f["status"].value,
    alocacao:Math.max(1,Number(f["alocacao"].value||100)),
    tags: [...new Set(tags)],
    version: atVersion,
    updatedAt: nowAtTs,
    deletedAt: atDeletedAt
  };
  if(!at.titulo) return alert("Informe o título.");
  if(!resourceName) return alert("Selecione o recurso.");
  if(!recByName) return alert("O nome do recurso não corresponde a nenhum recurso cadastrado.");
  if(fromYMD(at.fim)<fromYMD(at.inicio)) return alert("Fim não pode ser menor que início.");

  const rec = recByName;
  if(rec){
    if(rec.inicioAtivo && fromYMD(at.inicio) < fromYMD(rec.inicioAtivo)){
      return alert(`Início da atividade (${at.inicio}) menor que início ativo do recurso (${rec.inicioAtivo}).`);
    }
    if(rec.fimAtivo && fromYMD(at.fim) > fromYMD(rec.fimAtivo)){
      return alert(`Fim da atividade (${at.fim}) maior que fim ativo do recurso (${rec.fimAtivo}).`);
    }
  }
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
      window.__pendingAt=at;
      window.__pendingIdx=idx;
      justResumo.textContent=`${prev.titulo} — Início: ${prev.inicio} → ${at.inicio} | Fim: ${prev.fim} → ${at.fim}`;
      formJust.elements["just"].value="";
      dlgJust.showModal();
      return; 
    } else {
      activities[idx]=at;
      saveLS(LS.act,activities);
      // Registra evento de atualização de atividade
      recordEvent('activity','update', at.id, at);
      dlgAtividade.close();
      renderAll();
      saveBDDebounced();
    }
  } else {
    activities.push(at);
    saveLS(LS.act,activities);
    // Registra evento de criação de atividade
    recordEvent('activity','create', at.id, at);
    dlgAtividade.close();
    renderAll();
    saveBDDebounced();
  }
};

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
  // Registra evento de atualização após justificativa de alteração de datas
  recordEvent('activity','update', at.id, at);
  dlgJust.close();
  dlgAtividade.close();
  renderAll();
  saveBDDebounced();
};

// ===== NOVA FUNÇÃO renderTables PARA LAYOUT DE CARDS =====
function renderTables(filteredActs){
  const recursosContainer = document.getElementById('recursos-container');
  const atividadesContainer = document.getElementById('atividades-container');
  if (!recursosContainer || !atividadesContainer) return;

  // ===== Renderização dos Cards de Recursos =====
  recursosContainer.innerHTML = "";
  // Filtra recursos de acordo com os mesmos critérios aplicados às atividades.
  // Além dos filtros de tipo/senioridade/nomes, exibe apenas recursos que tenham
  // pelo menos uma atividade no conjunto filtrado (filteredActs). Isso garante
  // que a listagem de recursos reflita fielmente a seleção feita nas buscas.
  const recursosComAtividades = new Set(filteredActs.map(a => a.resourceId));
  const visibleResources = resources.filter(r => {
    // Ignorar recursos marcados como excluídos
    if(r.deletedAt) return false;
    // Mostrar apenas recursos ativos
    if(!r.ativo) return false;
    // Filtrar por tipo (Interno/Externo) se definido
    if(filtroTipo && (r.tipo||'').toLowerCase() !== filtroTipo.toLowerCase()) return false;
    // Filtrar por senioridade se definido
    if(filtroSenioridade && r.senioridade !== filtroSenioridade) return false;
    // Filtrar por busca de recurso se definido
    if(buscaRecurso && !(r.nome || '').toLowerCase().includes(buscaRecurso)) return false;
    // Se houver atividades filtradas, normalmente o componente escondia recursos
    // que não possuíssem nenhuma atividade visível. Entretanto, isso impedia
    // editar/inativar/excluir recursos sem atividades. Agora, sempre exibe
    // recursos que passam pelos demais filtros, mesmo que não tenham
    // nenhuma atividade no conjunto filtrado. Isso permite manipular
    // recursos "órfãos" de forma explícita.
    return true;
  });
  const groupedResources = visibleResources.reduce((acc, r) => {
    const tipo = r.tipo || 'Outros';
    if (!acc[tipo]) acc[tipo] = [];
    acc[tipo].push(r);
    return acc;
  }, {});

  // Ordem desejada: Interno, Externo, e depois outros
  const groupOrder = ['Interno', 'Externo'];
  const sortedGroups = [...Object.keys(groupedResources)].sort((a, b) => {
    const aIndex = groupOrder.indexOf(a);
    const bIndex = groupOrder.indexOf(b);
    if (aIndex > -1 && bIndex > -1) return aIndex - bIndex;
    if (aIndex > -1) return -1;
    if (bIndex > -1) return 1;
    return a.localeCompare(b);
  });

  for (const tipo of sortedGroups) {
    const group = groupedResources[tipo];
    const details = document.createElement('details');
    details.className = 'resource-group';
    details.open = true; // Inicia aberto

    const summary = document.createElement('summary');
    summary.textContent = tipo + ` (${group.length})`;
    details.appendChild(summary);

    group.forEach(r => {
      const card = document.createElement('div');
      card.className = `resource-card type-${tipo.toLowerCase()}`;
      card.innerHTML = `
        <div class="card-header">
          <strong class="resource-name">👤 ${r.nome}</strong>
          <div class="card-actions">
            <button class="btn-icon edit" title="Editar"></button>
            <button class="btn-icon duplicate" title="Duplicar"></button>
            <button class="btn-icon delete" title="Excluir"></button>
          </div>
        </div>
        <div class="card-body">
          <span class="chip">${r.senioridade}</span>
          <span class="chip">${r.ativo ? "Ativo" : "Inativo"}</span>
          <span class="chip">📊 Cap: ${r.capacidade || 100}%</span>
        </div>
        ${(r.inicioAtivo || r.fimAtivo) ? `<div class="card-footer muted small">📅 ${r.inicioAtivo || '...'} → ${r.fimAtivo || '...'}</div>` : ''}
      `;
      // Adiciona eventos aos botões de ação
      card.querySelector('.edit').onclick = () => {
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
      card.querySelector('.duplicate').onclick = () => {
        const nowTs = Date.now();
        const copy={...r,id:uuid(),nome:"Cópia de "+r.nome, version:1, updatedAt: nowTs, deletedAt:null};
        resources.push(copy);
        saveLS(LS.res,resources);
        // Registra evento de criação (duplicação) de recurso
        recordEvent('resource','create',copy.id, copy);
        renderAll();
        // Persiste no BD assíncrono para evitar reaparecimento após sincronização
        saveBDDebounced();
      };
      card.querySelector('.delete').onclick = () => {
        if(!confirm("Remover recurso e suas alocações?")) return;
        const nowTs = Date.now();
        // marca recurso como deletado sem removê-lo
        resources = resources.map(item => {
          if(item.id === r.id){
            const v = (item.version || 0) + 1;
            return { ...item, deletedAt: nowTs, updatedAt: nowTs, version: v };
          }
          return item;
        });
        // marca todas as atividades desse recurso como deletadas
        activities = activities.map(item => {
          if(item.resourceId === r.id){
            const v = (item.version || 0) + 1;
            return { ...item, deletedAt: nowTs, updatedAt: nowTs, version: v };
          }
          return item;
        });
        saveLS(LS.res, resources);
        saveLS(LS.act, activities);
        // Registra eventos de deleção: recurso e suas atividades
        try {
          recordEvent('resource','delete', r.id, { deletedAt: nowTs });
          activities.forEach(item => {
            if(item.resourceId === r.id && item.deletedAt === nowTs) {
              recordEvent('activity','delete', item.id, { deletedAt: nowTs });
            }
          });
        } catch (e) {
          console.error('Erro ao registrar eventos de deleção', e);
        }
        renderAll();
        // Persiste no BD assíncrono para evitar reaparecimento após sincronização
        saveBDDebounced();
      };
      details.appendChild(card);
    });
    recursosContainer.appendChild(details);
  }

  // ===== Renderização dos Cards de Atividades =====
  atividadesContainer.innerHTML = "";
  filteredActs.forEach(a => {
    const r = resources.find(x => x.id === a.resourceId);
    const card = document.createElement('div');
    card.className = 'activity-card';
    const statusClass = (a.status || '').toLowerCase().replace(/\s+/g, '-');

    const tagsHtml = (a.tags && a.tags.length)
      ? `<div class="tags-container">${a.tags.map(t => `<span class="chip tag">${t}</span>`).join(' ')}</div>`
      : '';

    card.innerHTML = `
      <div class="card-header">
        <strong class="activity-title">${a.titulo}</strong>
        <div class="card-actions">
          <button class="btn-icon edit" title="Editar"></button>
          <button class="btn-icon history" title="Histórico"></button>
          <button class="btn-icon duplicate" title="Duplicar"></button>
          <button class="btn-icon delete" title="Excluir"></button>
        </div>
      </div>
      <div class="card-body">
        <div class="activity-meta">
          <span class="status-badge status-${statusClass}">${a.status}</span>
          <span class="muted small">👤 ${r ? r.nome : 'N/A'}</span>
        </div>
        <div class="allocation-bar-container" title="Alocação: ${a.alocacao || 100}%">
          <div class="allocation-bar" style="width: ${Math.min(a.alocacao || 100, 100)}%;"></div>
          ${(a.alocacao || 100) > 100 ? '<div class="allocation-overload"></div>' : ''}
        </div>
        ${tagsHtml}
      </div>
      <div class="card-footer muted small">
        📅 ${a.inicio} → ${a.fim}
      </div>
    `;
    // Adiciona eventos aos botões de ação
    card.querySelector('.edit').onclick = () => {
      dlgAtividadeTitulo.textContent="Editar Atividade";
      fillRecursoOptions();
      formAtividade.elements["id"].value=a.id;
      formAtividade.elements["titulo"].value=a.titulo;
      // Preenche campo de recurso com o nome correspondente
      const recObj = resources.find(x=>x.id===a.resourceId);
      formAtividade.elements["resourceName"].value = recObj ? recObj.nome : '';
      formAtividade.elements["inicio"].value=a.inicio;
      formAtividade.elements["fim"].value=a.fim;
      formAtividade.elements["status"].value=a.status;
      formAtividade.elements["alocacao"].value=a.alocacao||100;
      formAtividade.elements["tags"].value = (a.tags || []).join(', ');
      dlgAtividade.showModal();
    };
    card.querySelector('.history').onclick = () => {
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
      const btn=document.getElementById('histApply'); if(btn){ btn.onclick=()=>{ card.querySelector('.history').onclick(); }; }
    };
    card.querySelector('.duplicate').onclick = () => {
      const nowTs = Date.now();
      const copy={...a,id:uuid(),titulo:"Cópia de "+a.titulo, version:1, updatedAt: nowTs, deletedAt:null};
      activities.push(copy);
      saveLS(LS.act,activities);
      // Registra evento de criação de atividade (cópia)
      recordEvent('activity','create', copy.id, copy);
      renderAll();
      // Persiste no BD assíncrono para evitar reaparecimento após sincronização
      saveBDDebounced();
    };
    card.querySelector('.delete').onclick = () => {
      if(!confirm("Remover atividade?")) return;
      const nowTs = Date.now();
      activities = activities.map(item => {
        if(item.id === a.id){
          const v = (item.version || 0) + 1;
          return { ...item, deletedAt: nowTs, updatedAt: nowTs, version: v };
        }
        return item;
      });
      saveLS(LS.act,activities);
      // Registra evento de deleção de atividade
      try {
        recordEvent('activity','delete', a.id, { deletedAt: nowTs });
      } catch (e) {
        console.error('Erro ao registrar evento de deleção de atividade', e);
      }
      renderAll();
      // Persiste no BD assíncrono para evitar reaparecimento após sincronização
      saveBDDebounced();
    };
    atividadesContainer.appendChild(card);
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
  // -------------------------------------------------------------------------
  // Header row showing month names.  We display a short month label
  // ("jan", "fev", "mar", etc.) on the first cell of each month and leave
  // other cells blank.  This helps users identify the month boundaries
  // without introducing confusing year labels (e.g. "fev. 25").  If your
  // locale is set to Portuguese, this will show month abbreviations in
  // Portuguese; otherwise it will use the browser's default locale.
  const rowMonths=document.createElement("div");
  rowMonths.className="row-months";
  rowMonths.style.display="grid";
  rowMonths.style.gridTemplateColumns=`repeat(${days.length}, 28px)`;
  days.forEach((d,i)=>{
    const isFirstOfMonth=d.getDate()===1 || i===0;
    const cell=document.createElement("div");
    cell.className="cell-day";
    cell.style.fontWeight=isFirstOfMonth?"600":"400";
    // Use only the short month name (no year) when this cell marks
    // the beginning of a month.  For example, on 1st February it will
    // show "fev".  Leaving other cells blank preserves alignment.
    cell.textContent=isFirstOfMonth? d.toLocaleDateString(undefined,{month:"short"}):"";
    rowMonths.appendChild(cell);
  });
  const rowDays=document.createElement("div");
  rowDays.style.display="grid";
  rowDays.style.gridTemplateColumns=`repeat(${days.length}, 28px)`;
  days.forEach(d=>{
    const cell=document.createElement("div");
    cell.className="cell-day";
    cell.textContent=String(d.getDate()).padStart(2,"0");
    // Marcar finais de semana com uma classe dedicada e aplicar uma cor
    // de fundo mais clara para distinguir visualmente sábados e domingos.
    const dow = d.getDay();
    if(dow === 0 || dow === 6){
      cell.classList.add("weekend");
      // Cor de fundo mais suave para cabeçalho de finais de semana
      cell.style.background = "#f8fafc";
    }
    rowDays.appendChild(cell);
  });
  // Append the month row (showing only month names) and then the row of
  // day numbers.  The month row helps users understand which month the
  // numeric day sequence belongs to without showing the year.
  right.appendChild(rowMonths);
  right.appendChild(rowDays);
  header.appendChild(left); header.appendChild(right);
  gantt.appendChild(header);

  const byRes=Object.fromEntries(resources.map(r=>[r.id,[]]));
  filteredActs.forEach(a=>{ if(byRes[a.resourceId]) byRes[a.resourceId].push(a); });
  Object.keys(byRes).forEach(k=>byRes[k].sort((a,b)=>fromYMD(a.inicio)-fromYMD(b.inicio)));

  resources.forEach(r=>{
    // Aplicar filtros básicos: tipo (Interno/Externo), senioridade, ativo e busca por nome do recurso.
    if(filtroTipo && (r.tipo||"").toLowerCase() !== filtroTipo.toLowerCase()) return;
    if(filtroSenioridade && r.senioridade!==filtroSenioridade) return;
    if(!r.ativo) return;
    if(buscaRecurso && !(r.nome||'').toLowerCase().includes(buscaRecurso)) return;
    const acts=byRes[r.id]||[];
    // Se após aplicar todos os filtros de atividades não restaram tarefas para este recurso,
    // não exibir a linha no Gantt. Isso garante que somente recursos com atividades
    // correspondentes ao filtro apareçam.
    if(acts.length===0) return;

    const row=document.createElement("div");
    row.className="row";

    const info=document.createElement("div");
    info.className="info";
    info.innerHTML=`<div style="font-weight:600">${r.nome}</div>
      <div class="muted" style="font-size:12px">${r.tipo} • ${r.senioridade} • Cap: ${r.capacidade}%${(r.inicioAtivo||r.fimAtivo)? " • Janela: " + (r.inicioAtivo||"…") + " → " + (r.fimAtivo||"…") : ""}</div>`;

    const bargrid=document.createElement("div");
    bargrid.className="bargrid";

    const cap=r.capacidade||100;
    days.forEach((d,i)=>{
      const dy=toYMD(d);
      const activeActs = acts.filter(a=>fromYMD(a.inicio)<=d && d<=fromYMD(a.fim));
      const sum=activeActs.reduce((acc,a)=>acc+(a.alocacao||100),0);
      const perc=cap? (sum/cap)*100 : 0;
      const heat=document.createElement("div");
      heat.className="heatcell";
      heat.style.left=`${i*28}px`; heat.style.width="28px";
      // Garante que os blocos de ocupação fiquem acima do plano de fundo da grade
      heat.style.zIndex = "1";
      // A classe "weekend" não é aplicada aos blocos de ocupação (heatcell).
      // A marcação visual de finais de semana é feita apenas no fundo da grade (gridBg).
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
      const c=activeActs.length;
      if(c>1){
        const cc=document.createElement("div");
        cc.className="ccell "+(c>=4?"high":(c>=3?"med":"low"));
        cc.style.left=`${i*28}px`; cc.style.width="28px";
        cc.textContent=String(c);
        cc.title=`${c} atividades simultâneas`;
        // Garante que o contador fique acima do plano de fundo da grade
        cc.style.zIndex = "1";
        bargrid.appendChild(cc);
      }
    });

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
    const lanes = [];
    const placed = intervals.sort((x,y)=>x.sIdx - y.sIdx).map(iv=>{
      let lane = 0;
      while(lane < lanes.length && !(lanes[lane] < iv.sIdx - 0)) lane++;
      if(lane === lanes.length) lanes.push(-Infinity);
      lanes[lane] = iv.eIdx;
      return {...iv, lane};
    });

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
      // As lacunas devem ficar acima do plano de fundo da grade
      el.style.zIndex = "1";
      bargrid.appendChild(el);
    });

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
        // Preenche campo de recurso com o nome correspondente
        const recObj = resources.find(x=>x.id===a.resourceId);
        formAtividade.elements["resourceName"].value = recObj ? recObj.nome : '';
        formAtividade.elements["inicio"].value=a.inicio;
        formAtividade.elements["fim"].value=a.fim;
        formAtividade.elements["status"].value=a.status;
        formAtividade.elements["alocacao"].value=a.alocacao||100;
        formAtividade.elements["tags"].value = (a.tags || []).join(', ');
        dlgAtividade.showModal();
      };
      // As atividades devem estar sempre acima do plano de fundo da grade e de outros elementos
      b.style.zIndex = "2";
      bargrid.appendChild(b);
    });

    const lanesH = Math.max(42, lanes.length*22 + 6);
    bargrid.style.height = lanesH + "px";

    const gridBg=document.createElement("div");
    gridBg.style.position="absolute"; gridBg.style.top="0"; gridBg.style.bottom="0"; gridBg.style.left="0"; gridBg.style.right="0";
    gridBg.style.display="grid"; gridBg.style.gridTemplateColumns=`repeat(${days.length}, 28px)`; gridBg.style.pointerEvents="none";
    // Coloca o fundo da grade atrás de outros elementos
    gridBg.style.zIndex = "0";
    for(let i=0;i<days.length;i++){
      const v=document.createElement("div");
      v.style.borderLeft="1px solid #f1f5f9";
      // Aplica uma tonalidade mais clara ao fundo das colunas de finais de semana
      const dt = days[i];
      const dwd = dt.getDay();
      if(dwd === 0 || dwd === 6){
        v.style.background = "#f8fafc";
      }
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
    // Ignorar atividades excluídas
    if(a.deletedAt) return false;
    if(!selectedStatus.has((a.status||"").trim())) return false;
    const r=resources.find(x=>x.id===a.resourceId);
    if(!r) return false;
    // também ignorar se recurso estiver marcado como deletado
    if(r.deletedAt) return false;
    if(filtroTipo && (r.tipo||"").toLowerCase() !== filtroTipo.toLowerCase()) return false;
    if(filtroSenioridade && r.senioridade!==filtroSenioridade) return false;
    if(buscaRecurso && !(r.nome||"").toLowerCase().includes(buscaRecurso)) return false;
    if(buscaTitulo && !a.titulo.toLowerCase().includes(buscaTitulo)) return false;
    if (buscaTag && (!a.tags || a.tags.length === 0 || !a.tags.some(t => t.toLowerCase().includes(buscaTag)))) {
        return false;
    }
    const s=fromYMD(a.inicio), e=fromYMD(a.fim);
    if(e<fromYMD(rangeStart) || s>fromYMD(rangeEnd)) return false;
    return true;
  });
}

// ===== Helpers para exportações =====
// Retorna lista de recursos que não foram marcados como excluídos.  Útil para
// exportação de dados completos, evitando incluir registros tombados no CSV/XLS.
function getActiveResources() {
  return (resources || []).filter(r => !r.deletedAt);
}

// Retorna lista de atividades que não foram marcadas como excluídas e cujo
// recurso associado também está ativo.  Evita exportar itens que foram
// excluídos ou cujos recursos foram removidos.
function getActiveActivities() {
  const activeResIds = new Set(getActiveResources().map(r => r.id));
  return (activities || []).filter(a => !a.deletedAt && activeResIds.has(a.resourceId));
}

/**
 * Garante que, se houver um banco de dados externo apontado, os dados em
 * memória estejam sincronizados com o arquivo no momento da exportação.
 * Verifica a última modificação do arquivo e, se for mais recente do que
 * __bdLastWrite, recarrega o conteúdo e reaplica snapshot/eventos. Caso
 * contrário, nada faz. Essa função é assíncrona porque a leitura do arquivo
 * utiliza a API File System Access.
 */
async function refreshFromBDIfNeeded() {
  if (!bdHandle) return;
  try {
    const file = await bdHandle.getFile();
    const lm = file.lastModified;
    if (typeof window !== 'undefined' && lm > (window.__bdLastWrite || 0)) {
      let parsed;
      const text = await file.text();
      if (bdFileExt === 'csv') {
        parsed = parseCSVBDUnico(text);
      } else {
        parsed = parseHTMLBDTables(text);
      }
      const newResources = (parsed.recursos || []).map(coerceResource);
      const newActivities = (parsed.atividades || []).map(coerceActivity);
      const newTrails = {};
      (parsed.historico || []).forEach(h => {
        const id = h.activityId;
        if (!id) return;
        if (!newTrails[id]) newTrails[id] = [];
        newTrails[id].push({
          ts: h.timestamp,
          oldInicio: h.oldInicio,
          oldFim: h.oldFim,
          newInicio: h.newInicio,
          newFim: h.newFim,
          justificativa: h.justificativa,
          user: h.user
        });
      });
      // Atualiza arrays globais e persiste
      resources = newResources;
      activities = newActivities;
      trails = newTrails;
      saveLS(LS.res, resources);
      saveLS(LS.act, activities);
      saveLS(LS.trail, trails);
      // Atualiza marcador de última escrita
      window.__bdLastWrite = lm;
      // Aplica snapshot e eventos pendentes
      try {
        loadSnapshotAndEvents();
      } catch (e) {
        console.error('Erro ao aplicar snapshot/eventos após refresh BD', e);
      }
    }
  } catch (e) {
    console.warn('Falha ao sincronizar com BD antes da exportação', e);
  }
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

document.getElementById("btnExportCSV").onclick = async () => {
  // Garante que estamos sincronizados com o BD antes de exportar
  await refreshFromBDIfNeeded();
  // Seleciona apenas recursos e atividades não deletados
  const activeResources = getActiveResources();
  const activeActivities = getActiveActivities();
  const rec = activeResources.map(r => ({
    id: r.id,
    nome: r.nome,
    tipo: r.tipo,
    senioridade: r.senioridade,
    ativo: r.ativo,
    capacidade: r.capacidade || 100,
    inicioAtivo: r.inicioAtivo || "",
    fimAtivo: r.fimAtivo || "",
    version: r.version || 1,
    updatedAt: r.updatedAt || 0,
    deletedAt: r.deletedAt || ""
  }));
  const atv = activeActivities.map(a => ({
    id: a.id,
    titulo: a.titulo,
    resourceId: a.resourceId,
    inicio: a.inicio,
    fim: a.fim,
    status: a.status,
    alocacao: a.alocacao || 100,
    tags: (a.tags || []).join(', '),
    version: a.version || 1,
    updatedAt: a.updatedAt || 0,
    deletedAt: a.deletedAt || ""
  }));
  download("recursos.csv", toCSV(rec, ["id","nome","tipo","senioridade","ativo","capacidade","inicioAtivo","fimAtivo","version","updatedAt","deletedAt"]), "text/csv;charset=utf-8");
  download("atividades.csv", toCSV(atv, ["id","titulo","resourceId","inicio","fim","status","alocacao", "tags","version","updatedAt","deletedAt"]), "text/csv;charset=utf-8");
  alert("Exportados: recursos.csv e atividades.csv");
};

function tableHTML(name, rows, cols){
  const header = cols.map(c=>`<th>${c}</th>`).join("");
  const body = rows.map(r=>`<tr>${cols.map(c=>`<td>${(r[c]??"")}</td>`).join("")}</tr>`).join("");
  return `
  <html xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:x="urn:schemas-microsoft-com:office:excel"
        xmlns="http://www.w3.org/TR/REC-html40">
  <head><meta charset="utf-8">
  </head><body>
    <table border="1"><thead><tr>${header}</tr></thead><tbody>${body}</tbody></table>
  </body></html>`;
}
document.getElementById("btnExportXLS").onclick = async () => {
  await refreshFromBDIfNeeded();
  const activeResources = getActiveResources();
  const activeActivities = getActiveActivities();
  const rec = activeResources.map(r => ({
    id: r.id,
    nome: r.nome,
    tipo: r.tipo,
    senioridade: r.senioridade,
    ativo: r.ativo,
    capacidade: r.capacidade || 100,
    inicioAtivo: r.inicioAtivo || "",
    fimAtivo: r.fimAtivo || ""
  }));
  const atv = activeActivities.map(a => ({
    id: a.id,
    titulo: a.titulo,
    resourceId: a.resourceId,
    inicio: a.inicio,
    fim: a.fim,
    status: a.status,
    alocacao: a.alocacao || 100,
    tags: (a.tags || []).join(', ')
  }));
  download("recursos.xls", tableHTML("Recursos", rec, ["id","nome","tipo","senioridade","ativo","capacidade","inicioAtivo","fimAtivo"]), "application/vnd.ms-excel");
  download("atividades.xls", tableHTML("Atividades", atv, ["id","titulo","resourceId","inicio","fim","status","alocacao", "tags"]), "application/vnd.ms-excel");
  alert("Exportados: recursos.xls e atividades.xls");
};

document.getElementById("btnExportPBI").onclick = async () => {
  await refreshFromBDIfNeeded();
  const rows = [];
  const start = fromYMD(rangeStart), end = fromYMD(rangeEnd);
  // Usa apenas recursos e atividades ativos (sem deletados)
  const activeResources = getActiveResources();
  const byId = Object.fromEntries(activeResources.map(r => [r.id, r]));
  getActiveActivities().forEach(a => {
    const s = fromYMD(a.inicio), e = fromYMD(a.fim);
    const r = byId[a.resourceId];
    if (!r) return;
    for (let d = new Date(Math.max(s, start)); d <= Math.min(e, end); d = addDays(d, 1)) {
      rows.push({
        data: toYMD(d),
        atividadeId: a.id,
        atividadeTitulo: a.titulo,
        status: a.status,
        alocacao: a.alocacao || 100,
        tags: (a.tags || []).join('|'),
        recursoId: r.id,
        recursoNome: r.nome,
        recursoTipo: r.tipo,
        recursoSenioridade: r.senioridade,
        recursoCapacidade: r.capacidade || 100
      });
    }
  });
  download("powerbi_atividades_diarias.csv",
    toCSV(rows, ["data","atividadeId","atividadeTitulo","status","alocacao","tags","recursoId","recursoNome","recursoTipo","recursoSenioridade","recursoCapacidade"]),
    "text/csv;charset=utf-8");
  alert(`Exportado: powerbi_atividades_diarias.csv (${rows.length} linhas)`);
};

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

const __fileImportEl = document.getElementById("fileImport");
if(__fileImportEl) __fileImportEl.onchange=(ev)=>{
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
            alocacao: Number(cols[idx("alocacao")]||100),
            tags: (cols[idx("tags")] || "").split(',').map(t => t.trim()).filter(Boolean)
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

// Formata rótulos do eixo X para os gráficos de capacidade agregada
function formatBucketLabel(key, gran){
  if(gran === "weekly"){
    // key: "W YYYY-MM-DD" (segunda-feira da semana)
    const raw = String(key).replace(/^W\s+/, "");
    const dt = fromYMD(raw);
    const dd = String(dt.getDate()).padStart(2,"0");
    const mm = String(dt.getMonth()+1).padStart(2,"0");
    return `${dd}/${mm}`;
  }
  // monthly: "YYYY-MM"
  const m = String(key);
  const parts = m.split("-");
  if(parts.length === 2){
    return `${parts[1]}/${parts[0]}`;
  }
  return String(key);
}
// FUNÇÃO CORRIGIDA
function renderAggregates(){
  aggCharts.innerHTML="";
  const gran=aggGran.value;
  const days=buildDays();
  // Constrói o mapa de recursos a serem exibidos, ignorando recursos inativos ou marcados como excluídos
  const byRes = Object.fromEntries(
    resources
      .filter(r => r.ativo && !r.deletedAt)
      .map(r => [r.id, {}])
  );

  // Pré-calcula o número de dias em cada "bucket" (semana/mês) para evitar contagem duplicada da capacidade
  const daysPerBucket = {};
  days.forEach(d => {
    const key = bucketKey(d, gran);
    if (!daysPerBucket[key]) daysPerBucket[key] = 0;
    daysPerBucket[key]++;
  });

  // Calcula a soma total de alocação (numerador)
  activities.forEach(a => {
    // Ignora atividades removidas
    if (a.deletedAt) return;
    // Busca o recurso associado, certificando-se que ele está ativo e não excluído
    const r = resources.find(x => x.id === a.resourceId && x.ativo && !x.deletedAt);
    if (!r) return;
    for (let d of days) {
      if (fromYMD(a.inicio) <= d && d <= fromYMD(a.fim)) {
        const key = bucketKey(d, gran);
        const map = byRes[r.id];
        if (!map[key]) map[key] = { sum: 0, capDays: 0 }; // Inicializa
        map[key].sum += (a.alocacao || 100);
      }
    }
  });

  // Calcula a capacidade total correta (denominador), contando cada dia apenas uma vez
  Object.keys(byRes).forEach(resourceId => {
    const resource = resources.find(r => r.id === resourceId);
    if (!resource) return;
    const cap = resource.capacidade || 100;
    const map = byRes[resourceId];
    Object.keys(map).forEach(key => {
      const numDays = daysPerBucket[key] || 0;
      map[key].capDays = numDays * cap;
    });
  });
  
  // Renderiza os gráficos
  resources
    .filter(r => r.ativo && !r.deletedAt)
    .forEach(r => {
    const card=document.createElement("div");
    card.className="card";
    const h=document.createElement("h3");
    h.textContent=`${r.nome} — ${gran==="weekly"?"Semanal":"Mensal"}`;
    const canvas=document.createElement("canvas");
    // Ajuste de eixo X (legendas): mais próximo do gráfico e sem recorte.
    // Em vez de empurrar as datas muito para baixo e rotacionar, deixamos
    // a legenda mais próxima e aplicamos "pulo" de rótulos quando há muitos.
    canvas.width=600; canvas.height=180; canvas.className="chart";
    const ctx=canvas.getContext("2d");
    const entries=Object.entries(byRes[r.id]||{});
    entries.sort((a,b)=>a[0]>b[0]?1:-1);
    const W=canvas.width, H=canvas.height;
    const marginL = 34;
    const marginT = 12;
    const marginB = 48; // espaço para legenda sem afastar demais do gráfico
    const axisY = H - marginB;
    ctx.clearRect(0,0,W,H);
    ctx.beginPath();
    ctx.moveTo(marginL, marginT);
    ctx.lineTo(marginL, axisY);
    ctx.lineTo(W-10, axisY);
    ctx.stroke();
    const barW=Math.max(8, (W - marginL - 20) / Math.max(1, entries.length) - 6);
    // Se houver muitos pontos, mostramos apenas 1 a cada N rótulos para evitar sobreposição.
    const labelStep = entries.length > 20 ? 3 : (entries.length > 12 ? 2 : 1);
    entries.forEach((kv,idx)=>{
      const key=kv[0]; const v=kv[1];
      const label = formatBucketLabel(key, gran);
      const perc = v.capDays > 0 ? (v.sum / v.capDays) * 100 : 0;
      const x = marginL + 10 + idx*(barW+6);
      const y = axisY - (Math.min(100, perc)/100)*(axisY - marginT - 18); // Limita a barra em 100% de altura visualmente
      ctx.fillStyle = perc > 100 ? '#ef4444' : '#2563eb'; // Cor vermelha para sobrecarga
      ctx.fillRect(x, y, barW, axisY - y);
      // Legenda do eixo X (datas): mais próxima do eixo e sem recorte.
      // Para não embolar, exibimos apenas 1 a cada "labelStep" quando necessário.
      if(idx % labelStep === 0){
        ctx.save();
        // Mantém a legenda próxima do eixo, mas com folga suficiente para não recortar.
        ctx.translate(x + barW/2, axisY + 16);
        ctx.textAlign="center";
        ctx.textBaseline="top";
        ctx.font="9px sans-serif";
        ctx.fillText(label, 0, 0);
        ctx.restore();
      }
      ctx.font="10px sans-serif"; ctx.fillText(Math.round(perc)+"%", x, y-4);
    });
    card.appendChild(h); card.appendChild(canvas);
    aggCharts.appendChild(card);
  });
}

// ===== Auto-sugestão de Tags =====
function updateTagDatalist() {
    const allTags = new Set();
    activities.forEach(a => {
        if (a.tags) {
            a.tags.forEach(t => allTags.add(t));
        }
    });
  
    const datalist = document.getElementById('existingTags');
    if (datalist) {
        datalist.innerHTML = '';
        allTags.forEach(tag => {
            const option = document.createElement('option');
            option.value = tag;
            datalist.appendChild(option);
        });
    }
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
    const wd = d.getDay();
    return wd>=1 && wd<=5;
  }

  function sumAllocationOn(resourceId, date) {
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
      if(tipo && (r.tipo||"").toLowerCase() !== tipo.toLowerCase()) return;
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
        const acts=activities.filter(a=>a.resourceId===r.id && a.status !== 'Concluída' && a.status !== 'Cancelada' && fromYMD(a.inicio)<=d && d<=fromYMD(a.fim));
        const sum=acts.reduce((acc,a)=>acc+(a.alocacao||100),0);
        if(sum>cap){sobre++; break;}
      }
    });
    const el3=document.getElementById("kpiSobrecarga"); if(el3) el3.textContent=sobre;
    try{
      renderOverloadDetails();
    }catch(err){ console.error(err); }
  }catch(e){ /* ignora falhas */ }
}

// ===== Risco por Recurso (Fases 1 & 3) =====
/**
 * Calcula os scores de risco para recursos ativos. Para cada recurso, conta quantos dias do período
 * visualizado o total de alocação ultrapassa a capacidade (dias de sobrecarga). Adiciona uma
 * penalidade fixa de 5 pontos se o recurso for do tipo Externo. Retorna uma lista de objetos
 * {recurso, tipo, score, dias} apenas para recursos com dias>0.
 */
function computeRiskScores(){
  const out=[];
  try{
    const days=buildDays();
    resources.filter(r=>r.ativo).forEach(r=>{
      const cap=r.capacidade||100;
      let overload=0;
      days.forEach(d=>{
        let sum=0;
        activities.forEach(a=>{
          if(a.resourceId===r.id && a.status!=='Concluída' && a.status!=='Cancelada'){
            if(fromYMD(a.inicio)<=d && d<=fromYMD(a.fim)) sum += (a.alocacao||100);
          }
        });
        if(sum>cap) overload++;
      });
      if(overload>0){
        let score=overload;
        if((r.tipo||'').toLowerCase()==='externo') score+=5;
        out.push({recurso:r.nome,tipo:r.tipo,score:score,dias:overload});
      }
    });
  }catch(e){ console.error(e); }
  return out;
}

/**
 * Renderiza a tabela de risco no painel "Risco por Recurso". Caso a caixa de seleção
 * "Mostrar apenas recursos com risco" esteja marcada, lista somente recursos com dias de
 * sobrecarga &gt; 0; caso contrário, lista todos os recursos ativos com score calculado (0 para
 * aqueles sem sobrecarga).
 */
function renderRiskScores(){
  const tbl=document.getElementById('tblRisco');
  if(!tbl) return;
  const only=document.getElementById('riskOnlyToggle')?.checked;
  // Construir lista de objetos de risco para todos os recursos ativos
  const items=[];
  const scoreMap = {};
  const computed = computeRiskScores();
  // computed contém apenas recursos com dias>0
  computed.forEach(s=>{
    scoreMap[s.recurso] = {score: s.score, dias: s.dias, tipo: s.tipo};
  });
  resources.filter(r=>r.ativo).forEach(r=>{
    const found = scoreMap[r.nome];
    const dias = found? found.dias: 0;
    let score = dias;
    const tipoLower = (r.tipo||'').toLowerCase();
    if(tipoLower==='externo' && dias>0) score += 5;
    items.push({ recurso: r.nome, tipo: r.tipo, score: score || 0, dias: dias });
  });
  // se apenas com risco: filtrar itens com dias > 0
  const list = only ? items.filter(i=> i.dias > 0) : items;
  // ordenar por score ascendente, depois por recurso
  list.sort((a,b)=>{
    if(a.score !== b.score) return a.score - b.score;
    if(a.dias !== b.dias) return a.dias - b.dias;
    return a.recurso.localeCompare(b.recurso);
  });
  let html='';
  list.forEach(s=>{
    html += `<tr><td>${s.recurso}</td><td>${s.tipo}</td><td>${s.score}</td><td>${s.dias}</td></tr>`;
  });
  tbl.querySelector('tbody').innerHTML = html;
}

// ===== Exportação filtrada (Fase 4) =====
/**
 * Obtém dados filtrados conforme período, lista de recursos e tipo escolhidos na seção de exportação filtrada.
 */
function getFilteredReport(){
  const start=document.getElementById('expStart')?.value||'';
  const end=document.getElementById('expEnd')?.value||'';
  const names=(document.getElementById('expResources')?.value||'').split(',').map(s=>s.trim().toLowerCase()).filter(Boolean);
  const type=(document.getElementById('expType')?.value||'').toLowerCase();
  // Filtra recursos ativos (ignorando deletados) e aplica filtros de nome e tipo
  const resFilt=getActiveResources().filter(r=>{
    if(names.length && !names.some(n=> (r.nome||'').toLowerCase().includes(n))) return false;
    if(type && type!=='' && (r.tipo||'').toLowerCase()!==type) return false;
    return true;
  });
  // Filtra atividades ativas (ignorando deletadas) e dentro do período, e que pertençam a um recurso filtrado
  const actsFilt=getActiveActivities().filter(a=>{
    const rec=resFilt.find(x=>x.id===a.resourceId);
    if(!rec) return false;
    if(start && fromYMD(a.fim)<fromYMD(start)) return false;
    if(end && fromYMD(a.inicio)>fromYMD(end)) return false;
    return true;
  });
  return {resources: resFilt, activities: actsFilt};
}

/**
 * Exporta os dados filtrados em dois CSVs (recursos e atividades). Inclui cabeçalhos para nome,
 * tipo, senioridade, capacidade e dados das atividades correspondentes.
 */
function exportFilteredCSV(){
  const data=getFilteredReport();
  const recRows=data.resources.map(r=>({nome:r.nome, tipo:r.tipo, senioridade:r.senioridade, capacidade:(r.capacidade||100)}));
  const actRows=data.activities.map(a=>({
    titulo:a.titulo,
    recurso:(resources.find(x=>x.id===a.resourceId)?.nome)||'',
    status:a.status,
    inicio:a.inicio,
    fim:a.fim,
    alocacao:(a.alocacao||100)
  }));
  if(recRows.length===0 && actRows.length===0){
    alert('Nenhum dado encontrado para o filtro aplicado.');
    return;
  }
  download('recursos_filtrados.csv', toCSV(recRows, ['nome','tipo','senioridade','capacidade']), 'text/csv;charset=utf-8');
  download('atividades_filtradas.csv', toCSV(actRows, ['titulo','recurso','status','inicio','fim','alocacao']), 'text/csv;charset=utf-8');
  alert('Exportados: recursos_filtrados.csv e atividades_filtradas.csv');
}

/**
 * Gera um relatório em HTML (em nova aba) dos dados filtrados. Calcula indicadores como % de execução
 * (com base nas atividades filtradas), quantidade de recursos ativos e recursos sobrecarregados. Se
 * possível, usa a janela para exibir as tabelas; caso contrário, faz download de um arquivo HTML.
 */
function exportFilteredPDF(){
  const data=getFilteredReport();
  const win=window.open('', '_blank');
  const now=new Date().toLocaleString();
  const totalActs=data.activities.length;
  const concl=data.activities.filter(a=>a.status==='Concluída').length;
  const execPerc=totalActs?Math.round((concl/totalActs)*100):0;
  const riskScores=computeRiskScores();
  // número de recursos sobrecarregados dentro do filtro
  const overCount=data.resources.filter(r=>riskScores.some(s=>s.recurso===r.nome)).length;
  const htmlParts=[];
  htmlParts.push('<html><head><meta charset="utf-8"><title>Relatório Planejador de Recursos (Filtrado)</title>');
  htmlParts.push('<style>body{font-family:sans-serif;padding:20px;} table{border-collapse:collapse;width:100%;margin-top:12px;} th,td{border:1px solid #ccc;padding:4px;} th{background:#f2f2f2;} .kpi span{font-weight:bold;margin-right:4px;}</style>');
  htmlParts.push('</head><body>');
  htmlParts.push('<h2>Relatório Planejador de Recursos (Filtrado)</h2>');
  htmlParts.push('<p>Gerado em: '+now+'</p>');
  htmlParts.push('<h3>KPIs</h3>');
  htmlParts.push('<p><span class="kpi">% Execução:</span> '+execPerc+'%<br>');
  htmlParts.push('<span class="kpi">Recursos Ativos:</span> '+data.resources.length+'<br>');
  htmlParts.push('<span class="kpi">Recursos Sobrecarregados:</span> '+overCount+'</p>');
  if(data.resources.length>0){
    htmlParts.push('<h3>Recursos</h3>');
    htmlParts.push('<table><thead><tr><th>Nome</th><th>Tipo</th><th>Senioridade</th><th>Capacidade%</th></tr></thead><tbody>');
    data.resources.forEach(r=>{
      htmlParts.push('<tr><td>'+r.nome+'</td><td>'+r.tipo+'</td><td>'+r.senioridade+'</td><td>'+ (r.capacidade||100) +'</td></tr>');
    });
    htmlParts.push('</tbody></table>');
  }
  if(data.activities.length>0){
    htmlParts.push('<h3>Atividades</h3>');
    htmlParts.push('<table><thead><tr><th>Título</th><th>Recurso</th><th>Status</th><th>Início</th><th>Fim</th><th>Alocação%</th></tr></thead><tbody>');
    data.activities.forEach(a=>{
      const rName=(resources.find(x=>x.id===a.resourceId)?.nome)||'';
      htmlParts.push('<tr><td>'+a.titulo+'</td><td>'+rName+'</td><td>'+a.status+'</td><td>'+a.inicio+'</td><td>'+a.fim+'</td><td>'+ (a.alocacao||100) +'</td></tr>');
    });
    htmlParts.push('</tbody></table>');
  }
  htmlParts.push('</body></html>');
  const html=htmlParts.join('');
  if(win){
    win.document.write(html);
    win.document.close();
  } else {
    download('relatorio_filtrado.html', html, 'text/html;charset=utf-8');
  }
}

// ===== Exportar PDF =====
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
      renderKPIs();
      doc.text("KPIs:", 14, 40);
      const k1=document.getElementById("kpiExecucao")?.textContent||"";
      const k2=document.getElementById("kpiRecursos")?.textContent||"";
      const k3=document.getElementById("kpiSobrecarga")?.textContent||"";
      doc.text("% Execução: "+k1, 20, 48);
      doc.text("Recursos Ativos: "+k2, 20, 56);
      doc.text("Recursos Sobrecarregados: "+k3, 20, 64);
      doc.text("Recursos:", 14, 78);
      let y=86;
      resources.forEach(r=>{ doc.text("- "+r.nome+" ("+r.tipo+", "+r.senioridade+", Cap "+(r.capacidade||100)+"%)", 20, y); y+=6; if(y>270){doc.addPage(); y=20;} });
      y+=6; doc.text("Atividades:", 14, y); y+=8;
      activities.forEach(a=>{
        const rec=resources.find(r=>r.id===a.resourceId);
        doc.text("- "+a.titulo+" ("+(rec?rec.nome:"—")+") ["+a.status+"] "+a.inicio+" → "+a.fim, 20, y);
        y+=6; if(y>270){doc.addPage(); y=20;}
      });
      if(hasHtml2Canvas){
        try{
          const canvas = await html2canvas(document.getElementById("gantt"));
          const img = canvas.toDataURL("image/png");
          doc.addPage(); doc.text("Visão Gantt",14,20); doc.addImage(img,"PNG",14,30,180,100);
        }catch(e){ /* ignora */ }
      }
      doc.save("planejador_relatorio.pdf");
    } else {
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

// ===== Hook no renderAll para atualizar KPIs e Tags sem quebrar fluxo =====
(function(){
  const _renderAll = renderAll;
  renderAll = function(){
    _renderAll();
    renderKPIs();
    updateTagDatalist();
    // Atualiza o painel de risco sempre que a tela for re-renderizada
    try { renderRiskScores(); } catch(e) {}
  };
})();

// === Registro de eventos para painel de risco e exportações filtradas ===
document.addEventListener('DOMContentLoaded', () => {
  try {
    const riskToggle=document.getElementById('riskOnlyToggle');
    if(riskToggle){ riskToggle.addEventListener('change', () => { renderRiskScores(); }); }
    const btnCsv=document.getElementById('btnExpCSV');
    if(btnCsv){ btnCsv.addEventListener('click', exportFilteredCSV); }
    const btnPdf=document.getElementById('btnExpPDF');
    if(btnPdf){ btnPdf.addEventListener('click', exportFilteredPDF); }
    // Preenche datas padrão com o intervalo de visão atual
    const s=document.getElementById('expStart');
    const e=document.getElementById('expEnd');
    if(s && !s.value) s.value=rangeStart;
    if(e && !e.value) e.value=rangeEnd;
    // Renderiza painel de risco na carga inicial
    try { renderRiskScores(); } catch(err) {}

  } catch(err) {
    console.error(err);
  }
});

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

// ===== BD por Excel/CSV (modelo único) =====
let bdFileHandle = null;

function updateBDStatus(msg){
  const el = document.getElementById('bdStatus');
  if(el) el.textContent = msg||'';
}

function parseHTMLExcelTables(htmlText){
  const doc = new DOMParser().parseFromString(htmlText, 'text/html');
  const tRec = doc.querySelector('#Recursos') || doc.querySelector('table[data-name="Recursos"]') || doc.querySelector('table:nth-of-type(1)');
  const tAtv = doc.querySelector('#Atividades') || doc.querySelector('table[data-name="Atividades"]') || doc.querySelector('table:nth-of-type(2)');
  function tableToObjects(tbl){
    if(!tbl) return [];
    const rows=[...tbl.querySelectorAll('tr')].map(tr=>[...tr.cells].map(td=>td.textContent.trim()));
    if(rows.length===0) return [];
    const headers=rows[0].map(h=>h.trim());
    return rows.slice(1).filter(r=>r.some(v=>v && v.trim().length)).map(r=>{
      const o={}; headers.forEach((h,i)=>o[h]=r[i]??''); return o;
    });
  }
  return { recursos: tableToObjects(tRec), atividades: tableToObjects(tAtv) };
}

// Normaliza datas vindas de arquivos do BD (CSV/HTML) para o formato ISO esperado (YYYY-MM-DD).
// Aceita datas no formato brasileiro (dd/mm/yyyy), no formato ISO (yyyy-mm-dd) ou já normalizadas. Para outros valores retorna a string original.
function normalizeDateField(val){
  if(!val) return '';
  const s = String(val).trim();
  // separa por barra ou hífen
  const parts = s.split(/[\/\-]/);
  if(parts.length === 3){
    // dd/mm/yyyy ou dd-mm-yyyy
    if(parts[0].length === 2 && parts[1].length === 2 && parts[2].length === 4){
      return `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
    }
    // yyyy-mm-dd ou yyyy/mm/dd
    if(parts[0].length === 4 && parts[1].length <= 2 && parts[2].length <= 2){
      return `${parts[0]}-${parts[1].padStart(2,'0')}-${parts[2].padStart(2,'0')}`;
    }
  }
  return s;
}

function coerceResource(r){
  return {
    id: String(r.id||r.ID||r.Id||''),
    nome: r.nome||r.Nome||r.NOME||'',
    tipo: (r.tipo||'').toLowerCase()||'interno',
    senioridade: (r.senioridade||'NA'),
    capacidade: Number(r.capacidade ?? r.Capacidade ?? 100),
    ativo: String(r.ativo||'S').toUpperCase().startsWith('S'),
    // normaliza datas para ISO; se vierem vazias permanecem strings vazias
    inicioAtivo: normalizeDateField(r.inicioAtivo||r.InicioAtivo||r.inicio||''),
    fimAtivo: normalizeDateField(r.fimAtivo||r.FimAtivo||r.fim||''),
    version: Number(r.version||r.versao||0) || 0,
    updatedAt: r.updatedAt ? Number(r.updatedAt) : 0,
    deletedAt: r.deletedAt ? Number(r.deletedAt) : null
  };
}

function coerceActivity(a){
  return {
    id: String(a.id||a.ID||a.Id||''),
    titulo: a.titulo||a.Titulo||a['TÍTULO']||'',
    resourceId: String(a.resourceId||a.RecursoID||a.Recurso||a.resource||''),
    // normaliza inicio e fim para formato ISO esperado
    inicio: normalizeDateField(a.inicio||a.Inicio||a['Início']||''),
    fim: normalizeDateField(a.fim||a.Fim||''),
    status: (a.status||'planejada'),
    alocacao: Number(a.alocacao ?? a.Alocacao ?? 100),
    tags: (a.tags || a.Tags || '').split(',').map(t => t.trim()).filter(Boolean),
    version: Number(a.version||a.versao||0) || 0,
    updatedAt: a.updatedAt ? Number(a.updatedAt) : 0,
    deletedAt: a.deletedAt ? Number(a.deletedAt) : null
  };
}

function parseCSVUnico(text){
  const lines = text.split(/\r?\n/).filter(l=>l.trim().length>0);
  if(lines.length===0) return {recursos:[], atividades:[]};
  const sep = lines[0].includes(';')?';':',';
  const headers = lines[0].split(sep).map(h=>h.trim());
  const rows = lines.slice(1).map(l=>{
    const cols = l.split(sep).map(c=>c.trim().replace(/^"|"$/g,''));
    const o={}; headers.forEach((h,i)=>o[h]=cols[i]||''); return o;
  });
  const recursos = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('recurso')).map(coerceResource);
  const atividades = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('atividade')).map(coerceActivity);
  return {recursos, atividades};
}

function parseHTMLBDTables(htmlText){
  const doc = new DOMParser().parseFromString(htmlText, 'text/html');
  const tRec = doc.querySelector('#Recursos') || doc.querySelector('table[data-name="Recursos"]') || doc.querySelector('table:nth-of-type(1)');
  const tAtv = doc.querySelector('#Atividades') || doc.querySelector('table[data-name="Atividades"]') || doc.querySelector('table:nth-of-type(2)');
  const tHoras = doc.querySelector('#HorasExternos') || doc.querySelector('table[data-name="HorasExternos"]') || doc.querySelector('table:nth-of-type(3)');
  const tHist = doc.querySelector('#HistoricoAtividades') || doc.querySelector('table[data-name="HistoricoAtividades"]');
  const tFeriados = doc.querySelector('#Feriados') || doc.querySelector('table[data-name="Feriados"]');
  function tableToObjects(tbl){
    if(!tbl) return [];
    const rows=[...tbl.querySelectorAll('tr')].map(tr=>[...tr.cells].map(td=>td.textContent.trim()));
    if(rows.length===0) return [];
    const headers=rows[0].map(h=>h.trim());
    return rows.slice(1).filter(r=>r.some(v=>v && v.trim().length)).map(r=>{
      const o={}; headers.forEach((h,i)=>o[h]=r[i]??''); return o;
    });
  }
  const recursos = tableToObjects(tRec);
  const atividades = tableToObjects(tAtv);
  const horasRows = tableToObjects(tHoras);
  const historico = tableToObjects(tHist);
  const feriados = tableToObjects(tFeriados);
  const horas = horasRows.map(h=>{
    const id = h.id || h.ID || h.resourceId || h.RecursoID || h.colaborador || h.Colaborador || '';
    const date = h.date || h.Date || h.data || h.Data || '';
    let minutos = h.minutos || h.Minutos || h.horas || h.Horas || '';
    const parseStr = (s) => {
      s = String(s||'').trim();
      if(!s) return 0;
      const m = s.match(/^(\d+):(\d{2})$/);
      if(m){ return parseInt(m[1],10)*60 + parseInt(m[2],10); }
      if (s.includes('.') || s.includes(',')) {
        const f = parseFloat(s.replace(',', '.'));
        if(!isNaN(f)) return Math.round(f*60);
      }
      const n = parseInt(s,10);
      return isNaN(n)?0:n;
    };
    minutos = parseStr(minutos);
    const tipo = h.tipo || h.Tipo || '';
    const projeto = h.projeto || h.Projeto || '';
    return { id: String(id), date: date, minutos: minutos, tipo: tipo, projeto: projeto };
  });
  const tCfg = doc.querySelector('#HorasExternosCfg') || doc.querySelector('table[data-name="HorasExternosCfg"]') || null;
  let cfg = [];
  if (tCfg) {
    const cfgRows = tableToObjects(tCfg);
    cfg = cfgRows.map(row => {
      const rid = row.id || row.ID || row.resourceId || '';
      const horasDia = row.horasDia || row.horasdia || row.horasDiaMin || row.horas_dia || row.horas_diarias || '';
      const dias = row.dias || row.Dias || row.dia || '';
      const projetos = row.projetos || row.Projetos || row.projeto_cfg || '';
      return { id: String(rid), horasDia: horasDia, dias: dias, projetos: projetos };
    });
  }
  return { recursos, atividades, horas, cfg, historico, feriados };
}

function parseCSVBDUnico(text){
  const lines = text.split(/\r?\n/).filter(l=>l.trim().length>0);
  if(lines.length===0) return {recursos:[], atividades:[], horas:[], historico:[], feriados:[]};
  const sep = lines[0].includes(';')?';':',';
  const headers = lines[0].split(sep).map(h=>h.trim());
  const rows = lines.slice(1).map(l=>{
    const cols = [];
    let cur = '';
    let inQuote = false;
    for(let i=0;i<l.length;i++){
      const ch = l[i];
      if(ch === '"') { inQuote = !inQuote; continue; }
      if(!inQuote && ch === sep){ cols.push(cur.trim()); cur=''; continue; }
      cur += ch;
    }
    cols.push(cur.trim());
    const o={}; headers.forEach((h,i)=>o[h]=cols[i]||''); return o;
  });
  const recursos = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('recurso')).map(coerceResource);
  const atividades = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('atividade')).map(coerceActivity);
  const historico = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('historico'));
  const feriados = rows.filter(r=>String(r.tabela||'').toLowerCase().startsWith('feriado'));
  const horas = rows.filter(r => {
    const tab = String(r.tabela || '').toLowerCase();
    return tab.startsWith('hora') && !tab.startsWith('hora_cfg');
  }).map(r => {
    const id = r.id || r.resourceId || r.colaborador || r.Id || r.ID || '';
    const date = r.date || r.Date || r.data || r.Data || '';
    let minutos = r.minutos || r.Minutos || r.horas || r.Horas || '';
    const parseStr = (s) => {
      s = String(s || '').trim();
      if (!s) return 0;
      const m = s.match(/^(\d+):(\d{2})$/);
      if (m) { return parseInt(m[1], 10) * 60 + parseInt(m[2], 10); }
      if (s.includes('.') || s.includes(',')) {
        const f = parseFloat(s.replace(',', '.'));
        if (!isNaN(f)) return Math.round(f * 60);
      }
      const n = parseInt(s, 10);
      return isNaN(n) ? 0 : n;
    };
    minutos = parseStr(minutos);
    const tipo = r.tipo || r.Tipo || '';
    const projeto = r.projeto || r.Projeto || '';
    return { id: String(id), date: date, minutos: minutos, tipo: tipo, projeto: projeto };
  });
  const cfg = rows.filter(r => String(r.tabela || '').toLowerCase().startsWith('hora_cfg')).map(r => {
    const rid = r.id || r.Id || r.resourceId || '';
    const horasDia = r.horasDia || r.horasdia || r.horasDiaMin || r.horas_dia || '';
    const dias = r.dias || r.Dias || r.dia || '';
    const projetos = r.projetos || r.Projetos || '';
    return { id: String(rid), horasDia: horasDia, dias: dias, projetos: projetos };
  });
  return { recursos, atividades, horas, cfg, historico, feriados };
}

async function saveBD() {
  if (!bdHandle) return;
  try {
    // Aguarda se outra sessão estiver salvando recentemente e previne concorrência
    await acquireBDLockIfBusy();
    // Checagem de versão: detecta alterações feitas por outra sessão desde o último salvamento
    try {
      const fchk = await bdHandle.getFile();
      const currentLm = fchk.lastModified;
      if (typeof window !== 'undefined' && currentLm > (window.__bdLastWrite || 0)) {
        /*
         * Em vez de abortar o salvamento, atualizamos nosso marcador de última escrita
         * para o valor atual. Isso evita perder alterações locais: se outra sessão
         * salvou recentemente, já esperamos pelo tempo definido em acquireBDLockIfBusy().
         * Ao prosseguir, escrevemos o nosso estado atual sobre a versão antiga.
         */
        window.__bdLastWrite = currentLm;
      }
    } catch (e) {
      // falha ao checar modificação; prosseguir com salvamento
    }
    let content = '';
    let mime = '';
    let horasList = [];
    let cfgList = [];
    let feriadosList = [];
    try {
      if (typeof window.getHorasExternosData === 'function') {
        const out = window.getHorasExternosData();
        if (Array.isArray(out)) horasList = out;
      }
    } catch(e){}
    try {
      if (typeof window.getHorasExternosConfig === 'function') {
        const outCfg = window.getHorasExternosConfig();
        if (Array.isArray(outCfg)) cfgList = outCfg;
      }
    } catch(e){}
    try {
        if (typeof window.getFeriados === 'function') {
            const outFeriados = window.getFeriados();
            if(Array.isArray(outFeriados)) feriadosList = outFeriados;
        }
    } catch(e) {}
    if (bdFileExt === 'csv') {
      // Cabeçalho CSV incluindo campos de versionamento e exclusão (version, updatedAt, deletedAt)
      const header = [
        'tabela','id','nome','tipo','senioridade','capacidade','ativo','inicioAtivo','fimAtivo',
        'version','updatedAt','deletedAt',
        'titulo','resourceId','inicio','fim','status','alocacao','tags',
        'date','minutos','tipoHora','projeto','horasDia','dias','projetos',
        'activityId','timestamp','oldInicio','oldFim','newInicio','newFim','justificativa','user','legend'
      ];
      const rows = [];
      // Gera linhas de recursos com campos extras
      resources.forEach(r => {
        rows.push({
          tabela:'recurso',
          id:r.id,
          nome:r.nome,
          tipo:r.tipo,
          senioridade:r.senioridade,
          capacidade:r.capacidade,
          ativo:r.ativo?'S':'N',
          inicioAtivo:r.inicioAtivo||'',
          fimAtivo:r.fimAtivo||'',
          version:r.version || 0,
          updatedAt:r.updatedAt || 0,
          deletedAt:r.deletedAt || '',
          titulo:'', resourceId:'', inicio:'', fim:'', status:'', alocacao:'', tags:'',
          date:'', minutos:'', tipoHora:'', projeto:'', horasDia:'', dias:'', projetos:'',
          activityId:'', timestamp:'', oldInicio:'', oldFim:'', newInicio:'', newFim:'', justificativa:'', user:'', legend:''
        });
      });
      // Gera linhas de atividades com campos extras
      activities.forEach(a => {
        rows.push({
          tabela:'atividade',
          id:a.id,
          nome:'', tipo:'', senioridade:'', capacidade:'', ativo:'', inicioAtivo:'', fimAtivo:'',
          version:a.version || 0,
          updatedAt:a.updatedAt || 0,
          deletedAt:a.deletedAt || '',
          titulo:a.titulo,
          resourceId:a.resourceId,
          inicio:a.inicio,
          fim:a.fim,
          status:a.status,
          alocacao:a.alocacao,
          tags:(a.tags || []).join(', '),
          date:'', minutos:'', tipoHora:'', projeto:'', horasDia:'', dias:'', projetos:'',
          activityId:'', timestamp:'', oldInicio:'', oldFim:'', newInicio:'', newFim:'', justificativa:'', user:'', legend:''
        });
      });
      horasList.forEach(h => {
        rows.push({
          tabela:'hora_externo', id:h.id, date:h.date || '', minutos:h.minutos, tipoHora:h.tipo || '', projeto:h.projeto || ''
        });
      });
      cfgList.forEach(cfg => {
        rows.push({
          tabela:'hora_cfg', id:cfg.id, horasDia: cfg.horasDia || '', dias: cfg.dias || '', projetos: cfg.projetos || ''
        });
      });
      Object.keys(trails).forEach(activityId => {
        (trails[activityId] || []).forEach(entry => {
          rows.push({
            tabela: 'historico',
            activityId: activityId,
            timestamp: entry.ts,
            oldInicio: entry.oldInicio,
            oldFim: entry.oldFim,
            newInicio: entry.newInicio,
            newFim: entry.newFim,
            justificativa: entry.justificativa,
            user: entry.user
          });
        });
      });
      feriadosList.forEach(f => {
          rows.push({
              tabela: 'feriado',
              date: f.date,
              legend: f.legend
          });
      });
      const csvRows = [];
      csvRows.push(header.join(','));
      rows.forEach(row => {
        const vals = header.map(h => {
          let v = row[h] || '';
          const needsQuote = String(v).includes(',') || String(v).includes(';') || String(v).includes('"');
          v = String(v).replace(/"/g, '""');
          return needsQuote ? '"'+v+'"' : v;
        });
        csvRows.push(vals.join(','));
      });
      content = csvRows.join('\n');
      mime = 'text/csv;charset=utf-8';
    } else {
      function tableHTML(title, headers, rows) {
        const thead = headers.map(h => `<th>${h}</th>`).join('');
        const tbody = rows.map(r => `<tr>${headers.map(h => `<td>${r[h] ?? ''}</td>`).join('')}</tr>`).join('');
        return `<h3>${title}</h3><table id='${title}' data-name='${title}' border='1'><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table>`;
      }
      // Tabelas HTML: adiciona version, updatedAt e deletedAt para recursos e atividades
      const headersRec = ['id','nome','tipo','senioridade','capacidade','ativo','inicioAtivo','fimAtivo','version','updatedAt','deletedAt'];
      const recRows = resources.map(r => ({
        id:r.id,
        nome:r.nome,
        tipo:r.tipo,
        senioridade:r.senioridade,
        capacidade:r.capacidade,
        ativo:r.ativo?'S':'N',
        inicioAtivo:r.inicioAtivo||'',
        fimAtivo:r.fimAtivo||'',
        version:r.version || 0,
        updatedAt:r.updatedAt || 0,
        deletedAt:r.deletedAt || ''
      }));
      const headersAtv = ['id','titulo','resourceId','inicio','fim','status','alocacao','tags','version','updatedAt','deletedAt'];
      const atvRows = activities.map(a => ({
        id:a.id,
        titulo:a.titulo,
        resourceId:a.resourceId,
        inicio:a.inicio,
        fim:a.fim,
        status:a.status,
        alocacao:a.alocacao,
        tags:(a.tags || []).join(', '),
        version:a.version || 0,
        updatedAt:a.updatedAt || 0,
        deletedAt:a.deletedAt || ''
      }));
      const headersHoras = ['id','date','minutos','tipo','projeto'];
      const horasRows = horasList.map(h => ({ id:h.id, date:h.date || '', minutos:h.minutos, tipo:h.tipo || '', projeto:h.projeto || '' }));
      const headersCfg = ['id','horasDia','dias','projetos'];
      const cfgRows = cfgList.map(cfg => ({ id: cfg.id, horasDia: cfg.horasDia || '', dias: cfg.dias || '', projetos: cfg.projetos || '' }));
      const headersHist = ['activityId', 'timestamp', 'oldInicio', 'oldFim', 'newInicio', 'newFim', 'justificativa', 'user'];
      const histRows = [];
      Object.keys(trails).forEach(activityId => {
          (trails[activityId] || []).forEach(entry => {
              histRows.push({
                  activityId: activityId,
                  timestamp: entry.ts,
                  oldInicio: entry.oldInicio,
                  oldFim: entry.oldFim,
                  newInicio: entry.newInicio,
                  newFim: entry.newFim,
                  justificativa: entry.justificativa,
                  user: entry.user
              });
          });
      });
      const headersFeriados = ['date', 'legend'];
      const feriadosRows = feriadosList.map(f => ({ date: f.date, legend: f.legend || '' }));
      
      content = `<!doctype html><html><head><meta charset='utf-8'><title>BD</title></head><body>`+
        tableHTML('Recursos', headersRec, recRows) +
        tableHTML('Atividades', headersAtv, atvRows) +
        tableHTML('HorasExternos', headersHoras, horasRows) +
        tableHTML('HorasExternosCfg', headersCfg, cfgRows) +
        tableHTML('HistoricoAtividades', headersHist, histRows) +
        tableHTML('Feriados', headersFeriados, feriadosRows) +
        `</body></html>`;
      mime = 'text/html;charset=utf-8';
    }
    const writable = await bdHandle.createWritable();
    await writable.write(new Blob([content], { type: mime }));
    await writable.close();
    updateBDStatus('Salvo em ' + (bdFileName || 'BD'));
    // após salvar com sucesso, atualiza a última modificação registrada
    if (typeof window !== 'undefined') {
      window.__bdLastWrite = Date.now();
    }
  } catch (e) {
    console.error('Erro ao salvar BD:', e);
    updateBDStatus('Erro ao salvar BD');
  }
}

function saveBDDebounced() {
  if (!bdHandle) return;
  clearTimeout(_saveBDTimer);
  _saveBDTimer = setTimeout(() => { saveBD(); }, 1000);
}

if (typeof window !== 'undefined') {
  try {
    window.onHorasExternosChange = () => {
      saveBDDebounced();
    };
  } catch(e){}
}

const fileBD = document.getElementById('fileBD');
if(fileBD){
  fileBD.onchange = async (ev)=>{
    const f = ev.target.files && ev.target.files[0];
    if(!f) return;
    try{
      const ext = f.name.toLowerCase().split('.').pop();
      const text = await f.text();
      let parsed;
      if(ext==='csv'){
        parsed = parseCSVBDUnico(text);
        resources = (parsed.recursos || []).map(coerceResource);
        activities = (parsed.atividades || []).map(coerceActivity);
        if(parsed.horas && typeof window.setHorasExternosData === 'function') window.setHorasExternosData(parsed.horas);
        if(parsed.cfg && typeof window.setHorasExternosConfig === 'function') window.setHorasExternosConfig(parsed.cfg);
        if(parsed.feriados && typeof window.setFeriados === 'function') window.setFeriados(parsed.feriados);
      } else {
        parsed = parseHTMLBDTables(text);
        resources = (parsed.recursos || []).map(coerceResource);
        activities = (parsed.atividades || []).map(coerceActivity);
        if(parsed.horas && typeof window.setHorasExternosData === 'function') window.setHorasExternosData(parsed.horas);
        if(parsed.cfg && typeof window.setHorasExternosConfig === 'function') window.setHorasExternosConfig(parsed.cfg);
        if(parsed.feriados && typeof window.setFeriados === 'function') window.setFeriados(parsed.feriados);
      }
      const newTrails = {};
      (parsed.historico || []).forEach(h => {
        const id = h.activityId;
        if (!id) return;
        if (!newTrails[id]) newTrails[id] = [];
        newTrails[id].push({
          ts: h.timestamp,
          oldInicio: h.oldInicio,
          oldFim: h.oldFim,
          newInicio: h.newInicio,
          newFim: h.newFim,
          justificativa: h.justificativa,
          user: h.user
        });
      });
      trails = newTrails;
      saveLS(LS.res, resources);
      saveLS(LS.act, activities);
      saveLS(LS.trail, trails);
      renderAll();
      updateBDStatus('BD carregado: '+ f.name);
    } catch(e){ alert('Erro ao ler arquivo BD: '+ e.message); }
  };
}

const btnSelectDirInBD = document.getElementById('btnSelectDirInBD');
if(btnSelectDirInBD){
  btnSelectDirInBD.onclick = async ()=>{
    try{
      const h = await window.showDirectoryPicker();
      if(!h) return;
      dirHandle = h;
      await idbSet(FSA_DB, FSA_STORE, 'dir', h);
      updateBDStatus('Pasta selecionada ✓ — Salvo');
      try{ await verifyPerm(dirHandle); }catch(e){}
    }catch(e){
      alert('Não foi possível selecionar a pasta.\nDica: abra pelo Chrome/Edge via http(s):// em vez de file://');
      console.warn(e);
    }
  };
}

const btnPickBDFile = document.getElementById('btnPickBDFile');
if(btnPickBDFile){
  btnPickBDFile.onclick = async () => {
    if (!('showOpenFilePicker' in window)) {
      alert('Seu navegador não suporta a abertura de arquivos com permissão de gravação. Use o Chrome/Edge via http(s)://');
      return;
    }
    try {
      const [handle] = await window.showOpenFilePicker({
        multiple: false,
        types: [
          {
            description: 'Arquivos de Banco de Dados',
            accept: {
              'application/vnd.ms-excel': ['.xls', '.html'],
              'text/csv': ['.csv'],
              'text/plain': ['.csv']
            }
          }
        ]
      });
      if (!handle) return;
      bdHandle = handle;
      const file = await handle.getFile();
      bdFileName = file.name || '';
      bdFileExt = (bdFileName.split('.').pop() || '').toLowerCase();
      const text = await file.text();
      let parsed;
      if (bdFileExt === 'csv') {
        parsed = parseCSVBDUnico(text);
      } else {
        parsed = parseHTMLBDTables(text);
      }
      resources = (parsed.recursos || []).map(coerceResource);
      activities = (parsed.atividades || []).map(coerceActivity);
      if (parsed.horas && typeof window.setHorasExternosData === 'function') {
        window.setHorasExternosData(parsed.horas);
      }
      if (parsed.cfg && typeof window.setHorasExternosConfig === 'function') {
        window.setHorasExternosConfig(parsed.cfg);
      }
      if (parsed.feriados && typeof window.setFeriados === 'function') {
        window.setFeriados(parsed.feriados);
      }
      const newTrails = {};
      (parsed.historico || []).forEach(h => {
        const id = h.activityId;
        if (!id) return;
        if (!newTrails[id]) newTrails[id] = [];
        newTrails[id].push({
          ts: h.timestamp,
          oldInicio: h.oldInicio,
          oldFim: h.oldFim,
          newInicio: h.newInicio,
          newFim: h.newFim,
          justificativa: h.justificativa,
          user: h.user
        });
      });
      trails = newTrails;
      // Ao apontar um novo BD, reinicia o log de eventos e o snapshot para evitar reaplicar
      // eventos antigos (anteriores à escolha do BD) que poderiam trazer dados "fantasmas".
      resetEventLogAndSnapshot();
      saveLS(LS.res, resources);
      saveLS(LS.act, activities);
      saveLS(LS.trail, trails);
      renderAll();
      updateBDStatus('BD carregado e pronto: ' + bdFileName);
      // registra o lastModified atual como a última modificação conhecida do arquivo
      try {
        const lm = file.lastModified;
        if (typeof window !== 'undefined') {
          window.__bdLastWrite = lm;
        }
      } catch(e){}
      startBDWatcher();
      updateDBStatusBanner(true);
    } catch (e) {
      if (e && e.name !== 'AbortError') {
        alert('Erro ao abrir arquivo BD: ' + e.message);
      }
    }
  };
}

// Permitir ao usuário definir um arquivo de BD como padrão para carregamento automático
const btnSetDefaultBD = document.getElementById('btnSetDefaultBD');
if(btnSetDefaultBD){
  btnSetDefaultBD.onclick = async () => {
    if (!('showOpenFilePicker' in window)) {
      alert('Seu navegador não suporta a abertura de arquivos com permissão de gravação. Use o Chrome/Edge via http(s)://');
      return;
    }
    try {
      const [handle] = await window.showOpenFilePicker({
        multiple: false,
        types: [
          {
            description: 'Arquivos de Banco de Dados',
            accept: {
              'application/vnd.ms-excel': ['.xls', '.html'],
              'text/csv': ['.csv'],
              'text/plain': ['.csv']
            }
          }
        ]
      });
      if(!handle) return;
      bdHandle = handle;
      // Persistir handle como BD padrão
      try { await idbSet(FSA_DB, FSA_STORE, 'bd', handle); } catch(e) { console.warn('Erro ao salvar BD padrão', e); }
      const file = await handle.getFile();
      bdFileName = file.name || '';
      bdFileExt = (bdFileName.split('.').pop() || '').toLowerCase();
      const text = await file.text();
      let parsed;
      if(bdFileExt === 'csv') parsed = parseCSVBDUnico(text);
      else parsed = parseHTMLBDTables(text);
      resources = (parsed.recursos || []).map(coerceResource);
      activities = (parsed.atividades || []).map(coerceActivity);
      if (parsed.horas && typeof window.setHorasExternosData === 'function') {
        window.setHorasExternosData(parsed.horas);
      }
      if (parsed.cfg && typeof window.setHorasExternosConfig === 'function') {
        window.setHorasExternosConfig(parsed.cfg);
      }
      if (parsed.feriados && typeof window.setFeriados === 'function') {
        window.setFeriados(parsed.feriados);
      }
      const newTrails = {};
      (parsed.historico || []).forEach(h => {
        const id = h.activityId;
        if (!id) return;
        if (!newTrails[id]) newTrails[id] = [];
        newTrails[id].push({
          ts: h.timestamp,
          oldInicio: h.oldInicio,
          oldFim: h.oldFim,
          newInicio: h.newInicio,
          newFim: h.newFim,
          justificativa: h.justificativa,
          user: h.user
        });
      });
      trails = newTrails;
      // Ao definir um BD como padrão, reinicia o log de eventos e o snapshot para
      // evitar reaplicar eventos antigos e trazer dados não pertencentes ao BD.
      resetEventLogAndSnapshot();
      saveLS(LS.res, resources);
      saveLS(LS.act, activities);
      saveLS(LS.trail, trails);
      renderAll();
      updateBDStatus('BD padrão definido: ' + bdFileName);
      // registra o lastModified atual como a última modificação conhecida do arquivo
      try {
        const lm = file.lastModified;
        if (typeof window !== 'undefined') {
          window.__bdLastWrite = lm;
        }
      } catch(e){}
      startBDWatcher();
      updateDBStatusBanner(true);
    } catch(e) {
      if(e && e.name !== 'AbortError'){
        alert('Erro ao definir BD padrão: ' + e.message);
      }
    }
  };
}

async function ensureDirOrAsk(){
  if(dirHandle) return true;
  try{
    const h = await window.showDirectoryPicker();
    if(!h) return false;
    dirHandle = h;
    await idbSet(FSA_DB, FSA_STORE, 'dir', h);
    updateBDStatus('Pasta selecionada ✓ — Salvo');
    return true;
  }catch(e){
    alert('Defina a pasta de dados para salvar o modelo.');
    return false;
  }
}

const btnExportModeloXLS = document.getElementById('btnExportModeloXLS');
if(btnExportModeloXLS){
  btnExportModeloXLS.onclick = () => {
    const headersRec = ['id','nome','tipo','senioridade','capacidade','ativo','inicioAtivo','fimAtivo'];
    const headersAtv = ['id','titulo','resourceId','inicio','fim','status','alocacao', 'tags'];
    const headersHoras = ['id','date','minutos','tipo','projeto'];
    const exampleRec = [{id:'R1',nome:'Recurso Exemplo',tipo:'interno',senioridade:'Pl',capacidade:100,ativo:'S',inicioAtivo:'2025-01-01',fimAtivo:''}];
    const exampleAtv = [{id:'A1',titulo:'Atividade Exemplo',resourceId:'R1',inicio:'2025-01-10',fim:'2025-01-20',status:'planejada',alocacao:100, tags: 'SAP, Manutenção'}];
    const exampleHoras = [{id:'R1',date:'2025-01-15',minutos:480,tipo:'trabalho',projeto:'Alca Analitico'}];
    const headersCfg = ['id','horasDia','dias','projetos'];
    const exampleCfg = [{id:'R1',horasDia:'08:00',dias:'seg,ter,qua,qui,sex',projetos:'Alca Analitico:300:00'}];
    const headersHist = ['activityId','timestamp','oldInicio','oldFim','newInicio','newFim','justificativa','user'];
    const exampleHist = [{activityId:'A1', timestamp:new Date().toISOString(), oldInicio:'2025-01-10', oldFim:'2025-01-20', newInicio:'2025-01-11', newFim:'2025-01-22', justificativa:'Ajuste de escopo', user:'usuário'}];
    const headersFeriados = ['date', 'legend'];
    const exampleFeriados = [{date: '2025-12-25', legend: 'Natal'}];
    function table(title, headers, rows){
      const thead = headers.map(h=>`<th>${h}</th>`).join('');
      const tbody = rows.map(r=>`<tr>${headers.map(h=>`<td>${(r[h]??'')}</td>`).join('')}</tr>`).join('');
      return `<h3>${title}</h3><table id="${title}" data-name="${title}" border='1'><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table>`;
    }
    const html = `<!doctype html><html><head><meta charset='utf-8'><title>Modelo BD</title></head><body>`+
      table('Recursos', headersRec, exampleRec) +
      table('Atividades', headersAtv, exampleAtv) +
      table('HorasExternos', headersHoras, exampleHoras) +
      table('HorasExternosCfg', headersCfg, exampleCfg) +
      table('HistoricoAtividades', headersHist, exampleHist) +
      table('Feriados', headersFeriados, exampleFeriados) +
      `</body></html>`;
    download('modelo_bd.xls', html, 'application/vnd.ms-excel');
    alert('Modelo de BD (Excel) gerado: modelo_bd.xls');
  };
}

const btnExportModeloCSV = document.getElementById('btnExportModeloCSV');
if(btnExportModeloCSV){
  btnExportModeloCSV.onclick = () => {
    const headers = ['tabela','id','nome','tipo','senioridade','capacidade','ativo','inicioAtivo','fimAtivo','titulo','resourceId','inicio','fim','status','alocacao','tags','date','minutos','tipoHora','projeto','horasDia','dias','projetos', 'activityId','timestamp','oldInicio','oldFim','newInicio','newFim','justificativa','user', 'legend'];
    const sample = [
      {tabela:'recurso',id:'R1',nome:'Recurso Exemplo',tipo:'interno',senioridade:'Pl',capacidade:100,ativo:'S',inicioAtivo:'2025-01-01',fimAtivo:''},
      {tabela:'atividade',id:'A1',titulo:'Atividade Exemplo',resourceId:'R1',inicio:'2025-01-10',fim:'2025-01-20',status:'planejada',alocacao:100, tags: 'SAP, Manutenção'},
      {tabela:'hora_externo',id:'R1',date:'2025-01-15',minutos:480,tipoHora:'trabalho',projeto:'Alca Analitico'},
      {tabela:'hora_cfg',id:'R1',horasDia:'08:00',dias:'seg,ter,qua,qui,sex',projetos:'Alca Analitico:300:00'},
      {tabela:'historico', activityId:'A1', timestamp:new Date().toISOString(), oldInicio:'2025-01-10', oldFim:'2025-01-20', newInicio:'2025-01-11', newFim:'2025-01-22', justificativa:'Ajuste de escopo', user:'usuário'},
      {tabela:'feriado', date:'2025-12-25', legend:'Natal'}
    ];
    const rows = [headers.join(',')];
    sample.forEach(obj => {
      const line = headers.map(h => {
        const v = obj[h] ?? '';
        const cleanV = String(v).replace(/"/g, '""');
        return '"' + cleanV + '"';
      }).join(',');
      rows.push(line);
    });
    const csv = rows.join('\n');
    download('modelo_bd.csv', csv, 'text/csv;charset=utf-8');
    alert('Modelo de BD (CSV único) gerado: modelo_bd.csv');
  };
}

(() => {
  if(dirHandle){ updateBDStatus('Pasta selecionada ✓ — Salvo'); }
})();

// ===== Importação em massa de dados =====
const fileImportData = document.getElementById('fileImportData');
if (fileImportData) {
  fileImportData.addEventListener('change', async (ev) => {
    const f = ev.target.files && ev.target.files[0];
    if (!f) return;
    try {
      const text = await f.text();
      await importCSVData(text);
      // Limpa o valor para permitir importar o mesmo arquivo novamente se necessário
      ev.target.value = '';
    } catch (e) {
      alert('Erro ao importar dados: ' + (e && e.message ? e.message : e));
      console.error(e);
    }
  });
}

const btnExportImportTemplate = document.getElementById('btnExportImportTemplate');
if (btnExportImportTemplate) {
  btnExportImportTemplate.onclick = () => {
    // Cabeçalho e dados de exemplo para o modelo de importação. Utiliza ponto e vírgula
    // como separador para ser compatível com planilhas em português (Excel)
    const headers = ['tipo','nome','tipoRecurso','senioridade','capacidade','ativo','inicioAtivo','fimAtivo','titulo','recursoNome','inicio','fim','status','alocacao','tags'];
    const samples = [
      {
        tipo: 'recurso',
        nome: 'Recurso Exemplo',
        tipoRecurso: 'Interno',
        senioridade: 'Pl',
        capacidade: 100,
        ativo: 'S',
        inicioAtivo: '2025-01-01',
        fimAtivo: ''
      },
      {
        tipo: 'atividade',
        titulo: 'Atividade Exemplo',
        recursoNome: 'Recurso Exemplo',
        inicio: '2025-01-10',
        fim: '2025-01-20',
        status: 'planejada',
        alocacao: 100,
        tags: 'ProjetoX, API'
      }
    ];
    const sep = ';';
    const rows = [headers.join(sep)];
    samples.forEach(obj => {
      const line = headers.map(h => {
        let v = obj[h] !== undefined && obj[h] !== null ? obj[h] : '';
        let s = String(v);
        // Escapa aspas
        s = s.replace(/"/g, '""');
        // Se contiver separador, vírgula ou quebra de linha, envolve em aspas
        if (s.includes(sep) || s.includes(',') || s.includes('\n')) {
          s = '"' + s + '"';
        }
        return s;
      }).join(sep);
      rows.push(line);
    });
    const csv = rows.join('\n');
    download('modelo_importacao.csv', csv, 'text/csv;charset=utf-8');
    alert('Modelo de importação gerado: modelo_importacao.csv');
  };
}

/**
 * Importa dados de um CSV para recursos e atividades. Cada linha deve conter
 * um campo "tipo" definindo se é "recurso" ou "atividade" (pode também
 * utilizar "tabela" com "recurso" ou "atividade"). Para recursos,
 * campos como nome, tipoRecurso, senioridade, capacidade, ativo,
 * inicioAtivo e fimAtivo são considerados. Para atividades, campos
 * titulo, recursoNome, inicio, fim, status, alocacao e tags são
 * utilizados. Recursos são mesclados pelo nome (case-insensitive);
 * atividades são mescladas pela combinação (titulo normalizado,
 * resourceId, inicio e fim). Novos IDs são gerados via uuid().
 * Para cada criação ou atualização, um evento é registrado. Ao final,
 * as listas resources e activities são salvas e a UI atualizada.
 * @param {string} csvText Conteúdo do arquivo CSV
 */
async function importCSVData(csvText) {
  try {
    const lines = csvText.split(/\r?\n/).filter(l => l.trim().length > 0);
    if (lines.length === 0) {
      alert('Arquivo vazio');
      return;
    }
    // Detecta separador (ponto e vírgula ou vírgula)
    const sep = lines[0].includes(';') ? ';' : ',';
    // Função para parsear uma linha com suporte a aspas
    function parseLine(line) {
      const cols = [];
      let cur = '';
      let inQuote = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          inQuote = !inQuote;
          continue;
        }
        if (!inQuote && ch === sep) {
          cols.push(cur);
          cur = '';
          continue;
        }
        cur += ch;
      }
      cols.push(cur);
      return cols.map(c => c.trim());
    }
    const headers = parseLine(lines[0]);
    let createdRes = 0, updatedRes = 0, createdAct = 0, updatedAct = 0, errors = 0;

    // Normaliza datas no formato ISO (YYYY-MM-DD). Aceita entradas no formato
    // brasileiro (dd/mm/yyyy) ou ISO (yyyy-mm-dd) e retorna sempre yyyy-mm-dd.
    // Para outros formatos ou entradas vazias, retorna a string original.
    function normalizeDateStr(d) {
      if (!d) return '';
      // remove aspas e espaços
      const s = String(d).trim().replace(/\"/g, '');
      // detecta separador ( '/' ou '-' )
      const parts = s.split(/[\/\-]/);
      if (parts.length === 3) {
        // dd/mm/yyyy -> yyyy-mm-dd
        if (parts[0].length === 2 && parts[1].length === 2 && parts[2].length === 4) {
          return `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
        }
        // yyyy-mm-dd ou yyyy/mm/dd -> yyyy-mm-dd
        if (parts[0].length === 4) {
          return `${parts[0]}-${parts[1].padStart(2,'0')}-${parts[2].padStart(2,'0')}`;
        }
      }
      return s;
    }
    for (let i = 1; i < lines.length; i++) {
      const rowLine = lines[i];
      if (!rowLine.trim()) continue;
      const cols = parseLine(rowLine);
      const row = {};
      headers.forEach((h, idx) => {
        row[h] = cols[idx] !== undefined ? cols[idx] : '';
      });
      let rectype = '';
      if (row.tipo) rectype = String(row.tipo).toLowerCase();
      if (!rectype && row.tabela) rectype = String(row.tabela).toLowerCase();
      if (rectype.startsWith('recurso') || rectype === 'recurso' || rectype === 'resource') {
        // Importar recurso
        const name = row.nome || row.nomeRecurso || row.recursoNome || '';
        if (!name) {
          errors++;
          continue;
        }
        let existing = (resources || []).find(r => normalizeName(r.nome) === normalizeName(name));
        if (existing) {
          let changed = false;
          if (row.tipoRecurso) {
            existing.tipo = row.tipoRecurso;
            changed = true;
          }
          if (row.senioridade) {
            existing.senioridade = row.senioridade;
            changed = true;
          }
          if (row.capacidade) {
            const cap = parseFloat(row.capacidade.replace(',', '.'));
            if (!isNaN(cap)) {
              existing.capacidade = cap;
              changed = true;
            }
          }
          if (row.ativo) {
            const val = String(row.ativo).toLowerCase();
            existing.ativo = ['s','sim','y','yes','1','true'].includes(val);
            changed = true;
          }
          if (row.inicioAtivo) {
            existing.inicioAtivo = normalizeDateStr(row.inicioAtivo);
            changed = true;
          }
          if (row.fimAtivo) {
            existing.fimAtivo = normalizeDateStr(row.fimAtivo);
            changed = true;
          }
          if (changed) {
            existing.version = (existing.version || 0) + 1;
            existing.updatedAt = Date.now();
            recordEvent('resource', 'update', existing.id, { ...existing });
            updatedRes++;
          }
        } else {
          const newId = uuid();
          const nr = {
            id: newId,
            nome: name,
            tipo: row.tipoRecurso || 'Interno',
            senioridade: row.senioridade || '',
            capacidade: row.capacidade ? parseFloat(String(row.capacidade).replace(',', '.')) || 0 : 0,
            ativo: row.ativo ? ['s','sim','y','yes','1','true'].includes(String(row.ativo).toLowerCase()) : true,
            inicioAtivo: normalizeDateStr(row.inicioAtivo || ''),
            fimAtivo: normalizeDateStr(row.fimAtivo || ''),
            version: 1,
            updatedAt: Date.now(),
            deletedAt: null
          };
          (resources = resources || []).push(nr);
          recordEvent('resource', 'create', nr.id, { ...nr });
          createdRes++;
        }
      } else if (rectype.startsWith('atividade') || rectype === 'atividade' || rectype === 'activity') {
        // Importar atividade
        const title = row.titulo || row.title || '';
        const rName = row.recursoNome || row.resourceName || row.nomeRecurso || row.nome || '';
        if (!title || !rName) {
          errors++;
          continue;
        }
        // Garante que o recurso exista
        let res = (resources || []).find(r => normalizeName(r.nome) === normalizeName(rName));
        let resCreated = false;
        if (!res) {
          // Cria novo recurso com valores mínimos
          const newResId = uuid();
          const nr2 = {
            id: newResId,
            nome: rName,
            tipo: 'Interno',
            senioridade: '',
            capacidade: 0,
            ativo: true,
            inicioAtivo: '',
            fimAtivo: '',
            version: 1,
            updatedAt: Date.now(),
            deletedAt: null
          };
          (resources = resources || []).push(nr2);
          recordEvent('resource', 'create', nr2.id, { ...nr2 });
          createdRes++;
          res = nr2;
          resCreated = true;
        }
        const start = normalizeDateStr(row.inicio || row.start || '');
        const end = normalizeDateStr(row.fim || row.end || '');
        // Determina o status da atividade. Se não informado, assume 'Planejada'.
        let statusRaw = row.status || row.situacao || '';
        statusRaw = String(statusRaw).trim();
        let status = 'Planejada';
        if (statusRaw) {
          // Procura um status igualando case e acentos. Usa a lista STATUS para validar.
          const foundStatus = STATUS.find(s => s.toLowerCase() === statusRaw.toLowerCase());
          if (foundStatus) {
            status = foundStatus;
          }
        }
        let alloc = 0;
        if (row.alocacao) {
          const num = parseFloat(String(row.alocacao).replace(',', '.'));
          if (!isNaN(num)) alloc = num;
        }
        const tagsStr = row.tags || '';
        const tagsArr = tagsStr ? tagsStr.split(/[,;]/).map(t => t.trim()).filter(Boolean) : [];
        // Tenta achar atividade existente por chave composta
        let existingAct = (activities || []).find(a => normalizeName(a.titulo) === normalizeName(title) && a.resourceId === res.id && a.inicio === start && a.fim === end);
        if (existingAct) {
          let changed = false;
          if (status) {
            existingAct.status = status;
            changed = true;
          }
          if (alloc) {
            existingAct.alocacao = alloc;
            changed = true;
          }
          if (tagsArr && tagsArr.length > 0) {
            existingAct.tags = tagsArr;
            changed = true;
          }
          if (changed) {
            existingAct.version = (existingAct.version || 0) + 1;
            existingAct.updatedAt = Date.now();
            recordEvent('activity', 'update', existingAct.id, { ...existingAct });
            updatedAct++;
          }
        } else {
          const newActId = uuid();
          const na = {
            id: newActId,
            titulo: title,
            resourceId: res.id,
            inicio: start,
            fim: end,
            status: status,
            alocacao: alloc,
            tags: tagsArr,
            version: 1,
            updatedAt: Date.now(),
            deletedAt: null
          };
          (activities = activities || []).push(na);
          recordEvent('activity', 'create', na.id, { ...na });
          createdAct++;
        }
      } else {
        // Tipo não reconhecido
        errors++;
      }
    }
    // Persiste e atualiza UI
    saveLS(LS.res, resources);
    saveLS(LS.act, activities);
    renderAll();
    // Salva no BD, se houver handle definido
    saveBDDebounced();
    alert(`Importação concluída:\n${createdRes} recursos criados, ${updatedRes} atualizados;\n${createdAct} atividades criadas, ${updatedAct} atualizadas;\n${errors} linhas ignoradas ou inválidas.`);
  } catch (err) {
    console.error('Erro ao processar importação CSV', err);
    alert('Erro ao processar importação: ' + (err && err.message ? err.message : err));
  }
}