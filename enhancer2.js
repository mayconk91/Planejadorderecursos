/* enhancer2.js - Gestão de Horas (Externos) + Toolbar FSA (robusta) */
(() => {
  const DB = 'rv-enhancer-v2';
  // Persisted state (threshold and per-recurso/project data). Additional UI filter
  // values live outside of persisted state to avoid polluting storage.
  const state = { thresholdMin: 10*60, externos: {}, folder: null };
  // UI-only filters. These are not persisted but retained across renders.
  state.selResId = '';
  state.selProjName = '';
  state.filterText = '';

  // Utils
  const q = (sel, root=document) => root.querySelector(sel);
  const qa = (sel, root=document) => Array.from(root.querySelectorAll(sel));
  const pad2 = n => String(n).padStart(2,'0');
  // Parse a string into minutes. Accepts HH:MM (unlimited hours) or decimal hours.
  const parseHHMM = s => {
    if (!s) return 0;
    s = String(s).trim();
    // Accept arbitrary digits before the colon (unlimited hours)
    let m = s.match(/^(\d+):([0-5]\d)$/);
    if (m) {
      // Convert hours and minutes to total minutes
      return (parseInt(m[1],10) * 60) + parseInt(m[2],10);
    }
    // Support decimals using dot or comma (e.g. 2.5 -> 2h30)
    const f = parseFloat(s.replace(',', '.'));
    return isNaN(f) ? 0 : Math.round(f * 60);
  };
  const fmtHHMM = mins => {
    const sign = mins < 0 ? "-" : "";
    mins = Math.abs(mins);
    return sign + pad2(Math.floor(mins/60)) + ":" + pad2(mins%60);
  };

  // Debounce helper: waits for a pause before executing a function. Useful for text filtering.
  const debounce = (fn, ms = 200) => {
    let timer;
    return (...args) => {
      clearTimeout(timer);
      timer = setTimeout(() => { fn.apply(this, args); }, ms);
    };
  };

  /**
   * Persist current filter selections (resource, project and text) to localStorage.
   * This avoids resetting filters when the user reloads or navigates away.
   */
  const persistFilterState = () => {
    try {
      const data = {
        selResId: state.selResId || '',
        selProjName: state.selProjName || '',
        filterText: state.filterText || ''
      };
      localStorage.setItem('rv-filters', JSON.stringify(data));
    } catch(e){}
  };

  /**
   * Restore filter selections from localStorage into the state. If there is no
   * saved data, defaults remain unchanged. Should be called before the first render.
   */
  const restoreFilterState = () => {
    try {
      const raw = localStorage.getItem('rv-filters');
      if (!raw) return;
      const data = JSON.parse(raw) || {};
      if (typeof data.selResId === 'string') state.selResId = data.selResId;
      if (typeof data.selProjName === 'string') state.selProjName = data.selProjName;
      if (typeof data.filterText === 'string') state.filterText = data.filterText.toLowerCase();
    } catch(e){}
  };

  /**
   * Add data-label attributes to each <td> in a table with class .tbl based on
   * the column header text. This is used by the responsive CSS to display stacked
   * tables on small screens with context labels.
   */
  const makeTableResponsive = (root = document) => {
    try {
      root.querySelectorAll('table.tbl').forEach(tbl => {
        const heads = Array.from(tbl.querySelectorAll('thead th')).map(th => th.textContent.trim());
        tbl.querySelectorAll('tbody tr').forEach(tr => {
          Array.from(tr.children).forEach((td, i) => {
            if (!td.hasAttribute('data-label') && heads[i]) {
              td.setAttribute('data-label', heads[i]);
            }
          });
        });
      });
    } catch(e) {}
  };
  const save = () => {
    try {
      localStorage.setItem(DB, JSON.stringify({ thresholdMin: state.thresholdMin, externos: state.externos }));
    } catch (e) {}
    // Notify external listeners that the horas externals data has changed. This allows the main
    // application (app.js) to persist hours back into the BD file whenever the user edits
    // the ledger or contracted hours. Guard against undefined.
    try {
      if (typeof window.onHorasExternosChange === 'function') {
        window.onHorasExternosChange();
      }
    } catch (e) {}
  };
  const load = () => { try { const raw = localStorage.getItem(DB); if (raw){ const o=JSON.parse(raw); state.thresholdMin=o.thresholdMin||state.thresholdMin; state.externos=o.externos||{}; } } catch(e){} };

  /**
   * Normalize previously saved data structures for externals.
   * Earlier versions stored a single contractedMin value and a flat array of project names.
   * Convert those into a new structure where `projetos` is an object keyed by project name,
   * each containing a `contratadoMin`. If no projects were specified, create a default
   * project called "Geral" to hold any existing contracted minutes.
   */
  const normalize = () => {
    Object.keys(state.externos).forEach(id => {
      const ext = state.externos[id];
      if (!ext) return;
      // ensure dias, horasDiaMin and ledger exist
      if (!ext.dias) ext.dias = {seg:true,ter:true,qua:true,qui:true,sex:true,sab:false,dom:false};
      if (typeof ext.horasDiaMin !== 'number') ext.horasDiaMin = 0;
      if (!Array.isArray(ext.ledger)) ext.ledger = [];
      // If projetos is not an object (might be array or missing), convert it
      const isObj = ext.projetos && typeof ext.projetos === 'object' && !Array.isArray(ext.projetos);
      if (!isObj) {
        const list = Array.isArray(ext.projetos) ? ext.projetos : [];
        const totalContr = ext.contratadoMin || 0;
        const projObj = {};
        if (list.length) {
          // assign total contracted minutes to the first project, others get zero
          list.forEach((name, idx) => {
            if (!name) return;
            projObj[name] = { contratadoMin: idx === 0 ? totalContr : 0 };
          });
        } else {
          // default project name
          projObj['Geral'] = { contratadoMin: totalContr };
        }
        ext.projetos = projObj;
        delete ext.contratadoMin;
      }
    });
  };

  // File System Access API (optional)
  const setFolderStatus = (extra) => {
    const el = q('#rv-folder-status');
    if (el) el.textContent = (state.folder? 'Pasta selecionada ✓' : 'Nenhuma pasta selecionada') + (extra? ' — '+extra : '');
  };
  const pickFolder = async () => {
    if (!('showDirectoryPicker' in window)) { alert('Navegador sem suporte à pasta. Usaremos armazenamento local.'); return; }
    try {
      const h = await window.showDirectoryPicker({id:'rv-enh'});
      state.folder = h;
      setFolderStatus();
      await saveToFolder();
    } catch(e){ if (e && e.name!=='AbortError') alert('Falha: '+e.message); }
  };
  const saveToFolder = async () => {
    if (!state.folder) return;
    const fh = await state.folder.getFileHandle('dados_enhancer.json', {create:true});
    const w = await fh.createWritable();
    await w.write(JSON.stringify({thresholdMin: state.thresholdMin, externos: state.externos}));
    await w.close();
    setFolderStatus('Salvo');
  };
  const reloadFromFolder = async () => {
    if (!state.folder) return;
    try {
      const fh = await state.folder.getFileHandle('dados_enhancer.json', {create:false});
      const f = await fh.getFile();
      const txt = await f.text();
      const obj = JSON.parse(txt);
      state.thresholdMin = obj.thresholdMin || state.thresholdMin;
      state.externos = obj.externos || state.externos;
      save();
      render();
      setFolderStatus('Carregado');
    } catch(e){ alert('Falha ao carregar: '+e.message); }
  };

  // Read resources from multiple sources to avoid blanks
  const getExternos = () => {
    // 1) Try global variable
    if (Array.isArray(window.resources)) {
      const arr = window.resources.filter(r => String(r.tipo||'').toLowerCase()==='externo');
      if (arr.length) return arr.map(r => ({ id: r.id ?? r.nome, nome: r.nome, tipo: r.tipo, projetos: r.projetos||[] }));
    }
    // 2) Try localStorage the app may use
    try {
      const raw = localStorage.getItem('rp_resources_v2') || localStorage.getItem('rp_resources');
      if (raw) {
        const arr = JSON.parse(raw);
        if (Array.isArray(arr)) {
          const out = arr.filter(r => String(r.tipo||'').toLowerCase()==='externo').map(r => ({ id: r.id ?? r.nome, nome: r.nome, tipo: r.tipo, projetos: r.projetos||[] }));
          if (out.length) return out;
        }
      }
    } catch {}
    // 3) Try DOM table
    const rows = qa('#tblRecursos tbody tr');
    if (rows.length) {
      const out = [];
      rows.forEach(tr => {
        const tds = qa('td', tr);
        if (tds.length >= 2) {
          const nome = tds[0].textContent.trim();
          const tipo = tds[1].textContent.trim();
          const projetos = (tds[4]?.textContent || '').split(',').map(s=>s.trim()).filter(Boolean);
          if (nome && tipo.toLowerCase().includes('extern')) out.push({ id: nome, nome, tipo:'externo', projetos });
        }
      });
      if (out.length) return out;
    }
    return [];
  };

  // UI injection (robust): if nav.tabs not found, add a floating button
  // We intentionally disable the automatic injection of the top-level folder selection toolbar
  // (rv-toolbar) because the main application now provides its own controls to select a data
  // directory and perform exports. Keeping two sets of controls was causing confusion and
  // duplicate UI elements. The hours panel tab is still injected below.
  const ensureUI = () => {
    // Tab button
    let btn = q('#tab-horas-btn');
    if (!btn) {
      const tabs = q('nav.tabs');
      btn = document.createElement('button');
      btn.id = 'tab-horas-btn';
      btn.className = 'tab';
      btn.textContent = 'Gestão de Horas (Externos)';
      // Use addEventListener with stopPropagation to prevent the document-level tab switcher
      btn.addEventListener('click', (ev) => {
        // Stop the click from bubbling to the global tab handler in app.js
        ev.stopPropagation();
        openPanel();
      });
      if (tabs) { tabs.appendChild(btn); } else {
        // fallback floating button
        btn.style.cssText = 'position:fixed;bottom:16px;right:16px;z-index:9999;border-radius:999px;padding:10px 14px;background:#fff;border:1px solid #ddd;box-shadow:0 2px 8px rgba(0,0,0,.08)';
        document.body.appendChild(btn);
      }

      // Install a leave guard: when clicking on any other tab (not ours), remove the active
      // state from our panel and our tab button. This avoids being locked in the hours tab.
      const nav = q('nav.tabs');
      if (nav && !nav.__rvLeaveGuard) {
        nav.__rvLeaveGuard = true;
        nav.addEventListener('click', (ev) => {
          const target = ev.target.closest('.tab');
          if (!target) return;
          if (target.id !== 'tab-horas-btn') {
            const ourPanel = q('#tab-horas-panel');
            const ourBtn = q('#tab-horas-btn');
            ourPanel?.classList.remove('active');
            ourBtn?.classList.remove('active');
            // Let the app's own tab handler handle showing the clicked tab.
          }
        }, true);
      }
    }
    // Panel
    if (!q('#tab-horas-panel')) {
      const host = q('#tab-plan')?.parentElement || q('.container') || document.body;
      const panel = document.createElement('div');
      panel.id = 'tab-horas-panel';
      panel.className = 'tabpanel';
      // Do not set inline display, rely on the CSS class `.tabpanel` to hide panels.
      // Construct the panel markup. Show the alerts block first so users immediately see resources near the limit
      panel.innerHTML = `
        <section class="panel">
          <!-- Alerts: show on top -->
          <div class="panel" style="border:1px dashed #bbb;margin-bottom:10px">
            <h3>⚠️ Recursos com horas próximas do fim</h3>
            <div id="rv-alertas"></div>
          </div>
          <h2>Gestão de Horas (Somente Recursos Externos)</h2>
          <div class="actions" style="display:flex;gap:8px;align-items:center;margin:8px 0">
            <label>Limiar de alerta (h) <input id="rv-th" type="number" min="0" style="width:90px"></label>
            <button id="rv-th-apply">Aplicar</button>
          </div>
          <p class="muted small">Suporte a minutos (HH:MM). Ex.: 120:00 − 08:30 = 111:30.</p>
          <!-- Filters: recurso/projeto/busca -->
          <div id="rv-filters" style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:8px 0"></div>
          <div id="rv-externos"></div>
        </section>`;
      host.appendChild(panel);
      q('#rv-th-apply').onclick = () => { const v = parseInt(q('#rv-th').value||'0',10); state.thresholdMin=(isNaN(v)?0:v)*60; save(); renderAlerts(); };
    }
  };

  const openPanel = () => {
    // Show our panel by toggling the CSS class instead of messing with inline style.
    // Remove active class from all tab panels and add it only to ours.
    const panel = q('#tab-horas-panel');
    if (!panel) return;
    qa('.tabpanel').forEach(p => p.classList.remove('active'));
    panel.classList.add('active');
    // Update tab buttons: mark our button as active, others will be handled by app's own handlers.
    qa('nav.tabs .tab').forEach(b => b.classList.remove('active'));
    const tb = q('#tab-horas-btn');
    if (tb) tb.classList.add('active');
    render();
  };

  function render(){
    const cont = q('#rv-externos'); if (!cont) return;
    let recs = getExternos();
    q('#rv-th').value = Math.round((state.thresholdMin||0)/60);
    cont.innerHTML = '';
    // Populate filter controls
    const fDiv = q('#rv-filters');
    if (fDiv) {
      // Ensure selected resource is still valid (except for special '__all__' value)
      const ids = recs.map(r => r.id || r.nome);
      if (state.selResId && state.selResId !== '__all__' && !ids.includes(state.selResId)) {
        state.selResId = '__all__';
      }
      // If no resource selected (empty string or undefined) and there is at least one, pick the first one by default.
      // This reduces initial visual clutter by showing only the first resource. Users can choose "Todos" later.
      if ((!state.selResId || state.selResId === '') && ids.length > 0) {
        state.selResId = ids[0];
      }
      // Build resource options (including a "Todos" option using the special '__all__' value)
      let resOpts = '<option value="__all__"' + (state.selResId === '__all__' ? ' selected' : '') + '>Todos</option>';
      recs.forEach(r => {
        const id = r.id || r.nome;
        const selected = state.selResId === id ? ' selected' : '';
        resOpts += `<option value="${id}"${selected}>${r.nome}</option>`;
      });
      // Build project options based on selected resource
      let projOpts = '<option value="">Todos</option>';
      if (state.selResId && state.selResId !== '__all__') {
        const selExt = state.externos[state.selResId];
        const projNames = selExt ? Object.keys(selExt.projetos || {}) : [];
        projNames.forEach(name => {
          const sel = state.selProjName === name ? ' selected' : '';
          projOpts += `<option value="${name}"${sel}>${name}</option>`;
        });
        // Ensure selected project is valid
        if (state.selProjName && !projNames.includes(state.selProjName)) state.selProjName = '';
      } else {
        state.selProjName = '';
      }
      // Render filters
      fDiv.innerHTML = `
        <label>Recurso: <select id="rv-filter-res">${resOpts}</select></label>
        <label>Projeto: <select id="rv-filter-proj">${projOpts}</select></label>
        <label>Busca: <input id="rv-filter-text" value="${state.filterText||''}" placeholder="Buscar"/></label>
      `;
      // Attach handlers
      const resSel = fDiv.querySelector('#rv-filter-res');
      const projSel = fDiv.querySelector('#rv-filter-proj');
      const textInp = fDiv.querySelector('#rv-filter-text');
      if (resSel) resSel.onchange = e => {
        state.selResId = e.target.value;
        // reset project when resource changes
        state.selProjName = '';
        persistFilterState();
        render();
      };
      if (projSel) projSel.onchange = e => {
        state.selProjName = e.target.value;
        persistFilterState();
        render();
      };
      if (textInp) {
        // Debounce the text input to avoid re-rendering on every keystroke
        textInp.oninput = debounce(e => {
          state.filterText = (textInp.value || '').toLowerCase();
          persistFilterState();
          render();
        }, 200);
      }
    }
    // Apply filters to recs
    if (state.selResId && state.selResId !== '__all__') {
      recs = recs.filter(r => (r.id || r.nome) === state.selResId);
    }
    if (state.filterText) {
      const t = state.filterText;
      recs = recs.filter(r => {
        const nameMatch = r.nome.toLowerCase().includes(t);
        let projMatch = false;
        const ext = state.externos[r.id || r.nome];
        if (ext) {
          projMatch = Object.keys(ext.projetos || {}).some(pn => pn.toLowerCase().includes(t));
        } else if (Array.isArray(r.projetos)) {
          projMatch = r.projetos.some(pn => pn && pn.toLowerCase().includes(t));
        }
        return nameMatch || projMatch;
      });
    }
    if (!recs.length){
      cont.innerHTML = '<p class="muted">Não há recursos externos visíveis. Acesse a aba Planejamento primeiro ou cadastre recursos externos.</p>';
      renderAlerts();
      return;
    }
    // Use a document fragment to minimize DOM reflows when rendering many cards
    const _frag = document.createDocumentFragment();
    recs.forEach(r => {
      const id = r.id || r.nome;
      // initialize state for this external if absent
      if (!state.externos[id]) {
        const projObj = {};
        (r.projetos || []).forEach(name => {
          if (name && !projObj[name]) projObj[name] = { contratadoMin: 0 };
        });
        state.externos[id] = {
          horasDiaMin: 0,
          dias: { seg:true, ter:true, qua:true, qui:true, sex:true, sab:false, dom:false },
          ledger: [],
          projetos: projObj
        };
      }
      const m = state.externos[id];
      // compute used minutes per project
      const usedPerProj = {};
      (m.ledger || []).forEach(e => {
        const pname = e.projeto || '';
        if (!usedPerProj[pname]) usedPerProj[pname] = 0;
        usedPerProj[pname] += e.minutos || 0;
      });
      // compute total contracted, used and saldo across all projects
      const totalContract = Object.values(m.projetos || {}).reduce((sum, p) => sum + (p.contratadoMin || 0), 0);
      const totalUsed = Object.values(usedPerProj).reduce((sum, v) => sum + v, 0);
      const totalSaldo = totalContract - totalUsed;
      // build project rows
      const projectRows = Object.keys(m.projetos || {}).map(name => {
        const contrMin = m.projetos[name].contratadoMin || 0;
        const usedMin = usedPerProj[name] || 0;
        const saldoMin = contrMin - usedMin;
        return `<tr data-proj="${name}"><td>${name}</td><td><input class="rv-proj-contr" data-id="${id}" data-proj="${name}" value="${fmtHHMM(contrMin)}"/></td><td>${fmtHHMM(usedMin)}</td><td>${fmtHHMM(saldoMin)}</td><td><button class="rv-proj-del" data-id="${id}" data-proj="${name}">Excluir</button></td></tr>`;
      }).join('');
      // options for project select
      const projOptions = Object.keys(m.projetos || {}).map(name => `<option value="${name}">${name}</option>`).join('');
      const card = document.createElement('div');
      card.className = 'rv-card';
      card.innerHTML = `
        <div class="rv-grid">
          <div><span class="rv-badge">Recurso</span> <b>${r.nome}</b></div>
          <label>Horas por dia (HH:MM)<input class="rv-dia" data-id="${id}" value="${fmtHHMM(m.horasDiaMin)}"/></label>
          <div><div class="muted small">Dias de trabalho</div><div class="rv-days" data-id="${id}">${['seg','ter','qua','qui','sex','sab','dom'].map(d=>`<label><input type="checkbox" data-day="${d}" ${m.dias[d]?'checked':''}/> ${{seg:'Seg',ter:'Ter',qua:'Qua',qui:'Qui',sex:'Sex',sab:'Sáb',dom:'Dom'}[d]}</label>`).join(' ')}</div></div>
          <div><div><span class="rv-badge">Consumido total</span> <b>${fmtHHMM(totalUsed)}</b></div><div><span class="rv-badge">Saldo total</span> <b>${fmtHHMM(totalSaldo)}</b></div></div>
        </div>
        <div class="rv-proj-container">
          <table class="tbl rv-proj-table"><thead><tr><th>Projeto</th><th>Contratado</th><th>Consumido</th><th>Saldo</th><th></th></tr></thead><tbody>
            ${projectRows}
          </tbody></table>
          <button class="rv-proj-add" data-id="${id}">+ Projeto</button>
        </div>
        <div class="rv-entry">
          <input type="date" class="rv-date-start" data-id="${id}"/>
          <input type="date" class="rv-date-end" data-id="${id}"/>
          <select class="rv-rec" data-id="${id}">
            <option value="once">Única</option>
            <option value="daily">Diária</option>
            <option value="weekly">Semanal</option>
            <option value="monthly">Mensal</option>
          </select>
          <input class="rv-hours" placeholder="Horas (HH:MM ou decimal)" data-id="${id}"/>
          <select class="rv-tipo" data-id="${id}"><option value="normal">Normal</option><option value="extra">Extra</option></select>
          <select class="rv-proj-select" data-id="${id}"><option value="">Selecione o projeto</option>${projOptions}</select>
          <button class="rv-add" data-id="${id}">Adicionar</button>
        </div>
        <table class="tbl" style="margin-top:6px"><thead><tr><th>Data</th><th>Horas</th><th>Tipo</th><th>Projeto</th><th></th></tr></thead><tbody class="rv-hist" data-id="${id}">${(m.ledger||[]).map((e,i)=>`<tr><td>${e.date}</td><td>${fmtHHMM(e.minutos)}</td><td>${e.tipo}</td><td>${e.projeto||''}</td><td><button class="rv-del" data-id="${id}" data-index="${i}">Excluir</button></td></tr>`).join('')}</tbody></table>
      `;
      _frag.appendChild(card);
    });
    // Insert all cards at once
    cont.appendChild(_frag);
    // bindings for dynamic elements
    // update hours per day
    cont.querySelectorAll('.rv-dia').forEach(el => el.onchange = e => {
      const id = e.target.dataset.id;
      state.externos[id].horasDiaMin = parseHHMM(e.target.value);
      save();
      render();
      renderAlerts();
    });
    // update working days
    cont.querySelectorAll('.rv-days input[type=checkbox]').forEach(cb => cb.onchange = e => {
      const pid = e.target.closest('.rv-days').dataset.id;
      const day = e.target.dataset.day;
      state.externos[pid].dias[day] = e.target.checked;
      save();
    });
    // project contract edits
    cont.querySelectorAll('.rv-proj-contr').forEach(inp => inp.onchange = e => {
      const id = e.target.dataset.id;
      const proj = e.target.dataset.proj;
      const mins = parseHHMM(e.target.value);
      if (!state.externos[id].projetos[proj]) state.externos[id].projetos[proj] = { contratadoMin: 0 };
      state.externos[id].projetos[proj].contratadoMin = mins;
      save();
      render();
      renderAlerts();
    });
    // project deletion
    cont.querySelectorAll('.rv-proj-del').forEach(btn => btn.onclick = e => {
      const id = e.target.dataset.id;
      const proj = e.target.dataset.proj;
      if (confirm('Excluir projeto "' + proj + '"? Os lançamentos permanecerão.')) {
        delete state.externos[id].projetos[proj];
        save();
        render();
        renderAlerts();
      }
    });
    // add new project
    cont.querySelectorAll('.rv-proj-add').forEach(btn => btn.onclick = e => {
      const id = e.target.dataset.id;
      const name = prompt('Nome do projeto:');
      if (!name) return;
      const nameTrim = name.trim();
      if (!nameTrim) return;
      if (!state.externos[id].projetos) state.externos[id].projetos = {};
      if (state.externos[id].projetos[nameTrim]) {
        alert('Projeto já existente');
        return;
      }
      let horas = prompt('Horas contratadas para ' + nameTrim + ' (HH:MM ou decimal):');
      if (horas == null) return;
      horas = horas.trim();
      const mins = parseHHMM(horas);
      if (mins <= 0) {
        alert('Horas inválidas');
        return;
      }
      state.externos[id].projetos[nameTrim] = { contratadoMin: mins };
      save();
      render();
      renderAlerts();
    });
    // add ledger entries
    cont.querySelectorAll('.rv-add').forEach(btn => btn.onclick = e => {
      const id = e.target.dataset.id;
      const card = e.target.closest('.rv-card');
      const startInp = card.querySelector('.rv-date-start[data-id="' + id + '"]');
      const endInp = card.querySelector('.rv-date-end[data-id="' + id + '"]');
      const recSel = card.querySelector('.rv-rec[data-id="' + id + '"]');
      const hoursInp = card.querySelector('.rv-hours[data-id="' + id + '"]');
      const tipoSel = card.querySelector('.rv-tipo[data-id="' + id + '"]');
      const projSel = card.querySelector('.rv-proj-select[data-id="' + id + '"]');
      const startDateStr = startInp && startInp.value;
      const endDateStr = endInp && endInp.value;
      const rec = recSel ? recSel.value : 'once';
      const hoursStr = hoursInp ? hoursInp.value : '';
      const tipo = tipoSel ? tipoSel.value : 'normal';
      const proj = projSel ? projSel.value : '';
      if (!startDateStr) { alert('Informe data de início'); return; }
      if (!hoursStr) { alert('Informe horas'); return; }
      const minutes = parseHHMM(hoursStr);
      if (minutes <= 0) { alert('Horas inválidas'); return; }
      if (!proj) { alert('Selecione o projeto'); return; }
      const startDate = new Date(startDateStr);
      const endDate = endDateStr ? new Date(endDateStr) : new Date(startDateStr);
      if (isNaN(startDate) || isNaN(endDate)) { alert('Data inválida'); return; }
      if (endDate < startDate) { alert('Data final anterior à inicial'); return; }
      // Determine list of dates for the recurrence respecting working days (diasCfg)
      const fmtDate = d => {
        // Normalize to YYYY-MM-DD to avoid timezone issues
        const year = d.getFullYear();
        const month = d.getMonth();
        const day = d.getDate();
        return new Date(year, month, day).toISOString().slice(0, 10);
      };
      const diasCfg = (state.externos[id] && state.externos[id].dias) ? state.externos[id].dias : {};
      const dowMap = { 0:'dom', 1:'seg', 2:'ter', 3:'qua', 4:'qui', 5:'sex', 6:'sab' };
      const datesSet = new Set();
      if (rec === 'daily' || rec === 'weekly') {
        // For daily and weekly recurrences, iterate through each day in the range and include it if it's a working day.
        let d = new Date(startDate);
        while (d <= endDate) {
          const key = dowMap[d.getDay()];
          if (diasCfg[key]) {
            datesSet.add(fmtDate(d));
          }
          // Move to next day
          d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
        }
      } else if (rec === 'monthly') {
        // Monthly recurrence: once per month on the same day (or last day if shorter), respecting working days
        const targetDay = startDate.getDate();
        let current = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
        while (current <= endDate) {
          const y = current.getFullYear();
          const m = current.getMonth();
          // Determine the target day for this month (cap at last day)
          const lastDay = new Date(y, m + 1, 0).getDate();
          const day = Math.min(targetDay, lastDay);
          const d = new Date(y, m, day);
          if (d >= startDate && d <= endDate) {
            const key = dowMap[d.getDay()];
            if (diasCfg[key]) {
              datesSet.add(fmtDate(d));
            }
          }
          // Move to first day of next month
          current = new Date(y, m + 1, 1);
        }
      } else {
        // Single entry: allow any day (working or not)
        datesSet.add(fmtDate(startDate));
      }
      const dates = Array.from(datesSet);
      const ledger = state.externos[id].ledger;
      dates.forEach(dateStr => {
        // Sum minutes for this date across all entries
        const totalForDate = ledger.filter(e => e.date === dateStr).reduce((acc, e) => acc + (e.minutos || 0), 0);
        // Check if there is an existing entry for this date/project/tipo to update
        const existing = ledger.find(e => e.date === dateStr && e.projeto === proj && e.tipo === tipo);
        const existingMin = existing ? existing.minutos : 0;
        // Proposed total includes current minutes (if existing) and new minutes
        const proposedTotal = totalForDate + minutes;
        if (proposedTotal > 24 * 60) {
          alert('O lançamento para ' + dateStr + ' excede o limite de 24 horas no dia. Ajuste as horas ou escolha outro dia.');
          return; // Skip adding this date
        }
        if (existing) {
          existing.minutos += minutes;
        } else {
          ledger.push({ date: dateStr, minutos: minutes, tipo: tipo, projeto: proj });
        }
      });
      // Persist changes and re-render
      save();
      render();
      renderAlerts();
    });
    // delete ledger entries
    cont.querySelectorAll('.rv-del').forEach(btn => btn.onclick = e => {
      const id = e.target.dataset.id;
      const idx = +e.target.dataset.index;
      state.externos[id].ledger.splice(idx, 1);
      save();
      render();
      renderAlerts();
    });
    // finally, update alerts after binding
    renderAlerts();
    // Update table labels for responsive view
    makeTableResponsive(cont);
  };

  const renderAlerts = () => {
    const el = q('#rv-alertas'); if (!el) return;
    el.innerHTML = '';
    const recs = getExternos();
    const items = [];
    recs.forEach(r => {
      const id = r.id || r.nome;
      const m = state.externos[id]; if (!m) return;
      // calculate total contracted minutes across all projects
      const totalContract = Object.values(m.projetos || {}).reduce((sum, p) => sum + (p.contratadoMin || 0), 0);
      const used = (m.ledger||[]).reduce((acc,e)=>acc+(e.minutos||0),0);
      const saldo = totalContract - used;
      if (saldo <= (state.thresholdMin||0)) items.push({ nome: r.nome, saldo });
    });
    if (!items.length){ el.innerHTML = '<p class="muted">Nenhum recurso no limite configurado.</p>'; return; }
    items.sort((a,b)=>a.saldo-b.saldo);
    const tbl = document.createElement('table'); tbl.className = 'tbl';
    tbl.innerHTML = '<thead><tr><th>Recurso</th><th>Saldo restante</th></tr></thead>';
    const tb = document.createElement('tbody'); tbl.appendChild(tb);
    items.forEach(i=>{ const tr=document.createElement('tr'); tr.innerHTML = '<td>'+i.nome+'</td><td>'+fmtHHMM(i.saldo)+'</td>'; tb.appendChild(tr); });
    el.appendChild(tbl);
  };

  // Observe resources table to avoid blank panel if data loads later
  const observeResources = () => {
    const rec = q('#tblRecursos tbody');
    if (!rec) return;
    const mo = new MutationObserver(() => { const p = q('#tab-horas-panel'); if (p && p.classList.contains('active')) render(); });
    mo.observe(rec, {childList:true, subtree:true});
  };

  const ensureStyles = () => {
    if (!q('#rv-enhancer-css')) {
      const css = document.createElement('style'); css.id = 'rv-enhancer-css';
      css.textContent = `.rv-card{border:1px solid #eee;border-radius:10px;padding:10px;margin:8px 0;background:#fff}
.rv-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:8px;align-items:start}
.rv-badge{display:inline-block;border:1px solid #ddd;border-radius:999px;padding:2px 6px;font-size:12px;background:#fafafa}
/* entry grid: start date | end date | recurrence | horas | tipo | projeto | botão */
.rv-entry{display:grid;grid-template-columns:120px 120px 120px 100px 100px 1fr 100px;gap:6px;align-items:center;margin-top:6px}
`;
      document.head.appendChild(css);
    }
  };

  document.addEventListener('DOMContentLoaded', () => {
    load();
    // Restore previously saved filter selections (if any) before rendering
    restoreFilterState();
    // Normalize existing data to use the new per-project contract structure
    normalize();
    ensureUI();
    ensureStyles();
    observeResources();
    // Ensure table data-label attributes are set on initial load
    makeTableResponsive(document);
    // If user navigates directly to our tab via button, render will happen in openPanel()
    // To avoid blank on first click due to not-yet-loaded resources, render anyway once:
    setTimeout(()=>{
      const p=q('#tab-horas-panel');
      // Only render if the panel is currently active (visible)
      if (p && p.classList.contains('active')) render();
    }, 300);
  });

  /*
   * Expose helpers on the window object so the main application can
   * synchronize the Gestão de Horas (Externos) data into the BD file.  The
   * functions below allow reading the current ledger entries and replacing
   * the ledger with new data loaded from a BD.  They intentionally avoid
   * exposing or modifying other parts of the internal `state` such as
   * contracted hours and selected days of the week, since those are
   * application-specific. If you need to persist contracted minutes or
   * day-of-week availability, you can extend these helpers accordingly.
   */
  try {
    // Returns a flat array of ledger entries. Each entry contains: id, date, minutos, tipo, projeto.
    window.getHorasExternosData = () => {
      const list = [];
      try {
        Object.keys(state.externos || {}).forEach(id => {
          const ext = state.externos[id];
          if (!ext) return;
          (ext.ledger || []).forEach(item => {
            list.push({ id, date: item.date, minutos: item.minutos, tipo: item.tipo || '', projeto: item.projeto || '' });
          });
        });
      } catch (e) {}
      return list;
    };
    // Accepts a flat array of ledger entries and rebuilds the externos ledger. Each entry
    // should provide id (resource identifier), date, minutos (number of minutes), tipo and projeto.
    window.setHorasExternosData = (entries = []) => {
      try {
        const newExternos = {};
        entries.forEach(ent => {
          const id = ent.id || ent.resourceId || ent.colaborador || '';
          if (!id) return;
          if (!newExternos[id]) {
            newExternos[id] = { dias: {seg:true,ter:true,qua:true,qui:true,sex:true,sab:false,dom:false}, horasDiaMin: 0, ledger: [], projetos: {} };
          }
          newExternos[id].ledger.push({ date: ent.date, minutos: Number(ent.minutos) || 0, tipo: ent.tipo || '', projeto: ent.projeto || '' });
        });
        state.externos = newExternos;
        save();
        // Re-render if the panel is visible
        if (typeof render === 'function') {
          render();
        }
      } catch (e) {}
    };
    // Provide a default no-op for the onHorasExternosChange callback if not already defined.
    if (typeof window.onHorasExternosChange !== 'function') {
      window.onHorasExternosChange = null;
    }

    /**
     * Returns a list of configuration objects for each externo resource. Each configuration
     * entry contains the resource id, the contracted hours per day (horasDia) formatted
     * as HH:MM, a comma-separated string of working days (dias), and a semicolon-separated
     * list of project definitions (projeto:HH:MM). These configurations are used when
     * persisting and restoring the Gestão de Horas settings to and from the BD file.
     */
    window.getHorasExternosConfig = () => {
      const list = [];
      try {
        Object.keys(state.externos || {}).forEach(id => {
          const ext = state.externos[id];
          if (!ext) return;
          // Build list of working days in order
          const diasArr = [];
          ['seg','ter','qua','qui','sex','sab','dom'].forEach(d => {
            if (ext.dias && ext.dias[d]) diasArr.push(d);
          });
          // Build list of project:hours pairs
          const projPairs = [];
          if (ext.projetos && typeof ext.projetos === 'object') {
            Object.keys(ext.projetos).forEach(name => {
              const p = ext.projetos[name];
              const min = p && typeof p.contratadoMin === 'number' ? p.contratadoMin : 0;
              projPairs.push(name + ':' + fmtHHMM(min));
            });
          }
          list.push({
            id: id,
            horasDia: fmtHHMM(ext.horasDiaMin || 0),
            dias: diasArr.join(','),
            projetos: projPairs.join(';')
          });
        });
      } catch (e) {}
      return list;
    };

    /**
     * Accepts a list of configuration objects and applies them to the internal state. Each
     * configuration entry should provide an id, horasDia (string in HH:MM or decimal),
     * dias (comma/semicolon-separated list of day codes) and projetos (semicolon-separated
     * list of project:HH:MM definitions). Missing resources will be created if needed.
     * After applying, the state is saved and the UI re-rendered.
     */
    window.setHorasExternosConfig = (cfgList = []) => {
      try {
        cfgList.forEach(cfg => {
          const id = cfg.id || cfg.resourceId || cfg.colaborador || '';
          if (!id) return;
          if (!state.externos[id]) {
            // initialize structure if missing
            state.externos[id] = {
              dias: {seg:true,ter:true,qua:true,qui:true,sex:true,sab:false,dom:false},
              horasDiaMin: 0,
              ledger: [],
              projetos: {}
            };
          }
          const ext = state.externos[id];
          // horasDia
          const hd = cfg.horasDia || cfg.horasdia || cfg.horasDiaMin || '';
          let mins = 0;
          if (typeof hd === 'number') {
            mins = hd;
          } else if (hd) {
            mins = parseHHMM(hd);
          }
          ext.horasDiaMin = mins;
          // dias
          const diasStr = cfg.dias || cfg.Dias || cfg.dia || '';
          const diasObj = {seg:false,ter:false,qua:false,qui:false,sex:false,sab:false,dom:false};
          if (typeof diasStr === 'string' && diasStr.trim()) {
            diasStr.split(/[,;]/).forEach(d => {
              const dd = d.trim().toLowerCase();
              if (['seg','ter','qua','qui','sex','sab','dom'].includes(dd)) diasObj[dd] = true;
            });
          }
          ext.dias = diasObj;
          // projetos
          const projStr = cfg.projetos || cfg.Projetos || '';
          const projObj = {};
          if (typeof projStr === 'string' && projStr.trim()) {
            projStr.split(';').forEach(part => {
              if (!part) return;
              const idx = part.indexOf(':');
              if (idx < 0) return;
              const name = part.slice(0, idx).trim();
              const val = part.slice(idx + 1).trim();
              if (!name) return;
              const dur = parseHHMM(val);
              projObj[name] = { contratadoMin: dur };
            });
          }
          ext.projetos = projObj;
        });
        save();
        if (typeof render === 'function') render();
      } catch (e) {}
    };
  } catch (e) {}
})();
