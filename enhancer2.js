/* enhancer2.js - Gestão de Horas (Externos) + Toolbar FSA (robusta) + Previsão de Fim com Legendas de Feriado */
(() => {
  const DB = 'rv-enhancer-v2';
  // --- MODIFICAÇÃO: 'feriados' agora é um array de objetos {date, legend} ---
  const state = { thresholdMin: 10*60, externos: {}, folder: null, feriados: [] };
  state.selResId = '';
  state.selProjName = '';
  state.filterText = '';

  // Utils
  const q = (sel, root=document) => root.querySelector(sel);
  const qa = (sel, root=document) => Array.from(root.querySelectorAll(sel));
  const pad2 = n => String(n).padStart(2,'0');
  const toYMD = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  const fromYMD = (s) => new Date(`${s}T00:00:00`);

  const parseHHMM = s => {
    if (!s) return 0;
    s = String(s).trim();
    let m = s.match(/^(\d+):([0-5]\d)$/);
    if (m) {
      return (parseInt(m[1],10) * 60) + parseInt(m[2],10);
    }
    const f = parseFloat(s.replace(',', '.'));
    return isNaN(f) ? 0 : Math.round(f * 60);
  };
  const fmtHHMM = mins => {
    const sign = mins < 0 ? "-" : "";
    mins = Math.abs(mins);
    return sign + pad2(Math.floor(mins/60)) + ":" + pad2(mins%60);
  };

  const debounce = (fn, ms = 200) => {
    let timer;
    return (...args) => {
      clearTimeout(timer);
      timer = setTimeout(() => { fn.apply(this, args); }, ms);
    };
  };

  const persistFilterState = () => {
    try {
      const data = { selResId: state.selResId || '', selProjName: state.selProjName || '', filterText: state.filterText || '' };
      localStorage.setItem('rv-filters', JSON.stringify(data));
    } catch(e){}
  };

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
      localStorage.setItem(DB, JSON.stringify({ thresholdMin: state.thresholdMin, externos: state.externos, feriados: state.feriados }));
    } catch (e) {}
    try {
      if (typeof window.onHorasExternosChange === 'function') {
        window.onHorasExternosChange();
      }
    } catch (e) {}
  };
  const load = () => { 
    try { 
      const raw = localStorage.getItem(DB); 
      if (raw){ 
        const o=JSON.parse(raw); 
        state.thresholdMin=o.thresholdMin||state.thresholdMin; 
        state.externos=o.externos||{}; 
        state.feriados = Array.isArray(o.feriados) ? o.feriados : [];
      } 
    } catch(e){} 
  };

  const normalize = () => {
    Object.keys(state.externos).forEach(id => {
      const ext = state.externos[id];
      if (!ext) return;
      if (!ext.dias) ext.dias = {seg:true,ter:true,qua:true,qui:true,sex:true,sab:false,dom:false};
      if (typeof ext.horasDiaMin !== 'number') ext.horasDiaMin = 0;
      if (!Array.isArray(ext.ledger)) ext.ledger = [];
      const isObj = ext.projetos && typeof ext.projetos === 'object' && !Array.isArray(ext.projetos);
      if (!isObj) {
        const list = Array.isArray(ext.projetos) ? ext.projetos : [];
        const totalContr = ext.contratadoMin || 0;
        const projObj = {};
        if (list.length) {
          list.forEach((name, idx) => {
            if (!name) return;
            projObj[name] = { contratadoMin: idx === 0 ? totalContr : 0 };
          });
        } else {
          projObj['Geral'] = { contratadoMin: totalContr };
        }
        ext.projetos = projObj;
        delete ext.contratadoMin;
      }
    });
  };

  const setFolderStatus = (extra) => {
    const el = q('#rv-folder-status');
    if (el) el.textContent = (state.folder? 'Pasta selecionada ✓' : 'Nenhuma pasta selecionada') + (extra? ' — '+extra : '');
  };
  
  const saveToFolder = async () => {
    if (!state.folder) return;
    const fh = await state.folder.getFileHandle('dados_enhancer.json', {create:true});
    const w = await fh.createWritable();
    await w.write(JSON.stringify({thresholdMin: state.thresholdMin, externos: state.externos, feriados: state.feriados }));
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
      state.feriados = Array.isArray(obj.feriados) ? obj.feriados : [];
      save();
      render();
      setFolderStatus('Carregado');
    } catch(e){ alert('Falha ao carregar: '+e.message); }
  };
  
  const getExternos = () => {
    if (Array.isArray(window.resources)) {
      const arr = window.resources.filter(r => String(r.tipo||'').toLowerCase()==='externo');
      if (arr.length) return arr.map(r => ({ id: r.id ?? r.nome, nome: r.nome, tipo: r.tipo, projetos: r.projetos||[] }));
    }
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

  function calcularPrevisaoFim(resourceId) {
    const recurso = state.externos[resourceId];
    if (!recurso || !recurso.ledger || recurso.ledger.length === 0) {
      return null;
    }

    const totalContratadoMin = Object.values(recurso.projetos || {}).reduce((sum, p) => sum + (p.contratadoMin || 0), 0);
    const totalUtilizadoMin = recurso.ledger.reduce((sum, entry) => sum + (entry.minutos || 0), 0);
    const saldoMin = totalContratadoMin - totalUtilizadoMin;

    if (saldoMin <= 0) {
      return "Concluído";
    }

    const horasDiaMin = recurso.horasDiaMin;
    if (!horasDiaMin || horasDiaMin <= 0) {
      return "Horas/dia não definidas";
    }

    const sortedLedger = [...recurso.ledger].sort((a, b) => new Date(a.date) - new Date(b.date));
    const ultimaDataLancada = fromYMD(sortedLedger[sortedLedger.length - 1].date);
    
    let diasUteisNecessarios = Math.ceil(saldoMin / horasDiaMin);
    let dataCorrente = new Date(ultimaDataLancada);
    
    const diasTrabalho = recurso.dias || {};
    const dowMap = { 0:'dom', 1:'seg', 2:'ter', 3:'qua', 4:'qui', 5:'sex', 6:'sab' };
    const feriadosSet = new Set((state.feriados || []).map(f => f.date));

    while (diasUteisNecessarios > 0) {
      dataCorrente.setDate(dataCorrente.getDate() + 1);
      const diaDaSemana = dowMap[dataCorrente.getDay()];
      const dataYMD = toYMD(dataCorrente);

      if (diasTrabalho[diaDaSemana] && !feriadosSet.has(dataYMD)) {
        diasUteisNecessarios--;
      }
    }
    
    return toYMD(dataCorrente).split('-').reverse().join('/');
  }

  const ensureUI = () => {
    let btn = q('#tab-horas-btn');
    if (!btn) {
      const tabs = q('nav.tabs');
      btn = document.createElement('button');
      btn.id = 'tab-horas-btn';
      btn.className = 'tab';
      btn.textContent = 'Gestão de Horas (Externos)';
      btn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        openPanel();
      });
      if (tabs) { tabs.appendChild(btn); } else {
        btn.style.cssText = 'position:fixed;bottom:16px;right:16px;z-index:9999;border-radius:999px;padding:10px 14px;background:#fff;border:1px solid #ddd;box-shadow:0 2px 8px rgba(0,0,0,.08)';
        document.body.appendChild(btn);
      }
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
          }
        }, true);
      }
    }
    if (!q('#tab-horas-panel')) {
      const host = q('#tab-plan')?.parentElement || q('.container') || document.body;
      const panel = document.createElement('div');
      panel.id = 'tab-horas-panel';
      panel.className = 'tabpanel';
      panel.innerHTML = `
        <section class="panel">
          <div id="rv-alertas" class="panel" style="border:1px dashed #bbb;margin-bottom:10px"></div>
          <h2>Gestão de Horas (Somente Recursos Externos)</h2>
          <div class="actions" style="display:flex;gap:8px;align-items:center;margin:8px 0">
            <label>Limiar de alerta (h) <input id="rv-th" type="number" min="0" style="width:90px"></label>
            <button id="rv-th-apply" class="btn">Aplicar</button>
          </div>
          <p class="muted small">Suporte a minutos (HH:MM). Ex.: 120:00 − 08:30 = 111:30.</p>
          <div id="rv-filters" style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:8px 0"></div>
          <div id="rv-externos"></div>
        </section>
        <section class="panel">
            <h3>Cadastro de Feriados (para todos os recursos)</h3>
            <p class="muted small">Informe um feriado por linha, no formato AAAA-MM-DD Legenda (ex: 2025-12-25 Natal).</p>
            <textarea id="rv-feriados" rows="8" style="width:100%; font-family:monospace;"></textarea>
            <button id="rv-feriados-save" class="btn" style="margin-top:8px;">Salvar Feriados</button>
        </section>`;
      host.appendChild(panel);
      q('#rv-th-apply').onclick = () => { const v = parseInt(q('#rv-th').value||'0',10); state.thresholdMin=(isNaN(v)?0:v)*60; save(); renderAlerts(); };
      
      q('#rv-feriados-save').onclick = () => {
        const txt = q('#rv-feriados').value || '';
        state.feriados = txt.split('\n').map(line => {
          const trimmedLine = line.trim();
          if (!trimmedLine) return null;
          const firstSpaceIndex = trimmedLine.indexOf(' ');
          if (firstSpaceIndex === -1) {
            return { date: trimmedLine, legend: '' };
          }
          const date = trimmedLine.substring(0, firstSpaceIndex);
          const legend = trimmedLine.substring(firstSpaceIndex + 1).trim();
          return { date, legend };
        }).filter(Boolean);
        save();
        render(); 
        alert('Feriados salvos com sucesso!');
      };
    }
  };

  const openPanel = () => {
    const panel = q('#tab-horas-panel');
    if (!panel) return;
    qa('.tabpanel').forEach(p => p.classList.remove('active'));
    panel.classList.add('active');
    qa('nav.tabs .tab').forEach(b => b.classList.remove('active'));
    const tb = q('#tab-horas-btn');
    if (tb) tb.classList.add('active');
    render();
  };

  function render(){
    // ===== INÍCIO DA MODIFICAÇÃO =====
    // Adia o redesenho se o usuário estiver digitando em um campo de texto,
    // mas permite a atualização após cliques em botões.
    const panel = q('#tab-horas-panel');
    const activeEl = document.activeElement;
    if (panel && panel.contains(activeEl)) {
        const tagName = activeEl.tagName.toUpperCase();
        if (tagName === 'INPUT' || tagName === 'TEXTAREA') {
            return; // Interrompe o render se o foco estiver em um campo de texto
        }
    }
    // ===== FIM DA MODIFICAÇÃO =====

    const cont = q('#rv-externos'); if (!cont) return;
    const feriadosTextarea = q('#rv-feriados');
    if (feriadosTextarea) {
        if (document.activeElement !== feriadosTextarea) {
            feriadosTextarea.value = (state.feriados || []).map(f => `${f.date} ${f.legend}`.trim()).join('\n');
        }
    }

    let recs = getExternos();
    q('#rv-th').value = Math.round((state.thresholdMin||0)/60);
    cont.innerHTML = '';
    const fDiv = q('#rv-filters');
    if (fDiv) {
      const ids = recs.map(r => r.id || r.nome);
      if (state.selResId && state.selResId !== '__all__' && !ids.includes(state.selResId)) {
        state.selResId = '__all__';
      }
      if ((!state.selResId || state.selResId === '') && ids.length > 0) {
        state.selResId = ids[0];
      }
      let resOpts = '<option value="__all__"' + (state.selResId === '__all__' ? ' selected' : '') + '>Todos</option>';
      recs.forEach(r => {
        const id = r.id || r.nome;
        const selected = state.selResId === id ? ' selected' : '';
        resOpts += `<option value="${id}"${selected}>${r.nome}</option>`;
      });
      let projOpts = '<option value="">Todos</option>';
      if (state.selResId && state.selResId !== '__all__') {
        const selExt = state.externos[state.selResId];
        const projNames = selExt ? Object.keys(selExt.projetos || {}) : [];
        projNames.forEach(name => {
          const sel = state.selProjName === name ? ' selected' : '';
          projOpts += `<option value="${name}"${sel}>${name}</option>`;
        });
        if (state.selProjName && !projNames.includes(state.selProjName)) state.selProjName = '';
      } else {
        state.selProjName = '';
      }
      fDiv.innerHTML = `
        <label>Recurso: <select id="rv-filter-res">${resOpts}</select></label>
        <label>Projeto: <select id="rv-filter-proj">${projOpts}</select></label>
        <label>Busca: <input id="rv-filter-text" value="${state.filterText||''}" placeholder="Buscar"/></label>
      `;
      const resSel = fDiv.querySelector('#rv-filter-res');
      const projSel = fDiv.querySelector('#rv-filter-proj');
      const textInp = fDiv.querySelector('#rv-filter-text');
      if (resSel) resSel.onchange = e => { state.selResId = e.target.value; state.selProjName = ''; persistFilterState(); render(); };
      if (projSel) projSel.onchange = e => { state.selProjName = e.target.value; persistFilterState(); render(); };
      if (textInp) { textInp.oninput = debounce(e => { state.filterText = (textInp.value || '').toLowerCase(); persistFilterState(); render(); }, 200); }
    }
    if (state.selResId && state.selResId !== '__all__') {
      recs = recs.filter(r => (r.id || r.nome) === state.selResId);
    }
    if (state.filterText) {
      const t = state.filterText;
      recs = recs.filter(r => {
        const nameMatch = r.nome.toLowerCase().includes(t);
        let projMatch = false;
        const ext = state.externos[r.id || r.nome];
        if (ext) { projMatch = Object.keys(ext.projetos || {}).some(pn => pn.toLowerCase().includes(t)); } 
        else if (Array.isArray(r.projetos)) { projMatch = r.projetos.some(pn => pn && pn.toLowerCase().includes(t)); }
        return nameMatch || projMatch;
      });
    }
    if (!recs.length){
      cont.innerHTML = '<p class="muted">Não há recursos externos visíveis. Acesse a aba Planejamento primeiro ou cadastre recursos externos.</p>';
      renderAlerts();
      return;
    }
    const _frag = document.createDocumentFragment();
    recs.forEach(r => {
      const id = r.id || r.nome;
      if (!state.externos[id]) {
        const projObj = {};
        (r.projetos || []).forEach(name => { if (name && !projObj[name]) projObj[name] = { contratadoMin: 0 }; });
        state.externos[id] = { horasDiaMin: 0, dias: { seg:true, ter:true, qua:true, qui:true, sex:true, sab:false, dom:false }, ledger: [], projetos: projObj };
      }
      const m = state.externos[id];
      const usedPerProj = {};
      (m.ledger || []).forEach(e => {
        const pname = e.projeto || '';
        if (!usedPerProj[pname]) usedPerProj[pname] = 0;
        usedPerProj[pname] += e.minutos || 0;
      });
      const totalContract = Object.values(m.projetos || {}).reduce((sum, p) => sum + (p.contratadoMin || 0), 0);
      const totalUsed = Object.values(usedPerProj).reduce((sum, v) => sum + v, 0);
      const totalSaldo = totalContract - totalUsed;
      const projectRows = Object.keys(m.projetos || {}).map(name => {
        const contrMin = m.projetos[name].contratadoMin || 0;
        const usedMin = usedPerProj[name] || 0;
        const saldoMin = contrMin - usedMin;
        return `<tr data-proj="${name}"><td>${name}</td><td><input class="rv-proj-contr" data-id="${id}" data-proj="${name}" value="${fmtHHMM(contrMin)}"/></td><td>${fmtHHMM(usedMin)}</td><td>${fmtHHMM(saldoMin)}</td><td><button class="rv-proj-del btn danger" data-id="${id}" data-proj="${name}">Excluir</button></td></tr>`;
      }).join('');
      const projOptions = Object.keys(m.projetos || {}).map(name => `<option value="${name}">${name}</option>`).join('');
      
      const previsaoFim = calcularPrevisaoFim(id);

      const card = document.createElement('div');
      card.className = 'rv-card';
      card.innerHTML = `
        <div class="rv-grid">
          <div><span class="rv-badge">Recurso</span> <b>${r.nome}</b></div>
          <label>Horas por dia (HH:MM)<input class="rv-dia" data-id="${id}" value="${fmtHHMM(m.horasDiaMin)}"/></label>
          <div><div class="muted small">Dias de trabalho</div><div class="rv-days" data-id="${id}">${['seg','ter','qua','qui','sex','sab','dom'].map(d=>`<label><input type="checkbox" data-day="${d}" ${m.dias[d]?'checked':''}/> ${{seg:'Seg',ter:'Ter',qua:'Qua',qui:'Qui',sex:'Sex',sab:'Sáb',dom:'Dom'}[d]}</label>`).join(' ')}</div></div>
          <div>
            <div><span class="rv-badge">Consumido total</span> <b>${fmtHHMM(totalUsed)}</b></div>
            <div><span class="rv-badge">Saldo total</span> <b>${fmtHHMM(totalSaldo)}</b></div>
            <div><span class="rv-badge">Previsão de Fim</span> <b id="rv-previsao-${id}">${previsaoFim || '–'}</b></div>
          </div>
        </div>
        <div class="rv-proj-container">
          <table class="tbl rv-proj-table"><thead><tr><th>Projeto</th><th>Contratado</th><th>Consumido</th><th>Saldo</th><th></th></tr></thead><tbody>${projectRows}</tbody></table>
          <button class="rv-proj-add btn" data-id="${id}">+ Projeto</button>
        </div>
        <div class="rv-entry">
          <input type="date" class="rv-date-start" data-id="${id}"/>
          <input type="date" class="rv-date-end" data-id="${id}"/>
          <select class="rv-rec" data-id="${id}"><option value="once">Única</option><option value="daily">Diária</option><option value="weekly">Semanal</option><option value="monthly">Mensal</option></select>
          <input class="rv-hours" placeholder="Horas (HH:MM ou decimal)" data-id="${id}"/>
          <select class="rv-tipo" data-id="${id}"><option value="normal">Normal</option><option value="extra">Extra</option></select>
          <select class="rv-proj-select" data-id="${id}"><option value="">Selecione o projeto</option>${projOptions}</select>
          <button class="rv-add btn" data-id="${id}">Adicionar</button>
        </div>
        <table class="tbl" style="margin-top:6px"><thead><tr><th>Data</th><th>Horas</th><th>Tipo</th><th>Projeto</th><th></th></tr></thead><tbody class="rv-hist" data-id="${id}">${(m.ledger||[]).map((e,i)=>`<tr><td>${e.date}</td><td>${fmtHHMM(e.minutos)}</td><td>${e.tipo}</td><td>${e.projeto||''}</td><td><button class="rv-del btn danger" data-id="${id}" data-index="${i}">Excluir</button></td></tr>`).join('')}</tbody></table>
      `;
      _frag.appendChild(card);
    });
    cont.appendChild(_frag);
    cont.querySelectorAll('.rv-dia').forEach(el => el.onchange = e => { const id = e.target.dataset.id; state.externos[id].horasDiaMin = parseHHMM(e.target.value); save(); render(); renderAlerts(); });
    cont.querySelectorAll('.rv-days input[type=checkbox]').forEach(cb => cb.onchange = e => { const pid = e.target.closest('.rv-days').dataset.id; const day = e.target.dataset.day; state.externos[pid].dias[day] = e.target.checked; save(); render(); });
    cont.querySelectorAll('.rv-proj-contr').forEach(inp => inp.onchange = e => { const id = e.target.dataset.id; const proj = e.target.dataset.proj; const mins = parseHHMM(e.target.value); if (!state.externos[id].projetos[proj]) state.externos[id].projetos[proj] = { contratadoMin: 0 }; state.externos[id].projetos[proj].contratadoMin = mins; save(); render(); renderAlerts(); });
    cont.querySelectorAll('.rv-proj-del').forEach(btn => btn.onclick = e => { const id = e.target.dataset.id; const proj = e.target.dataset.proj; if (confirm('Excluir projeto "' + proj + '"? Os lançamentos permanecerão.')) { delete state.externos[id].projetos[proj]; save(); render(); renderAlerts(); } });
    cont.querySelectorAll('.rv-proj-add').forEach(btn => btn.onclick = e => { const id = e.target.dataset.id; const name = prompt('Nome do projeto:'); if (!name) return; const nameTrim = name.trim(); if (!nameTrim) return; if (!state.externos[id].projetos) state.externos[id].projetos = {}; if (state.externos[id].projetos[nameTrim]) { alert('Projeto já existente'); return; } let horas = prompt('Horas contratadas para ' + nameTrim + ' (HH:MM ou decimal):'); if (horas == null) return; horas = horas.trim(); const mins = parseHHMM(horas); if (mins <= 0) { alert('Horas inválidas'); return; } state.externos[id].projetos[nameTrim] = { contratadoMin: mins }; save(); render(); renderAlerts(); });
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
      const startDate = fromYMD(startDateStr);
      const endDate = endDateStr ? fromYMD(endDateStr) : fromYMD(startDateStr);
      if (isNaN(startDate) || isNaN(endDate)) { alert('Data inválida'); return; }
      if (endDate < startDate) { alert('Data final anterior à inicial'); return; }
      
      const diasCfg = (state.externos[id] && state.externos[id].dias) ? state.externos[id].dias : {};
      const dowMap = { 0:'dom', 1:'seg', 2:'ter', 3:'qua', 4:'qui', 5:'sex', 6:'sab' };
      const datesSet = new Set();
      if (rec === 'daily' || rec === 'weekly') {
        let d = new Date(startDate);
        while (d <= endDate) {
          const key = dowMap[d.getDay()];
          if (diasCfg[key]) { datesSet.add(toYMD(d)); }
          d.setDate(d.getDate() + 1);
        }
      } else if (rec === 'monthly') {
        const targetDay = startDate.getDate();
        let current = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
        while (current <= endDate) {
          const y = current.getFullYear();
          const m = current.getMonth();
          const lastDay = new Date(y, m + 1, 0).getDate();
          const day = Math.min(targetDay, lastDay);
          const d = new Date(y, m, day);
          if (d >= startDate && d <= endDate) {
            const key = dowMap[d.getDay()];
            if (diasCfg[key]) { datesSet.add(toYMD(d)); }
          }
          current = new Date(y, m + 1, 1);
        }
      } else {
        datesSet.add(toYMD(startDate));
      }
      const dates = Array.from(datesSet);
      const ledger = state.externos[id].ledger;
      dates.forEach(dateStr => {
        const totalForDate = ledger.filter(e => e.date === dateStr).reduce((acc, e) => acc + (e.minutos || 0), 0);
        const existing = ledger.find(e => e.date === dateStr && e.projeto === proj && e.tipo === tipo);
        const proposedTotal = totalForDate + minutes;
        if (proposedTotal > 24 * 60) { alert('O lançamento para ' + dateStr + ' excede o limite de 24 horas no dia.'); return; }
        if (existing) { existing.minutos += minutes; } 
        else { ledger.push({ date: dateStr, minutos: minutes, tipo: tipo, projeto: proj }); }
      });
      save();
      render();
      renderAlerts();
    });
    cont.querySelectorAll('.rv-del').forEach(btn => btn.onclick = e => { const id = e.target.dataset.id; const idx = +e.target.dataset.index; state.externos[id].ledger.splice(idx, 1); save(); render(); renderAlerts(); });
    renderAlerts();
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
      const totalContract = Object.values(m.projetos || {}).reduce((sum, p) => sum + (p.contratadoMin || 0), 0);
      const used = (m.ledger||[]).reduce((acc,e)=>acc+(e.minutos||0),0);
      const saldo = totalContract - used;
      if (saldo <= (state.thresholdMin||0)) items.push({ nome: r.nome, saldo });
    });
    if (!items.length){ el.innerHTML = '<h3>⚠️ Recursos com horas próximas do fim</h3><p class="muted">Nenhum recurso no limite configurado.</p>'; return; }
    items.sort((a,b)=>a.saldo-b.saldo);
    const tbl = document.createElement('table'); tbl.className = 'tbl';
    tbl.innerHTML = '<thead><tr><th>Recurso</th><th>Saldo restante</th></tr></thead>';
    const tb = document.createElement('tbody'); tbl.appendChild(tb);
    items.forEach(i=>{ const tr=document.createElement('tr'); tr.innerHTML = '<td>'+i.nome+'</td><td>'+fmtHHMM(i.saldo)+'</td>'; tb.appendChild(tr); });
    el.innerHTML = '<h3>⚠️ Recursos com horas próximas do fim</h3>';
    el.appendChild(tbl);
  };

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
.rv-entry{display:grid;grid-template-columns:120px 120px 120px 100px 100px 1fr 100px;gap:6px;align-items:center;margin-top:6px}
`;
      document.head.appendChild(css);
    }
  };

  document.addEventListener('DOMContentLoaded', () => {
    load();
    restoreFilterState();
    normalize();
    ensureUI();
    ensureStyles();
    observeResources();
    makeTableResponsive(document);
    setTimeout(()=>{
      const p=q('#tab-horas-panel');
      if (p && p.classList.contains('active')) render();
    }, 300);
  });

  try {
    window.getFeriados = () => {
      return state.feriados || [];
    };
    window.setFeriados = (feriadosList = []) => {
      try {
        state.feriados = Array.isArray(feriadosList) 
          ? feriadosList.map(f => (typeof f === 'object' ? {date: f.date, legend: f.legend || ''} : {date: f, legend: ''})).filter(f => f.date) 
          : [];
        save();
        if (typeof render === 'function') render();
      } catch (e) {}
    };
    
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
        if (typeof render === 'function') {
          render();
        }
      } catch (e) {}
    };
    if (typeof window.onHorasExternosChange !== 'function') {
      window.onHorasExternosChange = null;
    }
    window.getHorasExternosConfig = () => {
      const list = [];
      try {
        Object.keys(state.externos || {}).forEach(id => {
          const ext = state.externos[id];
          if (!ext) return;
          const diasArr = [];
          ['seg','ter','qua','qui','sex','sab','dom'].forEach(d => { if (ext.dias && ext.dias[d]) diasArr.push(d); });
          const projPairs = [];
          if (ext.projetos && typeof ext.projetos === 'object') {
            Object.keys(ext.projetos).forEach(name => {
              const p = ext.projetos[name];
              const min = p && typeof p.contratadoMin === 'number' ? p.contratadoMin : 0;
              projPairs.push(name + ':' + fmtHHMM(min));
            });
          }
          list.push({ id: id, horasDia: fmtHHMM(ext.horasDiaMin || 0), dias: diasArr.join(','), projetos: projPairs.join(';') });
        });
      } catch (e) {}
      return list;
    };
    window.setHorasExternosConfig = (cfgList = []) => {
      try {
        cfgList.forEach(cfg => {
          const id = cfg.id || cfg.resourceId || cfg.colaborador || '';
          if (!id) return;
          if (!state.externos[id]) {
            state.externos[id] = { dias: {seg:true,ter:true,qua:true,qui:true,sex:true,sab:false,dom:false}, horasDiaMin: 0, ledger: [], projetos: {} };
          }
          const ext = state.externos[id];
          const hd = cfg.horasDia || cfg.horasdia || cfg.horasDiaMin || '';
          let mins = 0;
          if (typeof hd === 'number') { mins = hd; } else if (hd) { mins = parseHHMM(hd); }
          ext.horasDiaMin = mins;
          const diasStr = cfg.dias || cfg.Dias || cfg.dia || '';
          const diasObj = {seg:false,ter:false,qua:false,qui:false,sex:false,sab:false,dom:false};
          if (typeof diasStr === 'string' && diasStr.trim()) {
            diasStr.split(/[,;]/).forEach(d => {
              const dd = d.trim().toLowerCase();
              if (['seg','ter','qua','qui','sex','sab','dom'].includes(dd)) diasObj[dd] = true;
            });
          }
          ext.dias = diasObj;
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