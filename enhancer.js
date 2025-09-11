/* enhancer.js - Gestão de Horas (Externos) */
(function(){
  const DB = 'rv-enhancer-v1';
  const state = { thresholdMin: 10*60, externos: {} };

  function parseHHMM(s) {
    if(!s) return 0;
    const m = /^([0-9]+):([0-5][0-9])$/.exec(s.trim());
    if(m) return parseInt(m[1],10)*60 + parseInt(m[2],10);
    const f = parseFloat(String(s).replace(',', '.'));
    return isNaN(f) ? 0 : Math.round(f*60);
  }
  function fmtHHMM(mins) {
    const sign = mins < 0 ? '-' : '';
    mins = Math.abs(mins);
    const h = Math.floor(mins / 60);
    const m = mins % 60;
    return sign + String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0');
  }
  function save(){
    try { localStorage.setItem(DB, JSON.stringify(state)); } catch(e){}
  }
  function load(){
    try {
      const raw = localStorage.getItem(DB);
      if(raw){
        const obj = JSON.parse(raw);
        if(obj && typeof obj === 'object'){
          state.thresholdMin = obj.thresholdMin || state.thresholdMin;
          state.externos = obj.externos || {};
        }
      }
    } catch(e){}
  }
  function getExternos(){
    try{
      const raw = localStorage.getItem('rp_resources_v2');
      if(raw){
        const arr = JSON.parse(raw);
        if(Array.isArray(arr)) return arr.filter(r => String(r.tipo||'').toLowerCase() === 'externo');
      }
    }catch(e){}
    if(Array.isArray(window.resources)){
      return window.resources.filter(r => String(r.tipo||'').toLowerCase() === 'externo');
    }
    try{
      const rows = document.querySelectorAll('#tblRecursos tbody tr');
      const list = [];
      rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        if(cells.length >= 2){
          const nome = (cells[0].textContent||'').trim();
          const tipo = (cells[1].textContent||'').trim();
          if(nome && tipo.toLowerCase() === 'externo'){
            list.push({ id:nome, nome:nome, tipo:tipo });
          }
        }
      });
      return list;
    }catch(e){}
    return [];
  }
  function ensureTab(){
    if(document.getElementById('tab-enh-externos')) return;
    const nav = document.querySelector('nav.tabs');
    if(nav){
      const btn = document.createElement('button');
      btn.textContent = 'Gestão de Horas (Externos)';
      btn.className = 'tab';
      btn.dataset.tab = 'enh';
      btn.addEventListener('click', function(){
        document.querySelectorAll('nav.tabs button.tab').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        document.querySelectorAll('.tabpanel').forEach(p => p.classList.remove('active'));
        const panel = document.getElementById('tab-enh-externos');
        if(panel) panel.classList.add('active');
        render();
      });
      nav.appendChild(btn);
    }
    const container = document.querySelector('.container');
    const panel = document.createElement('div');
    panel.id = 'tab-enh-externos';
    panel.className = 'tabpanel';
    panel.innerHTML = `
      <section class="panel">
        <h2>Gestão de Horas (Somente Recursos Externos)</h2>
        <div style="display:flex;gap:8px;align-items:center;margin-bottom:8px">
          <label>Limiar alerta (h): <input id="enh-threshold" type="number" min="0" style="width:80px"></label>
          <button id="enh-apply">Aplicar</button>
        </div>
        <p class="muted small">Horas suportam minutos (HH:MM). Ex.: contratado 120:00 e consumiu 08:30 → saldo 111:30.</p>
        <div id="enh-externos"></div>
        <div style="margin-top:12px;border:1px dashed #bbb;padding:10px;border-radius:8px">
          <h3>⚠️ Recursos com horas próximas do fim</h3>
          <div id="enh-alertas"></div>
        </div>
      </section>`;
    container.appendChild(panel);
    document.getElementById('enh-apply').addEventListener('click', function(){
      const v = parseInt(document.getElementById('enh-threshold').value || '0', 10);
      state.thresholdMin = (isNaN(v) ? 0 : v) * 60;
      save();
      renderAlerts();
    });
  }
  function render(){
    const cont = document.getElementById('enh-externos');
    if(!cont) return;
    const recs = getExternos();
    cont.innerHTML = '';
    document.getElementById('enh-threshold').value = Math.round((state.thresholdMin || 0) / 60);
    if(recs.length === 0){
      cont.innerHTML = '<p class="muted">Não há recursos externos cadastrados.</p>';
      renderAlerts();
      return;
    }
    recs.forEach(r => {
      if(!state.externos[r.id]){
        state.externos[r.id] = { contratadoMin:0, horasDiaMin:0, dias:{seg:true,ter:true,qua:true,qui:true,sex:true,sab:false,dom:false}, ledger:[], projetos:(r.projetos||[]) };
      }
      const data = state.externos[r.id];
      const used = (data.ledger || []).reduce((sum,e) => sum + (e.minutos || 0), 0);
      const saldo = (data.contratadoMin || 0) - used;
      const card = document.createElement('div');
      card.className = 'enh-card';
      card.innerHTML = `
        <div class="enh-grid">
          <div><span class="enh-badge">Recurso</span> <b>${r.nome}</b>
            <div class="muted small">Projetos: <input class="enh-projs" data-id="${r.id}" value="${(data.projetos||[]).join(', ')}"/></div>
          </div>
          <label>Horas contratadas (HH:MM)
            <input class="enh-contr" data-id="${r.id}" value="${fmtHHMM(data.contratadoMin)}"/>
          </label>
          <label>Horas por dia (HH:MM)
            <input class="enh-dia" data-id="${r.id}" value="${fmtHHMM(data.horasDiaMin)}"/>
          </label>
          <div>
            <div class="muted small">Dias de trabalho</div>
            <div class="enh-days" data-id="${r.id}">
              ${['seg','ter','qua','qui','sex','sab','dom'].map(day => {
                const labels = {seg:'Seg',ter:'Ter',qua:'Qua',qui:'Qui',sex:'Sex',sab:'Sáb',dom:'Dom'};
                return `<label><input type="checkbox" data-day="${day}" ${data.dias[day] ? 'checked' : ''}/> ${labels[day]}</label>`;
              }).join(' ')}
            </div>
          </div>
          <div>
            <div><span class="enh-badge">Consumido</span> <b>${fmtHHMM(used)}</b></div>
            <div><span class="enh-badge">Saldo</span> <b>${fmtHHMM(saldo)}</b></div>
          </div>
        </div>
        <div style="margin-top:6px">
          <div class="muted small">Lançamentos (inclui casos atípicos e horas extras)</div>
          <div class="enh-entry">
            <input type="date" class="enh-date" data-id="${r.id}" />
            <input type="time" step="60" class="enh-hours" data-id="${r.id}" />
            <select class="enh-type" data-id="${r.id}">
              <option value="normal">Normal</option>
              <option value="extra">Extra</option>
            </select>
            <input placeholder="Projeto (opcional)" class="enh-proj-entry" data-id="${r.id}" />
            <button class="enh-add" data-id="${r.id}">Adicionar</button>
          </div>
        </div>
        <div class="muted small" style="margin-top:4px">Histórico</div>
        <table class="tbl">
          <thead><tr><th>Data</th><th>Horas</th><th>Tipo</th><th>Projeto</th><th></th></tr></thead>
          <tbody class="enh-hist" data-id="${r.id}">
            ${(data.ledger||[]).map((e,i) => `<tr><td>${e.date}</td><td>${fmtHHMM(e.minutos)}</td><td>${e.tipo}</td><td>${e.projeto||''}</td><td><button class="enh-del" data-id="${r.id}" data-index="${i}">Excluir</button></td></tr>`).join('')}
          </tbody>
        </table>`;
      cont.appendChild(card);
    });
    // Bind events
    cont.querySelectorAll('.enh-contr').forEach(inp => {
      inp.addEventListener('change', function(){
        const id = this.dataset.id;
        state.externos[id].contratadoMin = parseHHMM(this.value);
        save(); render(); renderAlerts();
      });
    });
    cont.querySelectorAll('.enh-dia').forEach(inp => {
      inp.addEventListener('change', function(){
        const id = this.dataset.id;
        state.externos[id].horasDiaMin = parseHHMM(this.value);
        save(); render(); renderAlerts();
      });
    });
    cont.querySelectorAll('.enh-days input[type=checkbox]').forEach(cb => {
      cb.addEventListener('change', function(){
        const id = this.closest('.enh-days').dataset.id;
        const day = this.dataset.day;
        state.externos[id].dias[day] = this.checked;
        save();
      });
    });
    cont.querySelectorAll('.enh-projs').forEach(inp => {
      inp.addEventListener('change', function(){
        const id = this.dataset.id;
        state.externos[id].projetos = this.value.split(',').map(s => s.trim()).filter(Boolean);
        save();
      });
    });
    cont.querySelectorAll('.enh-add').forEach(btn => {
      btn.addEventListener('click', function(){
        const id = this.dataset.id;
        const card = this.closest('.enh-card');
        const date = card.querySelector('.enh-date[data-id="'+id+'"]').value;
        const time = card.querySelector('.enh-hours[data-id="'+id+'"]').value;
        const tipo = card.querySelector('.enh-type[data-id="'+id+'"]').value;
        const proj = card.querySelector('.enh-proj-entry[data-id="'+id+'"]').value.trim();
        if(!date || !time){
          alert('Informe data e horas');
          return;
        }
        const parts = time.split(':');
        const minutes = parseInt(parts[0],10)*60 + parseInt(parts[1],10);
        state.externos[id].ledger.push({ date: date, minutos: minutes, tipo: tipo, projeto: proj || null });
        save(); render(); renderAlerts();
      });
    });
    cont.querySelectorAll('.enh-del').forEach(btn => {
      btn.addEventListener('click', function(){
        const id = this.dataset.id;
        const index = parseInt(this.dataset.index, 10);
        state.externos[id].ledger.splice(index, 1);
        save(); render(); renderAlerts();
      });
    });
    renderAlerts();
  }
  function renderAlerts(){
    const cont = document.getElementById('enh-alertas');
    if(!cont) return;
    cont.innerHTML = '';
    const recs = getExternos();
    const list = [];
    recs.forEach(r => {
      const data = state.externos[r.id];
      if(!data) return;
      const used = (data.ledger||[]).reduce((sum,e) => sum + (e.minutos||0), 0);
      const saldo = (data.contratadoMin || 0) - used;
      if(saldo <= state.thresholdMin){
        list.push({ nome: r.nome, saldo: saldo });
      }
    });
    if(list.length === 0){
      cont.innerHTML = '<p class="muted">Nenhum recurso no limite configurado.</p>';
      return;
    }
    list.sort((a,b) => a.saldo - b.saldo);
    const table = document.createElement('table');
    table.className = 'tbl';
    table.innerHTML = '<thead><tr><th>Recurso</th><th>Saldo restante</th></tr></thead>';
    const tbody = document.createElement('tbody');
    list.forEach(item => {
      const tr = document.createElement('tr');
      tr.innerHTML = '<td>'+item.nome+'</td><td>'+fmtHHMM(item.saldo)+'</td>';
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    cont.appendChild(table);
  }
  document.addEventListener('DOMContentLoaded', function(){
    load(); ensureTab();
  });
})();
