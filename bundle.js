/**
 * SGO — bundle.js (COMPLETO)
 * Todos os módulos concatenados com regras de Auxiliar, Sábados (50%) e Menu Lateral.
 */

/* ═══════════════════════════════════════════════════════════
   STATE — constantes e estado global
   ═══════════════════════════════════════════════════════════ */
const TEAM_TYPES = {
  "AUXILIAR": "Auxiliar",
  "CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.": "Chefe de equipe/ Instalação cidade.",
  "CHEFE DE EQUIPE/ RURAL": "Chefe de equipe/ Rural",
  "CHEFE DE EQUIPE/ TECNICO 12/36H": "Chefe de equipe/ tecnico 12/36H",
  "SUPORTE MOTO": "Suporte Moto",
  "CHEFE DE EQUIPE/FAZ TUDO": "Chefe de equipe/Faz tudo"
};

const DEFAULT_SETTINGS = {
  metasDiarias: {
    "CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.": 5,
    "CHEFE DE EQUIPE/ RURAL": 3,
    "CHEFE DE EQUIPE/ TECNICO 12/36H": 4,
    "SUPORTE MOTO": 8,
    "CHEFE DE EQUIPE/FAZ TUDO": 5
  }
};

const DN = ['Dom','Seg','Ter','Qua','Qui','Sex','Sáb'];

const PAGE_TITLES = {
  dashboard:       'Matriz Operacional',
  analise:         'Análise Operacional',
  pendentes:       'Novos Colaboradores',
  reclassificacao: 'Reclassificar IA',
  config:          'Base de Equipes'
};

const state = {
  appSettings:           {},
  teamData:              {},
  rawExcelCache:         [],
  globalRawData:         [],
  globalTechStats:       {},
  pendingTechs:          [],
  reclassifySuggestions: [],
  activeMonthYear:       "",
  workerBlobUrl:         null,
  selectedCityTab:       "ALL",
  isCompactMode:         false,
  dailyChartInstance:    null,
  isDark:                false
};

/* ═══════════════════════════════════════════════════════════
   STORAGE — localStorage e import/export
   ═══════════════════════════════════════════════════════════ */
const STORAGE_KEYS = {
  settings: 'sgo_settings_pro_v3',
  team:     'sgo_team_pro_v3',
  gemini:   'sgo_gemini_key',
  dark:     'sgo_dark'
};

const TYPE_MIGRATION = {
  "COMERCIAL": "CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.",
  "INSTALAÇÃO CIDADE": "CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.",
  "PLANTAO": "CHEFE DE EQUIPE/ TECNICO 12/36H",
  "TECNICO 12/36H": "CHEFE DE EQUIPE/ TECNICO 12/36H",
  "SUPORTE": "SUPORTE MOTO",
  "RURAL": "CHEFE DE EQUIPE/ RURAL",
  "FAZ_TUDO": "CHEFE DE EQUIPE/FAZ TUDO",
  "FAZ TUDO": "CHEFE DE EQUIPE/FAZ TUDO"
};

function migrateTypes(teamObj) {
  let changed = false;
  Object.values(teamObj).forEach(t => {
    const nt = TYPE_MIGRATION[t.tipo];
    if (nt) { t.tipo = nt; changed = true; }
  });
  return changed;
}

function initLocalStorage() {
  const rawSettings = localStorage.getItem(STORAGE_KEYS.settings);
  if (rawSettings) {
    state.appSettings = JSON.parse(rawSettings);
  } else {
    state.appSettings = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
  }
  const rawTeam = localStorage.getItem(STORAGE_KEYS.team) || localStorage.getItem('sgo_team_pro_v2');
  state.teamData = rawTeam ? JSON.parse(rawTeam) : {};
  if (migrateTypes(state.teamData)) saveTeam();
}

function saveSettings() { localStorage.setItem(STORAGE_KEYS.settings, JSON.stringify(state.appSettings)); }
function saveTeam()     { localStorage.setItem(STORAGE_KEYS.team,     JSON.stringify(state.teamData)); }
function saveGeminiKey(k) { localStorage.setItem(STORAGE_KEYS.gemini, k); }
function loadGeminiKey()  { return localStorage.getItem(STORAGE_KEYS.gemini) || ''; }
function saveDarkPref(v)  { localStorage.setItem(STORAGE_KEYS.dark, v ? '1' : '0'); }
function loadDarkPref()   { return localStorage.getItem(STORAGE_KEYS.dark) === '1'; }

function exportTeamData() {
  if (!Object.keys(state.teamData).length) return alert('Sem equipes cadastradas.');
  const a = document.createElement('a');
  a.href = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(state.teamData, null, 2));
  a.download = 'equipes_sgo.json';
  a.click();
}

function importTeamData(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const data = JSON.parse(ev.target.result);
      migrateTypes(data);
      state.teamData = data;
      saveTeam();
      alert('Importado com sucesso!');
      syncEcosystem();
    } catch { alert('Arquivo JSON inválido.'); }
  };
  reader.readAsText(file);
}

/* ═══════════════════════════════════════════════════════════
   UI — navegação, tema, relógio, sheet mobile
   ═══════════════════════════════════════════════════════════ */
function isMobile() { return window.innerWidth <= 768; }

function initClock() {
  const tick = () => {
    const el = document.getElementById('clockEl');
    if (el) el.textContent = new Date().toLocaleTimeString('pt-BR', { hour:'2-digit', minute:'2-digit', second:'2-digit' });
  };
  setInterval(tick, 1000); tick();
}

function applyDarkTheme(isDark) {
  document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light');
  ['darkBtn','darkBtnMob'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = isDark ? '☀️' : '🌙';
  });
}

function toggleDark() {
  state.isDark = !state.isDark;
  applyDarkTheme(state.isDark);
  saveDarkPref(state.isDark);
  if (state.dailyChartInstance) { state.dailyChartInstance.destroy(); state.dailyChartInstance = null; }
  if (state.globalRawData.length) applyFilters();
}

function switchTab(tab) {
  ['dashboard','analise','pendentes','reclassificacao','config'].forEach(t => {
    document.getElementById('view-' + t)?.classList.remove('active');
    document.getElementById('tab-' + t)?.classList.remove('active');
    document.getElementById('mob-tab-' + t)?.classList.remove('active');
  });
  document.getElementById('view-' + tab)?.classList.add('active');
  document.getElementById('tab-' + tab)?.classList.add('active');
  document.getElementById('mob-tab-' + tab)?.classList.add('active');
  const pt = document.getElementById('mobPageTitle');
  if (pt) pt.textContent = PAGE_TITLES[tab] || tab;
  if (isMobile()) window.scrollTo({ top: 0, behavior: 'smooth' });
}

function openMobileFilterSheet() {
  const s = document.getElementById('mobileFilterSheet');
  if(s){ s.style.display = 'block'; setTimeout(() => s.classList.add('open'), 10); }
}
function closeMobileFilterSheet() {
  const s = document.getElementById('mobileFilterSheet');
  if(s){ s.classList.remove('open'); setTimeout(() => s.style.display = 'none', 300); }
}
function initMobileSheetSwipe() {
  let startY = 0;
  const s = document.getElementById('mobileFilterSheet');
  if(!s) return;
  s.addEventListener('touchstart', e => { startY = e.touches[0].clientY; }, { passive: true });
  s.addEventListener('touchmove',  e => { if (e.touches[0].clientY - startY > 80) closeMobileFilterSheet(); }, { passive: true });
}

function toggleCompactMode() {
  state.isCompactMode = !state.isCompactMode;
  document.getElementById('matrixWrapper')?.classList.toggle('compact', state.isCompactMode);
  const lbl = state.isCompactMode ? 'Padrão' : 'Compacto';
  const btn = document.getElementById('btnCompact'); if (btn) btn.textContent = lbl;
  const mbl = document.getElementById('mobCompactLbl'); if (mbl) mbl.textContent = lbl;
}

function showLoading(v) { document.getElementById('loadingOverlay').style.display = v ? 'flex' : 'none'; }

function setStatus(text) {
  const s = document.getElementById('statusText'); if (s) s.textContent = text;
  const m = document.getElementById('mobStatusText'); if (m) m.textContent = text;
}

function syncMobileFilters() {
  const ft = document.getElementById('filterType');
  const fm = document.getElementById('filterTypeMob');
  if (fm && ft) ft.value = fm.value;
}

/* ═══════════════════════════════════════════════════════════
   TEAM — CRUD de colaboradores
   ═══════════════════════════════════════════════════════════ */
function limparNome(n) {
  if (!n) return '';
  return String(n).split('-')[0].split('>')[0].split('(')[0].split('/')[0]
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
}

function openAddTechModal(name, city, tipo) {
  name = name || ''; city = city || ''; tipo = tipo || 'CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.';
  document.getElementById('modalTitle').textContent = 'Adicionar Colaborador';
  document.getElementById('modTechName').value = name;
  document.getElementById('modTechName').disabled = false;
  document.getElementById('modTechCity').value = city;
  document.getElementById('modTechType').value = tipo;
  document.getElementById('techModal').style.display = 'flex';
}

function editTech(key) {
  const t = state.teamData[key]; if (!t) return;
  document.getElementById('modalTitle').textContent = 'Editar Colaborador';
  document.getElementById('modTechName').value = t.originalName;
  document.getElementById('modTechName').disabled = true;
  document.getElementById('modTechCity').value = t.base;
  document.getElementById('modTechType').value = t.tipo || 'CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.';
  document.getElementById('techModal').style.display = 'flex';
}

function closeTechModal() { document.getElementById('techModal').style.display = 'none'; }

function saveTechForm() {
  const name = document.getElementById('modTechName').value.trim().toUpperCase();
  const base = document.getElementById('modTechCity').value.trim().toUpperCase() || 'BASE NÃO DEFINIDA';
  const tipo = document.getElementById('modTechType').value;
  if (!name) return alert('Nome obrigatório.');
  state.teamData[limparNome(name)] = { originalName: name, base, tipo };
  saveTeam(); closeTechModal(); syncEcosystem();
}

function deleteTech(key) {
  if (!confirm('Remover colaborador?')) return;
  delete state.teamData[key]; saveTeam(); syncEcosystem();
}

function acceptReclassification(key, tipo) {
  if (!state.teamData[key]) return;
  state.teamData[key].tipo = tipo; saveTeam();
  alert('Reclassificado: ' + (TEAM_TYPES[tipo] || tipo));
  syncEcosystem();
}

function renderTeamTable() {
  const ft  = limparNome(document.getElementById('searchTech')?.value || '');
  const fc  = document.getElementById('filterTeamCity')?.value || 'ALL';
  const grp = {};
  Object.values(state.teamData).forEach(t => {
    if (ft && !limparNome(t.originalName).includes(ft)) return;
    if (fc !== 'ALL' && t.base !== fc) return;
    if (!grp[t.base]) grp[t.base] = [];
    grp[t.base].push(t);
  });
  const container = document.getElementById('teamListContainer');
  if (!container) return;
  if (!Object.keys(grp).length) { container.innerHTML = '<div class="empty">Nenhum colaborador encontrado.</div>'; return; }
  container.innerHTML = Object.keys(grp).sort().map(base => {
    const members = grp[base].sort((a,b) => a.originalName.localeCompare(b.originalName));
    return `<div class="team-grp">
      <div class="tg-hdr"><span class="tg-name">${base}</span><span class="tg-cnt">${members.length} membros</span></div>
      ${members.map(t => {
        const key = limparNome(t.originalName).replace(/'/g,"\\'");
        return `<div class="t-row">
          <span class="t-name">${t.originalName}</span>
          <span class="t-type" style="color:var(--accent); font-size:10px;">${TEAM_TYPES[t.tipo] || t.tipo}</span>
          <div style="display:flex;gap:8px;">
            <button class="lbtn e" style="color:var(--blue);" onclick="App.editTech('${key}')">Editar</button>
            <button class="lbtn d" style="color:var(--red);" onclick="App.deleteTech('${key}')">Remover</button>
          </div>
        </div>`;
      }).join('')}
    </div>`;
  }).join('');
}

function populateFilters() {
  const bases = [...new Set(Object.values(state.teamData).map(t => t.base))].sort();
  const ct = document.getElementById('cityTabsContainer');
  if (ct) {
    ct.innerHTML = `<button class="city-tab active" data-city="ALL" onclick="App.selectCityTab('ALL')">Todas</button>`;
    bases.forEach(b => ct.innerHTML += `<button class="city-tab" data-city="${b}" onclick="App.selectCityTab('${b}')">${b}</button>`);
    if (state.selectedCityTab !== 'ALL' && !bases.includes(state.selectedCityTab)) state.selectedCityTab = 'ALL';
    document.querySelectorAll('.city-tab').forEach(t => t.classList.toggle('active', t.dataset.city === state.selectedCityTab));
  }
  const fc = document.getElementById('filterTeamCity');
  if (fc) {
    const prev = fc.value;
    fc.innerHTML = '<option value="ALL">Todas as Regionais</option>';
    bases.forEach(b => fc.innerHTML += `<option value="${b}">${b}</option>`);
    fc.value = bases.includes(prev) ? prev : 'ALL';
  }
  const dl = document.getElementById('regionaisDatalist');
  if (dl) { dl.innerHTML = ''; bases.forEach(b => dl.innerHTML += `<option value="${b}">`); }
}

function selectCityTab(city) {
  state.selectedCityTab = city;
  document.querySelectorAll('.city-tab').forEach(t => t.classList.toggle('active', t.dataset.city === city));
  applyFilters();
}

/* ═══════════════════════════════════════════════════════════
   CHARTS — sparklines, gráfico diário, ranking, KPIs
   ═══════════════════════════════════════════════════════════ */
function drawSparkline(id, data, color) {
  const c = document.getElementById(id); if (!c) return;
  const ctx = c.getContext('2d');
  const W = c.offsetWidth || 60, H = c.offsetHeight || 26;
  c.width = W * 2; c.height = H * 2; ctx.scale(2, 2);
  ctx.clearRect(0, 0, W, H);
  if (!data || data.length < 2) return;
  const mn = Math.min(...data), mx = Math.max(...data), rng = mx - mn || 1, pad = 3;
  const pts = data.map((v, i) => ({ x: pad + (i/(data.length-1))*(W-pad*2), y: H-pad-(v-mn)/rng*(H-pad*2) }));
  const grad = ctx.createLinearGradient(0,0,0,H);
  grad.addColorStop(0, color+'44'); grad.addColorStop(1, color+'00');
  ctx.beginPath(); ctx.moveTo(pts[0].x, pts[0].y);
  pts.slice(1).forEach(p => ctx.lineTo(p.x, p.y));
  ctx.lineTo(pts[pts.length-1].x, H); ctx.lineTo(pts[0].x, H); ctx.closePath();
  ctx.fillStyle = grad; ctx.fill();
  ctx.beginPath(); ctx.moveTo(pts[0].x, pts[0].y);
  pts.slice(1).forEach(p => ctx.lineTo(p.x, p.y));
  ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.stroke();
}

function drawDailyChart(dailyArr) {
  const canvas = document.getElementById('dailyChart'); if (!canvas) return;
  if (state.dailyChartInstance) { state.dailyChartInstance.destroy(); state.dailyChartInstance = null; }
  const dark = state.isDark, maxVal = Math.max(...dailyArr);
  const colors = dailyArr.map(v => {
    if (v === 0) return dark ? 'rgba(48,54,61,.4)' : 'rgba(228,232,239,.6)';
    const pct = v / maxVal;
    if (pct > 0.75) return 'rgba(0,255,170,.75)';
    if (pct > 0.45) return 'rgba(0,136,255,.7)';
    if (pct > 0.20) return 'rgba(255,213,0,.7)';
    return 'rgba(255,68,68,.7)';
  });
  state.dailyChartInstance = new Chart(canvas, {
    type: 'bar',
    data: { labels: dailyArr.map((_,i) => i+1), datasets:[{ data: dailyArr, backgroundColor: colors, borderRadius: 3, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend:{display:false}, tooltip:{ callbacks:{label:ctx=>`${ctx.raw} O.S.`}, bodyFont:{family:'Geist Mono'}, titleFont:{family:'Geist Mono',size:10} }},
      scales: {
        x: { ticks:{font:{family:'Geist Mono',size:isMobile()?7:8},color:dark?'#8A9BB5':'#9CA3AF',maxRotation:0}, grid:{display:false}, border:{display:false} },
        y: { ticks:{font:{family:'Geist Mono',size:isMobile()?7:8},color:dark?'#8A9BB5':'#9CA3AF',maxTicksLimit:5}, grid:{color:dark?'rgba(0,255,170,.15)':'rgba(228,232,239,.8)'}, border:{display:false} }
      }
    }
  });
}

function buildRanking(filtered) {
  const totals = {};
  filtered.forEach(i => { totals[i.techKey] = (totals[i.techKey]||0)+1; });
  const sorted = Object.entries(totals).sort((a,b)=>b[1]-a[1]).slice(0,8);
  if (!sorted.length) return;
  const maxVal = sorted[0][1];
  const medals = ['🥇','🥈','🥉'], posClass = ['gold','silver','bronze'];
  const rb = document.getElementById('rankBody'); if (!rb) return;
  rb.innerHTML = sorted.map(([tk,total],i) => {
    const nome = state.teamData[tk]?.originalName || tk;
    const base = state.teamData[tk]?.base || '';
    const pct  = Math.round(total/maxVal*100);
    const pos  = i<3 ? `<span class="rank-pos ${posClass[i]}">${medals[i]}</span>` : `<span class="rank-pos">${i+1}</span>`;
    return `<div class="rank-item" style="display:flex; align-items:center; gap:10px; padding:10px 0; border-bottom:1px solid var(--card-border);">${pos}<div style="flex:1;min-width:0;"><div class="rank-name" style="font-weight:bold;">${nome}</div><div class="rank-base" style="font-size:10px;color:var(--text-3);">${base}</div></div><div class="rank-bar-wrap" style="width:70px;height:6px;background:var(--card-bg2);border-radius:10px;overflow:hidden;"><div class="rank-bar-fill" style="width:${pct}%;background:var(--accent);height:100%;"></div></div><div class="rank-val" style="font-weight:bold;">${total}</div></div>`;
  }).join('');
  const rt = document.getElementById('rankTotal'); if (rt) rt.textContent = `${sorted.length} técnicos`;
}

function updateDashboardStats(filtered) {
  const grid = document.getElementById('kpiGrid'); if (grid) grid.style.display = 'grid';
  const tot = filtered.length;
  document.getElementById('dashTotalOs').textContent = tot || '—';
  
  if (!tot) { ['dashActiveTechs','dashAvgTech','dashTopBase','dashCriticalBase'].forEach(id=> { if(document.getElementById(id)) document.getElementById(id).textContent='—'}); return; }
  
  const ts = new Set(filtered.map(i=>i.techKey));
  document.getElementById('dashActiveTechs').textContent = ts.size;
  document.getElementById('dashAvgTech').textContent = Math.round(tot/ts.size);
  
  const [ys,ms] = state.activeMonthYear.split('-');
  const dIM = new Date(parseInt(ys), parseInt(ms), 0).getDate();
  const dayMap = {}; filtered.forEach(i=>{ dayMap[i.day]=(dayMap[i.day]||0)+1; });
  const dailyArr = Array.from({length:dIM},(_,i)=>dayMap[i+1]||0);
  
  setTimeout(()=>{
    drawSparkline('spark0', dailyArr, '#00FFAA');
    const tArr = Array.from(ts).map(tk=>{let c=0;filtered.forEach(i=>{if(i.techKey===tk)c++;});return c;}).sort((a,b)=>a-b);
    drawSparkline('spark1', tArr, '#0088FF');
    drawSparkline('spark2', dailyArr.map(v=>v>0?Math.round(v/ts.size):0), '#B266FF');
    drawDailyChart(dailyArr);
    buildRanking(filtered);
  }, 100);

  const bs = {};
  filtered.forEach(i=>{ if(!bs[i.cidade])bs[i.cidade]={total:0,techs:new Set()}; bs[i.cidade].total++; bs[i.cidade].techs.add(i.techKey); });
  const arr = Object.keys(bs).map(b=>({nome:b,media:bs[b].total/bs[b].techs.size})).sort((a,b)=>b.media-a.media);
  
  if (arr.length) {
    document.getElementById('dashTopBase').textContent = arr[0].nome;
    const tb = document.getElementById('dashTopBadge');
    if(tb){tb.style.display='inline-flex';tb.textContent='▲ '+Math.round(arr[0].media);}
    
    // Ocultar Menor Média se filtrado por apenas 1 cidade
    const critCard = document.getElementById('kpiCriticalBaseCard') || document.getElementById('dashCriticalBase')?.closest('.kpi');
    if (state.selectedCityTab !== 'ALL' || arr.length <= 1) {
        if(critCard) critCard.style.display = 'none';
    } else {
        if(critCard) critCard.style.display = 'block';
        document.getElementById('dashCriticalBase').textContent = arr[arr.length-1].nome;
        const cb = document.getElementById('dashCritBadge');
        if(cb){cb.style.display='inline-flex';cb.textContent='▼ '+Math.round(arr[arr.length-1].media);}
    }
  }

  const ar = document.getElementById('analysisRow'); if (ar) ar.style.display = 'grid';
  const cl = document.getElementById('chartMonthLabel'); if (cl) cl.textContent = state.activeMonthYear;
}

/* ═══════════════════════════════════════════════════════════
   MATRIX — geração da tabela e exportação Excel (Regra Sábado)
   ═══════════════════════════════════════════════════════════ */
function vC(v, m) {
  if (!v || v===0) return 'vd';
  if (v>=m)     return 'vg';
  if (v>=m-1)   return 'vb';
  if (v>=m-2)   return 'vy';
  return 'vr';
}
function tC(v, ce, cb, cm) {
  if (v>=ce) return 'tg';
  if (v>=cb) return 'tb';
  if (v>=cm) return 'ty';
  return 'tr';
}

function buildCityMap(filtered) {
  const cm = {};
  filtered.forEach(i=>{
    if(!cm[i.cidade])cm[i.cidade]={};
    if(!cm[i.cidade][i.techKey])cm[i.cidade][i.techKey]={dias:{},total:0,tipo:i.tipo};
    cm[i.cidade][i.techKey].dias[i.day]=(cm[i.cidade][i.techKey].dias[i.day]||0)+1;
    cm[i.cidade][i.techKey].total++;
  });
  return cm;
}

function renderHeaderHTML(dIM, fdw) {
  let r0=`<tr><th rowspan="3" class="cn" style="background:var(--mat-hd);color:var(--accent);text-align:left!important;">Equipe / Técnico</th>`;
  let r1=`<tr>`, r2=`<tr>`;
  let cw=1, cdw=0, dic=0;
  if(fdw>0){
    r0+=`<th colspan="7" class="week-th" style="background:var(--mat-wk);color:#fff;">SEM 1</th><th class="week-th"></th>`;
    for(let i=0;i<fdw;i++){const wk=cdw===0||cdw===6;r1+=`<th class="${wk?'wknd':''}">—</th>`;r2+=`<th class="${wk?'wknd':''}" style="font-size:7px;">${DN[cdw]}</th>`;cdw++;dic++;}
  }
  for(let d=1;d<=dIM;d++){
    if(cdw>6){r1+=`<th class="ctot" style="background:var(--mat-hd);border-left:1px solid rgba(255,255,255,.06)!important;">TOT</th>`;r2+=`<th style="background:var(--mat-hd);"></th>`;cdw=0;cw++;dic=0;}
    if(dic===0){r0+=`<th colspan="7" class="week-th" style="background:var(--mat-wk);color:#fff;">SEM ${cw}</th><th class="week-th"></th>`;}
    const wk=cdw===0||cdw===6;
    r1+=`<th class="${wk?'wknd':''}">${d}</th>`;r2+=`<th class="${wk?'wknd':''}" style="font-size:7px;letter-spacing:.03em;">${DN[cdw]}</th>`;
    cdw++;dic++;
  }
  if(dic>0){
    for(let i=0;i<7-dic;i++){const wk=cdw===0||cdw===6;r1+=`<th class="${wk?'wknd':''}">—</th>`;r2+=`<th class="${wk?'wknd':''}"></th>`;cdw++;}
    r1+=`<th class="ctot" style="background:var(--mat-hd);">TOT</th>`;r2+=`<th style="background:var(--mat-hd);"></th>`;
  }
  r0+=`<th rowspan="3" class="ctos" style="background:var(--mat-wk);color:var(--accent);min-width:50px;">TOTAL<br>O.S.</th>`;
  r0+=`<th colspan="3" class="week-th" style="background:var(--mat-wk);color:#fff;">CAP.</th></tr>`;
  r1+=`<th class="ccap" style="color:var(--green)!important;background:var(--mat-hd);">Exc.</th><th class="ccap" style="color:var(--blue)!important;background:var(--mat-hd);">Bom</th><th class="ccap" style="color:var(--yellow)!important;background:var(--mat-hd);">Med.</th></tr>`;
  r2+=`<th style="background:var(--mat-hd);"></th><th style="background:var(--mat-hd);"></th><th style="background:var(--mat-hd);"></th></tr>`;
  return r0+r1+r2;
}

function renderTechRow(tk, data, dIM, fdw, year, mIdx) {
  const metaBase = state.appSettings.metasDiarias[data.tipo] || 5;
  const isAuxiliar = data.tipo === 'AUXILIAR';
  const nome = state.teamData[tk].originalName;
  const tipo = TEAM_TYPES[data.tipo] || data.tipo;
  
  // Calcular capacidade total do mês dinamicamente (respeitando sábado 50%)
  let ce = 0;
  for(let d=1; d<=dIM; d++) {
     let dw = new Date(year, mIdx, d).getDay();
     if(dw !== 0) { // Menos domingo
        let isSat = dw === 6;
        ce += isAuxiliar ? 0 : (isSat ? Math.ceil(metaBase*0.5) : metaBase);
     }
  }
  if (data.tipo === 'CHEFE DE EQUIPE/ TECNICO 12/36H') ce = metaBase * 15;
  const cb = Math.floor(ce * 0.85);
  const cm2 = Math.floor(ce * 0.70);

  let row=`<tr><td class="cn"><div class="cn-name" style="font-weight:bold;">${nome}</div><span class="cn-type" style="font-size:9px;color:var(--text-3);display:block;">${tipo}</span></td>`;
  let wt=0, cur=0;
  for(let i=0;i<fdw;i++){row+=`<td></td>`;cur++;}
  for(let d=1;d<=dIM;d++){
    if(cur>6){row+=`<td class="ctot" style="color:var(--text-3); font-size:10px;">${wt>0?wt:''}</td>`;wt=0;cur=0;}
    const v = data.dias[d]||0; wt+=v;
    
    // REGRA DE SÁBADO 50% & AUXILIAR 0
    const isSaturday = new Date(year, mIdx, d).getDay() === 6;
    let metaDia = isAuxiliar ? 0 : (isSaturday ? Math.ceil(metaBase * 0.5) : metaBase);

    const vc=vC(v, metaDia), wk=cur===0||cur===6;
    row+=`<td class="${vc}${wk?' cwknd':''}">${v>0?v:''}</td>`;cur++;
  }
  if(cur>0){for(let i=0;i<7-cur;i++)row+=`<td></td>`;row+=`<td class="ctot" style="color:var(--text-3); font-size:10px;">${wt>0?wt:''}</td>`;}
  row+=`<td class="ctos ${tC(data.total,ce,cb,cm2)}" style="font-size:13px; font-weight:bold;">${data.total}</td>`;
  row+=`<td class="ccap" style="color:var(--green); font-size:10px;">${ce}</td><td class="ccap" style="color:var(--blue); font-size:10px;">${cb>0?cb:0}</td><td class="ccap" style="color:var(--yellow); font-size:10px;">${cm2>0?cm2:0}</td></tr>`;
  return row;
}

function generateMatrix(filtered) {
  const wrapper = document.getElementById('matrixWrapper');
  if (!filtered.length) {
    wrapper.innerHTML=`<div style="padding:40px;text-align:center;font-family:var(--mono);font-size:12px;color:var(--text-3);background:var(--card-bg);border:1px solid var(--card-border);border-radius:var(--r-xl);">Nenhum dado para os filtros selecionados.</div>`;
    return;
  }
  const [ys,ms]=state.activeMonthYear.split('-');
  const year=parseInt(ys), mIdx=parseInt(ms)-1;
  const dIM=new Date(year,mIdx+1,0).getDate(), fdw=new Date(year,mIdx,1).getDay();
  
  const cityMap=buildCityMap(filtered);
  let html='';
  Object.keys(cityMap).sort().forEach(cidade=>{
    if(state.selectedCityTab!=='ALL'&&cidade!==state.selectedCityTab)return;
    const techs=cityMap[cidade];
    const sorted=Object.entries(techs).sort((a,b)=>b[1].total-a[1].total);
    const totOS=Object.values(techs).reduce((s,t)=>s+t.total,0);
    
    const thead=renderHeaderHTML(dIM,fdw);
    const tbody=sorted.map(([tk,data])=>renderTechRow(tk,data,dIM,fdw,year,mIdx)).join('');
    
    html+=`<div class="mat-block"><div class="mat-hdr"><div class="mat-city">${cidade}</div><div style="font-size:10px; color:var(--text-3); border:1px solid var(--card-border); padding:3px 8px; border-radius:15px; margin-left:10px;">${sorted.length} técnicos · ${totOS} O.S.</div></div><div style="overflow-x:auto;"><table class="mat-table"><thead>${thead}</thead><tbody>${tbody}</tbody></table></div></div>`;
  });
  wrapper.innerHTML=html;
  if(isMobile())wrapper.classList.add('compact');
}

function exportToExcel() {
  if(!state.globalRawData.length)return alert('Sem dados para exportar.');
  const [ys,ms]=state.activeMonthYear.split('-');
  const year=parseInt(ys),mIdx=parseInt(ms)-1;
  const dIM=new Date(year,mIdx+1,0).getDate(),fdw=new Date(year,mIdx,1).getDay();
  const fc=state.selectedCityTab, ft=document.getElementById('filterType')?.value||'ALL';
  const filtered=state.globalRawData.filter(i=>i.monthStr===state.activeMonthYear&&(fc==='ALL'||i.cidade===fc)&&(ft==='ALL'||i.tipo===ft));
  if(!filtered.length)return alert('Sem dados nos filtros selecionados.');
  const cityMap=buildCityMap(filtered);
  const wb=XLSX.utils.book_new();
  const XL={navy:'FF000B29',navyMid:'FF001240',navyLight:'FF001A8A',navyTxt:'FFC4E0E5',weekTxt:'FF00FFAA',white:'FFFFFFFF',border:'FF0022FF',greenBg:'FF00FFAA',greenFg:'FF000000',blueBg:'FF0088FF',blueFg:'FFFFFFFF',yellowBg:'FFFFD500',yellowFg:'FF000000',redBg:'FFFF4444',redFg:'FFFFFFFF',totBg:'FF001A8A',totFg:'FF8A9BB5'};
  const cs=(fill,font,bold=false,align='center',sz=9)=>({fill:fill?{fgColor:{rgb:fill}}:{},font:{color:{rgb:font||XL.navy},bold,name:'Calibri',sz},border:{top:{style:'thin',color:{rgb:XL.border}},bottom:{style:'thin',color:{rgb:XL.border}},left:{style:'thin',color:{rgb:XL.border}},right:{style:'thin',color:{rgb:XL.border}}},alignment:{horizontal:align,vertical:'center',wrapText:false}});
  
  Object.keys(cityMap).sort().forEach(cidade=>{
    const techs=cityMap[cidade];
    const sorted=Object.entries(techs).sort((a,b)=>b[1].total-a[1].total);
    let wsData=[],merges=[],r0=['Técnico'],r1=[''],r2=[''];
    let cw=1,cdw=0,dic=0,ci=1;
    if(fdw>0){r0.push('SEM 1');merges.push({s:{r:0,c:ci},e:{r:0,c:ci+6}});for(let i=0;i<6;i++)r0.push('');r0.push('');for(let i=0;i<fdw;i++){r1.push('—');r2.push(DN[cdw]);cdw++;dic++;ci++;}}
    for(let d=1;d<=dIM;d++){
      if(cdw>6){r1.push('TOT');r2.push('');ci++;cdw=0;cw++;dic=0;}
      if(dic===0){r0.push(`SEM ${cw}`);merges.push({s:{r:0,c:ci},e:{r:0,c:ci+6}});for(let i=0;i<6;i++)r0.push('');r0.push('');}
      r1.push(d);r2.push(DN[cdw]);cdw++;dic++;ci++;
    }
    if(dic>0){let rem=7-dic;for(let i=0;i<rem;i++){r1.push('—');r2.push(DN[cdw]);cdw++;ci++;}r1.push('TOT');r2.push('');ci++;}
    r0.push('TOTAL');merges.push({s:{r:0,c:ci},e:{r:2,c:ci}});
    r0.push('CAPACIDADE');merges.push({s:{r:0,c:ci+1},e:{r:0,c:ci+3}});r0.push('');r0.push('');
    r1.push('');r1.push('Excelente');r1.push('Bom');r1.push('Mediano');
    r2.push('');r2.push('');r2.push('');r2.push('');
    wsData.push(r0);wsData.push(r1);wsData.push(r2);
    
    let mnl=18;
    sorted.forEach(([tk,data])=>{
      const meta = state.appSettings.metasDiarias[data.tipo]||5;
      const nd = state.teamData[tk].originalName;if(nd.length>mnl)mnl=nd.length;
      let row=[nd],wt=0,cur=0;
      for(let i=0;i<fdw;i++){row.push('');cur++;}
      for(let d=1;d<=dIM;d++){if(cur>6){row.push(wt>0?wt:'');wt=0;cur=0;}const v=data.dias[d]||0;wt+=v;row.push(v>0?v:'');cur++;}
      if(cur>0){for(let i=0;i<7-cur;i++)row.push('');row.push(wt>0?wt:'');}
      
      let ce=0;
      for(let d=1; d<=dIM; d++) {
         let dw = new Date(year, mIdx, d).getDay();
         if(dw !== 0) {
            let isSat = dw === 6;
            ce += (data.tipo === 'AUXILIAR') ? 0 : (isSat ? Math.ceil(meta*0.5) : meta);
         }
      }
      if (data.tipo === 'CHEFE DE EQUIPE/ TECNICO 12/36H') ce = meta * 15;
      const cb = Math.floor(ce * 0.85);
      const cm2 = Math.floor(ce * 0.70);
      
      row.push(data.total);row.push(ce);row.push(cb>0?cb:0);row.push(cm2>0?cm2:0);
      wsData.push(row);
    });
    
    const ws=XLSX.utils.aoa_to_sheet(wsData);ws['!merges']=merges;
    let cols=[{wch:mnl+4}];
    for(let i=1;i<r0.length;i++){if(r0[i]==='TOTAL')cols.push({wch:9});else if(['Excelente','Bom','Mediano'].includes(r1[i]))cols.push({wch:10});else if(r1[i]==='TOT')cols.push({wch:5});else cols.push({wch:4.5});}
    ws['!cols']=cols;ws['!rows']=[{hpt:22},{hpt:17},{hpt:13}];
    const range=XLSX.utils.decode_range(ws['!ref']);const tci=r0.indexOf('TOTAL');
    
    for(let R=range.s.r;R<=range.e.r;++R){for(let C=range.s.c;C<=range.e.c;++C){
      const ca=XLSX.utils.encode_cell({r:R,c:C});
      if(!ws[ca])ws[ca]={v:'',t:'s'};if(!ws[ca].s)ws[ca].s={};
      const isTot=r1[C]==='TOT',isCap=['Excelente','Bom','Mediano'].includes(r1[C]);
      if(R===0){ws[ca].s=cs(r0[C]?.startsWith('SEM')?XL.navyLight:XL.navyMid,r0[C]?.startsWith('SEM')?XL.weekTxt:XL.navyTxt,true,'center',9);}
      else if(R===1){if(r1[C]==='Excelente')ws[ca].s=cs(XL.greenBg,XL.greenFg,true,'center',10);else if(r1[C]==='Bom')ws[ca].s=cs(XL.blueBg,XL.blueFg,true,'center',10);else if(r1[C]==='Mediano')ws[ca].s=cs(XL.yellowBg,XL.yellowFg,true,'center',10);else ws[ca].s=cs(XL.navyMid,XL.navyTxt,true,'center',9);}
      else if(R===2){ws[ca].s=cs(XL.navyMid,'FF5A6478',false,'center',8);}
      else if(C===0){ws[ca].s=cs(XL.white,XL.navy,true,'left',10);}
      else if(C===tci&&R>2){const vt=Number(ws[ca].v||0),vce=Number(ws[XLSX.utils.encode_cell({r:R,c:tci+1})].v||0),vcb=Number(ws[XLSX.utils.encode_cell({r:R,c:tci+2})].v||0),vcm=Number(ws[XLSX.utils.encode_cell({r:R,c:tci+3})].v||0);let bg=XL.redBg,fg=XL.redFg;if(vt>=vce){bg=XL.greenBg;fg=XL.greenFg;}else if(vt>=vcb){bg=XL.blueBg;fg=XL.blueFg;}else if(vt>=vcm){bg=XL.yellowBg;fg=XL.yellowFg;}ws[ca].s=cs(bg,fg,true,'center',11);}
      else if(isCap){ws[ca].s=cs(XL.white,'FF9CA3AF',false,'center',10);}
      else if(isTot){ws[ca].s=cs(XL.totBg,XL.totFg,true,'center',9);}
      else if(R>2){
          const v=Number(ws[ca].v||0),rowIdx=R-3;
          if(rowIdx<sorted.length){
            const tipoDado = sorted[rowIdx][1].tipo;
            const meta=state.appSettings.metasDiarias[tipoDado]||5;
            let isAuxiliar = tipoDado === 'AUXILIAR';
            let metaDia = meta;
            if (parseInt(r1[C])) {
                let dayNum = parseInt(r1[C]);
                let isSaturday = new Date(year, mIdx, dayNum).getDay() === 6;
                metaDia = isAuxiliar ? 0 : (isSaturday ? Math.ceil(meta*0.5) : meta);
            } else if (isAuxiliar) {
                metaDia = 0;
            }
            let bg=XL.white,fg='FF8A9BB5',bold=false;
            if(v>0){
                if(v>=metaDia){bg=XL.greenBg;fg=XL.greenFg;bold=true;}
                else if(v>=metaDia-1){bg=XL.blueBg;fg=XL.blueFg;bold=true;}
                else if(v>=metaDia-2){bg=XL.yellowBg;fg=XL.yellowFg;bold=true;}
                else{bg=XL.redBg;fg=XL.redFg;bold=true;}
            }
            ws[ca].s=cs(bg,fg,bold,'center',9);
          }
      }
    }}
    const rowH=ws['!rows']||[];for(let R=3;R<=range.e.r;R++)rowH.push({hpt:15});
    ws['!rows']=rowH;ws['!freeze']={xSplit:1,ySplit:3};
    XLSX.utils.book_append_sheet(wb,ws,cidade.substring(0,31));
  });
  XLSX.writeFile(wb,`SGO_Matriz_${state.activeMonthYear}.xlsx`);
}

/* ═══════════════════════════════════════════════════════════
   ANALYSIS — análise local e Gemini
   ═══════════════════════════════════════════════════════════ */
function buildOperationalAnalysis(filtered) {
  if (!filtered.length) return;
  const techTotals={}, techDias={};
  filtered.forEach(i=>{techTotals[i.techKey]=(techTotals[i.techKey]||0)+1;if(!techDias[i.techKey])techDias[i.techKey]=new Set();techDias[i.techKey].add(i.day);});
  const sorted=Object.entries(techTotals).sort((a,b)=>b[1]-a[1]);
  const totalOS=filtered.length, totalTechs=sorted.length, avgOS=totalOS/totalTechs;
  const best=sorted[0], worst=sorted[sorted.length-1];
  const bestNome=state.teamData[best[0]]?.originalName||best[0];
  const worstNome=state.teamData[worst[0]]?.originalName||worst[0];
  const bestMeta=state.appSettings.metasDiarias[state.teamData[best[0]]?.tipo||'CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.']||5;
  const bestCap=best[1]/(bestMeta*24)*100;
  
  const [ys,ms]=state.activeMonthYear.split('-');
  const dIM=new Date(parseInt(ys),parseInt(ms),0).getDate();
  const dayMap={};filtered.forEach(i=>{dayMap[i.day]=(dayMap[i.day]||0)+1;});
  const dailyArr=Array.from({length:dIM},(_,i)=>dayMap[i+1]||0);
  const workDays=dailyArr.filter(v=>v>0);
  const weakIdx=dailyArr.indexOf(Math.min(...workDays.filter(v=>v>0)));
  const weakDay=weakIdx>=0?`Dia ${weakIdx+1} (${DN[new Date(parseInt(ys),parseInt(ms)-1,weakIdx+1).getDay()]})`:'—';
  const weakVal=weakIdx>=0?dailyArr[weakIdx]:0;
  
  const abovePct=sorted.filter(([k,v])=>{const m=state.appSettings.metasDiarias[state.teamData[k]?.tipo||'CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.']||5;return v/(m*24)>=0.8;});
  const effPct=Math.round(abovePct.length/totalTechs*100);
  
  document.getElementById('aEfic').textContent=effPct+'%';
  document.getElementById('aEficSub').textContent=`${abovePct.length} de ${totalTechs} técnicos acima de 80%`;
  document.getElementById('aBest').textContent=bestNome;
  document.getElementById('aBestSub').textContent=`${best[1]} OS · ${Math.round(bestCap)}% cap.`;
  document.getElementById('aWorst').textContent=worstNome;
  document.getElementById('aWorstSub').textContent=`${worst[1]} OS · ${Math.round(worst[1]/avgOS*100)}% da média`;
  document.getElementById('aWeakDay').textContent=weakDay;
  document.getElementById('aWeakDaySub').textContent=`${weakVal} OS · menor produção`;
  
  // Positivos (Top e Auxiliares)
  const topList=sorted.slice(0,5).map(([k,v])=>{const m=state.appSettings.metasDiarias[state.teamData[k]?.tipo||'CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.']||5;const cap=Math.round(v/(m*24)*100);return`<div class="obs-item" style="display:flex; align-items:center; gap:8px; margin-bottom:5px;"><div class="obs-dot" style="width:8px;height:8px;border-radius:50%;background:var(--green);"></div><span><b>${state.teamData[k]?.originalName||k}</b> — ${v} OS (${cap}% cap.)</span></div>`;}).join('');
  
  const auxStats = {};
  filtered.forEach(i => {
      if(state.teamData[i.techKey]?.tipo === 'AUXILIAR') {
          auxStats[i.techKey] = (auxStats[i.techKey] || 0) + 1;
      }
  });
  let auxHTML = '';
  Object.entries(auxStats).forEach(([tk, total]) => {
      const nm = state.teamData[tk]?.originalName || tk;
      auxHTML += `<div class="obs-item" style="display:flex; align-items:center; gap:8px; margin-bottom:5px;"><div class="obs-dot" style="width:8px;height:8px;border-radius:50%;background:var(--accent);"></div><span style="color:var(--accent); font-weight:bold;">🔥 Auxiliar Destaque: ${nm} fechou ${total} O.S!</span></div>`;
  });

  document.getElementById('obsPositivos').innerHTML = (auxHTML + topList) || '<div class="empty" style="padding:12px;">Sem dados.</div>';
  
  const botList=sorted.filter(([k,v])=>v<avgOS*0.6).slice(0,5).map(([k,v])=>{const diff=Math.round((1-v/avgOS)*100);return`<div class="obs-item" style="display:flex; align-items:center; gap:8px; margin-bottom:5px;"><div class="obs-dot" style="width:8px;height:8px;border-radius:50%;background:var(--red);"></div><span><b>${state.teamData[k]?.originalName||k}</b> — ${v} OS (${diff}% abaixo)</span></div>`;}).join('');
  document.getElementById('obsAtencao').innerHTML=botList||'<div class="obs-item" style="display:flex; align-items:center; gap:8px; margin-bottom:5px;"><div class="obs-dot" style="width:8px;height:8px;border-radius:50%;background:var(--green);"></div><span>Todos acima de 60% da média!</span></div>';
  
  const diasComOS=workDays.length;
  const diasUteis=dailyArr.filter((_,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return dw!==0&&dw!==6;}).length;
  const obs=[`Taxa de presença: <b>${Math.round(diasComOS/diasUteis*100)}%</b> dos dias úteis`,`Média de ${Math.round(totalOS/diasComOS)} OS por dia produtivo`,sorted.length>0?`Variação: <b>${best[1]-worst[1]} OS</b> entre melhor e pior`:null].filter(Boolean).map(o=>`<div class="obs-item" style="display:flex; align-items:center; gap:8px; margin-bottom:5px;"><div class="obs-dot" style="width:8px;height:8px;border-radius:50%;background:var(--yellow);"></div><span>${o}</span></div>`).join('');
  document.getElementById('obsRapidas').innerHTML=obs;
  
  const sug=[];
  if(worst[1]<avgOS*0.5)sug.push(`Investigar <b>${worstNome}</b> — abaixo de 50% da média`);
  if(effPct<50)sug.push('Revisar metas — menos de 50% atingiu 80% da capacidade');
  if(weakVal<avgOS*0.3)sug.push(`Analisar dia fraco (${weakDay})`);
  sug.push('Considerar redistribuição entre polos com desempenhos desiguais');
  document.getElementById('obsSugestoes').innerHTML=sug.slice(0,3).map(s=>`<div class="obs-item" style="display:flex; align-items:center; gap:8px; margin-bottom:5px;"><div class="obs-dot" style="width:8px;height:8px;border-radius:50%;background:var(--yellow);"></div><span>${s}</span></div>`).join('');
  
  document.getElementById('badgeAnalise')?.classList.remove('hidden');
  document.getElementById('mobBadgeAnalise')?.classList.remove('hidden');
}

function evaluateTechsByAI() {
  if(!Object.keys(state.globalTechStats).length)return;
  state.pendingTechs=[]; state.reclassifySuggestions=[];
  const cache={};
  const matchTeam=nome=>{if(cache[nome]!==undefined)return cache[nome];for(const k in state.teamData){if(nome===k||nome.includes(k)||k.includes(nome)){cache[nome]=k;return k;}}cache[nome]=null;return null;};
  
  Object.keys(state.globalTechStats).forEach(nome=>{
    const s=state.globalTechStats[nome];
    const pR=s.rural/s.total, da=[...s.days].sort((a,b)=>a-b), dt=da.length;
    let d2=0;for(let k=1;k<da.length;k++)if(da[k]-da[k-1]===2)d2++;
    const isP=dt>=4&&d2/(dt-1||1)>=0.5;
    
    let sug='CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.';
    if(isP)sug='CHEFE DE EQUIPE/ TECNICO 12/36H';else if(pR>=0.6)sug='CHEFE DE EQUIPE/ RURAL';else if(s.total/dt>7.5)sug='SUPORTE MOTO';
    
    const kc=matchTeam(nome);
    if(!kc&&s.total>1)state.pendingTechs.push({nome,total:s.total,suggestedType:sug});
    else if(kc&&s.total>=5){const ct=state.teamData[kc].tipo||'CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.';if(ct!==sug)state.reclassifySuggestions.push({key:kc,nome:state.teamData[kc].originalName,currentType:ct,suggestedType:sug});}
  });
  renderPendentes();
}

function renderPendentes() {
  const badge=(id,n)=>{const el=document.getElementById(id);if(!el)return;if(n>0){el.textContent=n;el.classList.remove('hidden');}else el.classList.add('hidden');};
  badge('badgePendentes',state.pendingTechs.length);badge('mobBadgePendentes',state.pendingTechs.length);
  badge('badgeReclassificacao',state.reclassifySuggestions.length);badge('mobBadgeReclass',state.reclassifySuggestions.length);
  const tb=document.getElementById('pendentesTableBody');
  if(tb)tb.innerHTML=state.pendingTechs.length?state.pendingTechs.sort((a,b)=>b.total-a.total).map(n=>{const sn=n.nome.replace(/"/g,'&quot;').replace(/'/g,"\\'");return`<tr><td style="font-weight:bold;">${n.nome}</td><td style="color:var(--blue); font-weight:bold;">${n.total}</td><td style="color:var(--accent); font-size:10px;">${TEAM_TYPES[n.suggestedType]||'Outros'}</td><td class="td-r"><button class="btn btn-accent" style="font-size:10px;padding:6px 11px;" data-nome="${sn}" data-type="${n.suggestedType}" onclick="App.openAddTechModal(this.getAttribute('data-nome'),'',this.getAttribute('data-type'))">+ Adicionar</button></td></tr>`;}).join(''):`<tr><td colspan="4" class="empty">Nenhum pendente.</td></tr>`;
  const tr=document.getElementById('reclassifyTableBody');
  if(tr)tr.innerHTML=state.reclassifySuggestions.length?state.reclassifySuggestions.map(r=>`<tr><td style="font-weight:bold;">${r.nome}</td><td style="font-size:10px; color:var(--text-3); text-decoration:line-through;">${TEAM_TYPES[r.currentType]}</td><td style="color:var(--purple); font-size:10px;">${TEAM_TYPES[r.suggestedType]}</td><td class="td-r"><button class="btn btn-purple" style="font-size:10px;padding:6px 11px;background:var(--purple-bg);border:1px solid var(--purple);color:var(--purple);" onclick="App.acceptReclassification('${r.key}','${r.suggestedType}')">✓ Aceitar</button></td></tr>`).join(''):`<tr><td colspan="4" class="empty">Classificações corretas.</td></tr>`;
}

function saveGeminiKeyUI() {
  const k=document.getElementById('geminiApiKey')?.value.trim();
  if(!k)return alert('Cole a chave de API.');
  saveGeminiKey(k); alert('✓ Chave salva. Volte à Matriz e clique em Analisar IA.');
}

function buildGeminiContext() {
  if(!state.globalRawData.length)return null;
  const fc=state.selectedCityTab, ft=document.getElementById('filterType')?.value||'ALL';
  const filtered=state.globalRawData.filter(i=>i.monthStr===state.activeMonthYear&&(fc==='ALL'||i.cidade===fc)&&(ft==='ALL'||i.tipo===ft));
  if(!filtered.length)return null;
  const cs={};
  filtered.forEach(i=>{if(!cs[i.cidade])cs[i.cidade]={total:0,techs:{}};cs[i.cidade].total++;if(!cs[i.cidade].techs[i.techKey])cs[i.cidade].techs[i.techKey]={total:0,tipo:i.tipo,dias:new Set()};cs[i.cidade].techs[i.techKey].total++;cs[i.cidade].techs[i.techKey].dias.add(i.day);});
  const m=state.appSettings.metasDiarias;
  let ctx=`Período: ${state.activeMonthYear}\nTotal O.S.: ${filtered.length}\nMetas: Inst.Cidade=${m["CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE."]}, Plantão=${m["CHEFE DE EQUIPE/ TECNICO 12/36H"]}, Suporte=${m["SUPORTE MOTO"]}, Rural=${m["CHEFE DE EQUIPE/ RURAL"]}, FazTudo=${m["CHEFE DE EQUIPE/FAZ TUDO"]}\n\n`;
  Object.keys(cs).sort().forEach(cidade=>{const c=cs[cidade],tl=Object.entries(c.techs);ctx+=`=== ${cidade} ===\nTotal: ${c.total} | Técnicos: ${tl.length} | Média: ${(c.total/tl.length).toFixed(1)}\n`;tl.sort((a,b)=>b[1].total-a[1].total).forEach(([k,d])=>{const meta=state.appSettings.metasDiarias[d.tipo]||5,bm=d.tipo==='CHEFE DE EQUIPE/ TECNICO 12/36H'?15:24,ce=meta*bm,dt=d.dias.size;ctx+=`  • ${state.teamData[k]?.originalName||k} (${TEAM_TYPES[d.tipo]||d.tipo}): ${d.total} OS / ${dt} dias | média/dia: ${dt>0?(d.total/dt).toFixed(1):0} | meta: ${meta}/dia | cap: ${Math.round(d.total/ce*100)}%\n`;});ctx+='\n';});
  return ctx;
}

async function analyzeWithGemini() {
  const ak=loadGeminiKey();
  if(!ak){alert('Configure a API Key na aba Base de Equipes.');switchTab('config');return;}
  const ctx=buildGeminiContext();
  if(!ctx){alert('Carregue uma planilha primeiro.');return;}
  const btn=document.getElementById('btnGeminiAnalyze');
  const panel=document.getElementById('geminiPanel'), content=document.getElementById('geminiInsightContent');
  if(btn){btn.disabled=true;btn.textContent='⏳ Analisando...';}
  panel.style.display='block';
  content.innerHTML=`<div style="display:flex;align-items:center;gap:10px;color:var(--purple);font-family:var(--mono);font-size:11px;">Processando com Gemini 2.5 Flash...</div>`;
  panel.scrollIntoView({behavior:'smooth',block:'start'});
  const prompt=`Você é um analista de produtividade de operações técnicas de provedor de internet. Analise os dados e gere um relatório objetivo. Responda em português do Brasil.\n\nRetorne apenas:\n\n1️⃣ Ranking de produtividade (Top 5 técnicos com OS e % capacidade)\n\n2️⃣ TOP 5 Técnicos abaixo da média (nome, OS, quanto abaixo)\n\n3️⃣ Observações rápidas (máx 3 pontos)\n\n4️⃣ Sugestões de melhoria (máx 5 ações concretas)\n\nUse **negrito** para nomes. Seja direto e objetivo.\n\nDADOS:\n${ctx}`;
  try{
    const res=await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${ak}`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({contents:[{parts:[{text:prompt}]}],generationConfig:{temperature:0.3,maxOutputTokens:50000}})});
    if(!res.ok){const e=await res.json();throw new Error(e?.error?.message||`HTTP ${res.status}`);}
    const data=await res.json();
    const text=data?.candidates?.[0]?.content?.parts?.[0]?.text;
    if(!text)throw new Error('Resposta vazia.');
    const fmt=text.replace(/\*\*(.*?)\*\*/g,'<strong>$1</strong>').replace(/\*(.*?)\*/g,'<em style="color:var(--text-2)">$1</em>').split('\n\n').filter(p=>p.trim()).map(p=>{if(p.includes('\n• ')||p.includes('\n- ')){return p.split('\n').map(l=>{if(l.startsWith('• ')||l.startsWith('- '))return`<div class="gb" style="display:flex; gap:8px;"><span class="gb-dot" style="color:var(--purple);">›</span><span>${l.slice(2)}</span></div>`;return`<p style="margin-bottom:3px;">${l}</p>`;}).join('');}return`<p style="margin-bottom:10px;">${p.replace(/\n/g,'<br>')}</p>`;}).join('');
    content.innerHTML=fmt;
  }catch(err){
    content.innerHTML=`<div style="color:var(--red);font-family:var(--mono);font-size:11px;"><strong style="display:block;margin-bottom:5px;">Erro ao chamar Gemini</strong>${err.message}</div>`;
  }finally{
    if(btn){btn.disabled=false;btn.textContent=`Analisar IA`;}
  }
}

/* ═══════════════════════════════════════════════════════════
   UPLOAD — worker de parsing e pipeline de dados
   ═══════════════════════════════════════════════════════════ */
const WORKER_CODE = `
importScripts('https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.min.js');
function cleanStr(s){if(!s)return"";return String(s).split('-')[0].split('>')[0].split('(')[0].split('/')[0].normalize("NFD").replace(/[\\u0300-\\u036f]/g,"").replace(/\\s+/g,' ').trim().toUpperCase();}
function extractDate(v){if(!v)return null;let s=String(v).trim();if(typeof v==='number'||!isNaN(Number(s))){let d=new Date(Math.round((Number(s)-25569)*86400*1000));d.setMinutes(d.getMinutes()+d.getTimezoneOffset());return d;}let b=s.match(/(\\d{1,2})\\/(\\d{1,2})\\/(\\d{2,4})/);if(b){let y=b[3].length===2?parseInt('20'+b[3]):parseInt(b[3]);return new Date(y,parseInt(b[2],10)-1,parseInt(b[1],10));}let i=s.match(/(\\d{4})-(\\d{1,2})-(\\d{1,2})/);if(i)return new Date(parseInt(i[1]),parseInt(i[2])-1,parseInt(i[3]));return null;}
self.onmessage=function(e){try{const{fileData}=e.data;const wb=XLSX.read(fileData,{type:'array',raw:true});const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:""});if(rows.length<=1)throw new Error("Arquivo sem dados");const hdrs=rows[0].map(h=>String(h).normalize("NFD").replace(/[\\u0300-\\u036f]/g,"").trim().toLowerCase());const iD=hdrs.findIndex(h=>h.includes('fechamento')||h.includes('conclusao')||h.includes('data'));const iR=hdrs.findIndex(h=>h.includes('responsavel')||h.includes('tecnico')||h.includes('executor')||h.includes('colaborador'));if(iD===-1||iR===-1)throw new Error("Colunas 'Responsável' e/ou 'Data' não encontradas.");let pm={},ts={},vr=[];const rR=/(RURAL|FAZENDA|S[IÍ]TIO|LINHA |GLEBA|PROJETO)/i;for(let i=1;i<rows.length;i++){const dt=rows[i][iD],rp=rows[i][iR];if(!dt||!rp||String(rp).toLowerCase().includes("filtros"))continue;const d=extractDate(dt);if(!d)continue;const nm=cleanStr(rp);const mk=d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,'0');pm[mk]=(pm[mk]||0)+1;let isR=rR.test(rows[i].join(" "));if(!ts[nm])ts[nm]={total:0,rural:0,days:new Set()};ts[nm].total++;if(isR)ts[nm].rural++;ts[nm].days.add(d.getDate());vr.push({nome:nm,day:d.getDate(),monthStr:mk});}if(!vr.length)throw new Error("Nenhuma OS válida.");let am=Object.keys(pm).reduce((a,b)=>pm[a]>pm[b]?a:b);vr=vr.filter(r=>r.monthStr===am);let ss={};for(let k in ts)ss[k]={total:ts[k].total,rural:ts[k].rural,days:Array.from(ts[k].days)};self.postMessage({success:true,activeMonth:am,allOS:vr,techStats:ss});}catch(err){self.postMessage({success:false,error:err.message});}};
`;

function initWorker() {
  const blob = new Blob([WORKER_CODE], { type: 'application/javascript' });
  state.workerBlobUrl = URL.createObjectURL(blob);
}

function parseFileWithWorker(event) {
  const file = event.target.files[0]; if (!file) return;
  showLoading(true);
  const reader = new FileReader();
  reader.onload = ev => {
    const ab = ev.target.result;
    const w  = new Worker(state.workerBlobUrl);
    w.onmessage = msg => {
      showLoading(false);
      ['fileInput','fileInputMob'].forEach(id=>{ const el=document.getElementById(id); if(el)el.value=''; });
      if (msg.data.success) {
        state.activeMonthYear = msg.data.activeMonth;
        state.rawExcelCache   = msg.data.allOS;
        state.globalTechStats = msg.data.techStats;
        ['filterMonth','filterMonthMob'].forEach(id=>{ const el=document.getElementById(id); if(el)el.value=state.activeMonthYear; });
        document.getElementById('cityTabsWrapper').style.display='block';
        document.getElementById('searchBar').style.display='flex';
        const ma=document.getElementById('mobActions'); if(ma)ma.style.display='flex';
        setStatus('Dados · '+state.activeMonthYear);
        syncEcosystem();
      } else { alert('Erro ao processar planilha: '+msg.data.error); }
      w.terminate();
    };
    w.postMessage({ fileData: ab }, [ab]);
  };
  reader.readAsArrayBuffer(file);
}

function rebuildGlobalRawData() {
  if (!state.rawExcelCache.length) return;
  state.globalRawData = [];
  const cache = {};
  state.rawExcelCache.forEach(os => {
    let tk = cache[os.nome];
    if (tk === undefined) {
      tk = null;
      for (const k in state.teamData) { if (os.nome===k||os.nome.includes(k)||k.includes(os.nome)){tk=k;break;} }
      cache[os.nome] = tk;
    }
    if (tk) state.globalRawData.push({ techKey:tk, cidade:state.teamData[tk].base, tipo:state.teamData[tk].tipo||'CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.', day:os.day, monthStr:os.monthStr });
  });
  applyFilters();
}

function applyFilters() {
  if (!state.globalRawData.length) {
    const w=document.getElementById('matrixWrapper');
    if(w)w.innerHTML=`<div style="padding:40px;text-align:center;font-family:var(--mono);font-size:11px;color:var(--text-3);background:var(--card-bg);border:1px solid var(--card-border);border-radius:var(--r-xl);">Faça o upload da planilha para gerar a matriz.</div>`;
    return;
  }
  const ft=document.getElementById('filterType')?.value||'ALL';
  const fm=document.getElementById('filterTypeMob'); if(fm)fm.value=ft;
  const fc=state.selectedCityTab;
  const filtered=state.globalRawData.filter(i=>i.monthStr===state.activeMonthYear&&(fc==='ALL'||i.cidade===fc)&&(ft==='ALL'||i.tipo===ft));
  updateDashboardStats(filtered);
  buildOperationalAnalysis(filtered);
  generateMatrix(filtered);
  document.getElementById('legendBar').style.display='flex';
  document.getElementById('matrixSection').style.display='block';
  document.getElementById('btnGeminiAnalyze')?.classList.remove('hidden');
}

function syncEcosystem() {
  populateFilters();
  renderTeamTable();
  evaluateTechsByAI();
  rebuildGlobalRawData();
}

function filterMatrixBySearch() {
  const q=document.getElementById('techSearchInput')?.value.trim().toLowerCase()||'';
  const rows=document.querySelectorAll('#matrixWrapper tbody tr');
  let count=0;
  rows.forEach(r=>{ const name=r.querySelector('.cn-name'); if(!name){r.style.display='';return;} const match=!q||name.textContent.toLowerCase().includes(q); r.style.display=match?'':'none'; if(match)count++; });
  const sc=document.getElementById('searchCount'); if(sc)sc.textContent=q?`${count} resultado${count!==1?'s':''}`:'';
}

function clearSearch() { const i=document.getElementById('techSearchInput'); if(i)i.value=''; filterMatrixBySearch(); }

/* ═══════════════════════════════════════════════════════════
   SETTINGS UI E INICIALIZAÇÃO APP
   ═══════════════════════════════════════════════════════════ */
function loadSettingsToUI() {
  const m = state.appSettings.metasDiarias;
  if(document.getElementById('cfgMetaDiaInstalacao')) document.getElementById('cfgMetaDiaInstalacao').value = m["CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE."] || 5;
  if(document.getElementById('cfgMetaDiaRural')) document.getElementById('cfgMetaDiaRural').value = m["CHEFE DE EQUIPE/ RURAL"] || 3;
  if(document.getElementById('cfgMetaDiaTecnico')) document.getElementById('cfgMetaDiaTecnico').value = m["CHEFE DE EQUIPE/ TECNICO 12/36H"] || 4;
  if(document.getElementById('cfgMetaDiaSuporte')) document.getElementById('cfgMetaDiaSuporte').value = m["SUPORTE MOTO"] || 8;
  if(document.getElementById('cfgMetaDiaFazTudo')) document.getElementById('cfgMetaDiaFazTudo').value = m["CHEFE DE EQUIPE/FAZ TUDO"] || 5;
}

function saveGlobalSettings() {
  const m = state.appSettings.metasDiarias;
  m["CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE."] = parseInt(document.getElementById('cfgMetaDiaInstalacao').value) || 5;
  m["CHEFE DE EQUIPE/ RURAL"] = parseInt(document.getElementById('cfgMetaDiaRural').value) || 3;
  m["CHEFE DE EQUIPE/ TECNICO 12/36H"] = parseInt(document.getElementById('cfgMetaDiaTecnico').value) || 4;
  m["SUPORTE MOTO"] = parseInt(document.getElementById('cfgMetaDiaSuporte').value) || 8;
  m["CHEFE DE EQUIPE/FAZ TUDO"] = parseInt(document.getElementById('cfgMetaDiaFazTudo').value) || 5;
  saveSettings();
  if (state.globalRawData.length) applyFilters();
  alert('Metas salvas com sucesso.');
}

window.App = {
  switchTab,
  toggleDark: () => toggleDark(() => { if (state.globalRawData.length) applyFilters(); }),
  openMobileFilterSheet,
  closeMobileFilterSheet,
  syncMobileFilters,
  toggleCompactMode,
  parseFileWithWorker,
  applyFilters,
  filterMatrixBySearch,
  clearSearch,
  exportToExcel,
  selectCityTab: city => { selectCityTab(city); applyFilters(); },
  openAddTechModal,
  editTech,
  closeTechModal,
  saveTechForm:           () => saveTechForm(syncEcosystem),
  deleteTech:             key => deleteTech(key, syncEcosystem),
  acceptReclassification: (key, tipo) => acceptReclassification(key, tipo, syncEcosystem),
  filterTeamTable:        renderTeamTable,
  exportTeamData,
  importTeamData:         event => importTeamData(event, syncEcosystem),
  saveGlobalSettings,
  saveGeminiKey:    saveGeminiKeyUI,
  analyzeWithGemini,
};

document.addEventListener('DOMContentLoaded', () => {
  state.isDark = loadDarkPref();
  if (state.isDark) applyDarkTheme(true);
  initLocalStorage();
  initClock();
  initWorker();
  initMobileSheetSwipe();
  loadSettingsToUI();
  renderTeamTable();
  populateFilters();
  
  const geminiInput = document.getElementById('geminiApiKey');
  if (geminiInput) geminiInput.value = loadGeminiKey();
  
  const pm = document.getElementById('pageMetaDate');
  if (pm) pm.textContent = new Date().toLocaleDateString('pt-BR', { weekday:'long', day:'2-digit', month:'long', year:'numeric' });
  if (isMobile()) state.isCompactMode = true;

  // Lógica de Toggle do Sidebar
  const mt = document.getElementById('menu-toggle');
  if(mt) mt.addEventListener('click', () => document.getElementById('sidebar').classList.toggle('collapsed'));
});
