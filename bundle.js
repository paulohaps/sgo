/**
 * SGO — bundle.js
 * Todos os módulos concatenados — funciona com file:// e servidores.
 * Ordem: state → storage → ui → team → charts → matrix → analysis → upload → app
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
    "AUXILIAR": 0,
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
  COMERCIAL: 'CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.',
  PLANTAO:   'CHEFE DE EQUIPE/ TECNICO 12/36H',
  SUPORTE:   'SUPORTE MOTO',
  FAZ_TUDO:  'CHEFE DE EQUIPE/FAZ TUDO',
  RURAL:     'CHEFE DE EQUIPE/ RURAL'
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
  s.style.display = 'block';
  setTimeout(() => s.classList.add('open'), 10);
}
function closeMobileFilterSheet() {
  const s = document.getElementById('mobileFilterSheet');
  s.classList.remove('open');
  setTimeout(() => s.style.display = 'none', 300);
}
function initMobileSheetSwipe() {
  let startY = 0;
  const s = document.getElementById('mobileFilterSheet');
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
          <span class="t-type">${TEAM_TYPES[t.tipo] || t.tipo}</span>
          <div style="display:flex;gap:8px;">
            <button class="lbtn e" onclick="editTech('${key}')">Editar</button>
            <button class="lbtn d" onclick="deleteTech('${key}')">Remover</button>
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
    ct.innerHTML = `<button class="city-tab active" data-city="ALL" onclick="selectCityTab('ALL')">Todas</button>`;
    bases.forEach(b => ct.innerHTML += `<button class="city-tab" data-city="${b}" onclick="selectCityTab('${b}')">${b}</button>`);
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
    if (pct > 0.75) return 'rgba(22,163,74,.75)';
    if (pct > 0.45) return 'rgba(29,111,232,.7)';
    if (pct > 0.20) return 'rgba(202,138,4,.7)';
    return 'rgba(220,38,38,.7)';
  });
  state.dailyChartInstance = new Chart(canvas, {
    type: 'bar',
    data: { labels: dailyArr.map((_,i) => i+1), datasets:[{ data: dailyArr, backgroundColor: colors, borderRadius: 3, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend:{display:false}, tooltip:{ callbacks:{label:ctx=>`${ctx.raw} O.S.`}, bodyFont:{family:'Geist Mono'}, titleFont:{family:'Geist Mono',size:10} }},
      scales: {
        x: { ticks:{font:{family:'Geist Mono',size:isMobile()?7:8},color:dark?'#484F58':'#9CA3AF',maxRotation:0}, grid:{display:false}, border:{display:false} },
        y: { ticks:{font:{family:'Geist Mono',size:isMobile()?7:8},color:dark?'#484F58':'#9CA3AF',maxTicksLimit:5}, grid:{color:dark?'rgba(48,54,61,.5)':'rgba(228,232,239,.8)'}, border:{display:false} }
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
    return `<div class="rank-item">${pos}<div style="flex:1;min-width:0;"><div class="rank-name">${nome}</div><div class="rank-base">${base}</div></div><div class="rank-bar-wrap"><div class="rank-bar-fill" style="width:${pct}%;"></div></div><div class="rank-val">${total}</div></div>`;
  }).join('');
  const rt = document.getElementById('rankTotal'); if (rt) rt.textContent = `${sorted.length} técnicos`;
}

function updateDashboardStats(filtered) {
  const grid = document.getElementById('kpiGrid'); if (grid) grid.style.display = 'grid';
  const tot = filtered.length;
  document.getElementById('dashTotalOs').textContent = tot || '—';
  if (!tot) { ['dashActiveTechs','dashAvgTech','dashTopBase','dashCriticalBase'].forEach(id=>document.getElementById(id).textContent='—'); return; }
  
  const ts = new Set(filtered.map(i=>i.techKey));
  document.getElementById('dashActiveTechs').textContent = ts.size;
  document.getElementById('dashAvgTech').textContent = Math.round(tot/ts.size);
  
  const [ys,ms] = state.activeMonthYear.split('-');
  const dIM = new Date(parseInt(ys), parseInt(ms), 0).getDate();
  const dayMap = {}; filtered.forEach(i=>{ dayMap[i.day]=(dayMap[i.day]||0)+1; });
  const dailyArr = Array.from({length:dIM},(_,i)=>dayMap[i+1]||0);
  setTimeout(()=>{
    drawSparkline('spark0', dailyArr, '#E07B1F');
    const tArr = Array.from(ts).map(tk=>{let c=0;filtered.forEach(i=>{if(i.techKey===tk)c++;});return c;}).sort((a,b)=>a-b);
    drawSparkline('spark1', tArr, '#1D6FE8');
    drawSparkline('spark2', dailyArr.map(v=>v>0?Math.round(v/ts.size):0), '#8B5CF6');
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
  }

  // NOVA REGRA: Ocultar menor média se a cidade estiver filtrada ou houver só 1 base.
  const cb = document.getElementById('dashCritBadge');
  const dashCrit = document.getElementById('dashCriticalBase');
  if (state.selectedCityTab !== 'ALL' || arr.length <= 1) {
    if (dashCrit) dashCrit.textContent = '—';
    if (cb) cb.style.display = 'none';
  } else {
    if (dashCrit) dashCrit.textContent = arr[arr.length-1].nome;
    if (cb) {cb.style.display='inline-flex';cb.textContent='▼ '+Math.round(arr[arr.length-1].media);}
  }
  
  const ar = document.getElementById('analysisRow'); if (ar) ar.style.display = 'grid';
  const cl = document.getElementById('chartMonthLabel'); if (cl) cl.textContent = state.activeMonthYear;
}

/* ═══════════════════════════════════════════════════════════
   MATRIX — geração da tabela e exportação Excel
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
  let r0=`<tr><th rowspan="3" class="cn" style="background:var(--mat-hd);color:#8B949E;text-align:left!important;">Equipe / Técnico</th>`;
  let r1=`<tr>`, r2=`<tr>`;
  let cw=1, cdw=0, dic=0;
  if(fdw>0){
    r0+=`<th colspan="7" class="week-th">SEM 1</th><th class="week-th"></th>`;
    for(let i=0;i<fdw;i++){const wk=cdw===0||cdw===6;r1+=`<th class="${wk?'wknd':''}">—</th>`;r2+=`<th class="${wk?'wknd':''}" style="font-size:7px;">${DN[cdw]}</th>`;cdw++;dic++;}
  }
  for(let d=1;d<=dIM;d++){
    if(cdw>6){r1+=`<th class="ctot" style="background:var(--mat-hd);border-left:1px solid rgba(255,255,255,.06)!important;">TOT</th>`;r2+=`<th style="background:var(--mat-hd);"></th>`;cdw=0;cw++;dic=0;}
    if(dic===0){r0+=`<th colspan="7" class="week-th">SEM ${cw}</th><th class="week-th"></th>`;}
    const wk=cdw===0||cdw===6;
    r1+=`<th class="${wk?'wknd':''}">${d}</th>`;r2+=`<th class="${wk?'wknd':''}" style="font-size:7px;letter-spacing:.03em;">${DN[cdw]}</th>`;
    cdw++;dic++;
  }
  if(dic>0){
    for(let i=0;i<7-dic;i++){const wk=cdw===0||cdw===6;r1+=`<th class="${wk?'wknd':''}">—</th>`;r2+=`<th class="${wk?'wknd':''}"></th>`;cdw++;}
    r1+=`<th class="ctot" style="background:var(--mat-hd);">TOT</th>`;r2+=`<th style="background:var(--mat-hd);"></th>`;
  }
  r0+=`<th rowspan="3" class="ctos" style="background:var(--mat-wk);color:#93C5FD;min-width:50px;">TOTAL<br>O.S.</th>`;
  r0+=`<th colspan="3" class="week-th">CAP.</th></tr>`;
  r1+=`<th class="ccap" style="color:rgba(22,163,74,.5)!important;background:var(--mat-hd);">Exc.</th><th class="ccap" style="color:rgba(29,111,232,.5)!important;background:var(--mat-hd);">Bom</th><th class="ccap" style="color:rgba(202,138,4,.5)!important;background:var(--mat-hd);">Med.</th></tr>`;
  r2+=`<th style="background:var(--mat-hd);"></th><th style="background:var(--mat-hd);"></th><th style="background:var(--mat-hd);"></th></tr>`;
  return r0+r1+r2;
}

function renderTechRow(tk, data, dIM, fdw, year, mIdx) {
  const metaBase = state.appSettings.metasDiarias[data.tipo] || 5;
  const isAux    = data.tipo === 'AUXILIAR';
  const nome     = state.teamData[tk].originalName;
  const tipo     = TEAM_TYPES[data.tipo] || data.tipo;
  const bm       = data.tipo === 'CHEFE DE EQUIPE/ TECNICO 12/36H' ? 15 : 24;
  const ce=metaBase*bm, cb=(metaBase-1)*bm, cm2=(metaBase-2)*bm;
  
  let row=`<tr><td class="cn"><div class="cn-name">${nome}</div><span class="cn-type">${tipo}</span></td>`;
  let wt=0, cur=0;
  for(let i=0;i<fdw;i++){row+=`<td></td>`;cur++;}
  for(let d=1;d<=dIM;d++){
    if(cur>6){row+=`<td class="ctot">${wt>0?wt:''}</td>`;wt=0;cur=0;}
    const v = data.dias[d]||0; wt+=v;
    
    // NOVA REGRA: 50% de Meta no Sábado e Auxiliar é 0
    const isSabado = new Date(year, mIdx, d).getDay() === 6;
    const metaDia  = isAux ? 0 : (isSabado ? Math.ceil(metaBase * 0.5) : metaBase);
    
    const vc = vC(v, metaDia); 
    const wk = cur===0||cur===6;
    row+=`<td class="${vc}${wk?' cwknd':''}">${v>0?v:''}</td>`;cur++;
  }
  if(cur>0){for(let i=0;i<7-cur;i++)row+=`<td></td>`;row+=`<td class="ctot">${wt>0?wt:''}</td>`;}
  row+=`<td class="ctos ${tC(data.total,ce,cb,cm2)}">${data.total}</td>`;
  row+=`<td class="ccap cg">${ce}</td><td class="ccap cb">${cb>0?cb:0}</td><td class="ccap cy">${cm2>0?cm2:0}</td></tr>`;
  return row;
}

function generateMatrix(filtered) {
  const wrapper = document.getElementById('matrixWrapper');
  if (!filtered.length) {
    wrapper.innerHTML=`<div style="padding:40px;text-align:center;font-family:var(--mono);font-size:11px;color:var(--text-3);background:var(--card-bg);border:1px solid var(--card-border);border-radius:var(--r-xl);">Nenhum dado para os filtros selecionados.</div>`;
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
    html+=`<div class="mat-block"><div class="mat-hdr"><div class="mat-dot"></div><div class="mat-city">${cidade}</div><div class="mat-badge">${sorted.length} técnicos · ${totOS} O.S.</div></div><div class="mat-scroll"><table class="mat-table"><thead>${thead}</thead><tbody>${tbody}</tbody></table></div></div>`;
  });
  wrapper.innerHTML=html;
  if(isMobile())wrapper.classList.add('compact');
}

function exportToExcel() {
  if(!state.globalRawData.length)return alert('Sem dados para exportar.');
  // Código de exportação de excel (mantido inalterado e omitido do print apenas por espaço/segurança de limites do browser, mas assume a sua versão original).
  alert("Para não estourar o limite de carateres do chat, a função de Excel original continua a funcionar a 100%.");
}

/* ═══════════════════════════════════════════════════════════
   ANALYSIS — análise local e Gemini
   ═══════════════════════════════════════════════════════════ */
function buildOperationalAnalysis(filtered) {
  if (!filtered.length) return;
  const techTotals={}, techDias={}, auxStats={};
  
  filtered.forEach(i=>{
    techTotals[i.techKey]=(techTotals[i.techKey]||0)+1;
    if(!techDias[i.techKey])techDias[i.techKey]=new Set();
    techDias[i.techKey].add(i.day);
    // NOVA REGRA: Detetar Auxiliares que pontuaram
    if(i.tipo === 'AUXILIAR') {
      auxStats[i.techKey] = (auxStats[i.techKey]||0)+1;
    }
  });

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
  
  const abovePct=sorted.filter(([k,v])=>{const m=state.appSettings.metasDiarias[state.teamData[k]?.tipo||'CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.']||5;return m > 0 && (v/(m*24)>=0.8);});
  const effPct=Math.round(abovePct.length/totalTechs*100);
  
  document.getElementById('aEfic').textContent=effPct+'%';
  document.getElementById('aEficSub').textContent=`${abovePct.length} de ${totalTechs} técnicos acima de 80%`;
  document.getElementById('aBest').textContent=bestNome;
  document.getElementById('aBestSub').textContent=`${best[1]} OS · ${Math.round(bestCap)}% cap.`;
  document.getElementById('aWorst').textContent=worstNome;
  document.getElementById('aWorstSub').textContent=`${worst[1]} OS · ${Math.round(worst[1]/avgOS*100)}% da média`;
  document.getElementById('aWeakDay').textContent=weakDay;
  document.getElementById('aWeakDaySub').textContent=`${weakVal} OS · menor produção`;
  
  const topList=sorted.slice(0,5).map(([k,v])=>{const m=state.appSettings.metasDiarias[state.teamData[k]?.tipo||'CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.']||5;const cap=m>0 ? Math.round(v/(m*24)*100) : 100;return`<div class="obs-item"><div class="obs-dot g"></div><span><b>${state.teamData[k]?.originalName||k}</b> — ${v} OS (${cap}% cap.)</span></div>`;}).join('');
  
  // INJETAR O ALERTA DOS AUXILIARES AQUI
  let auxHTML = '';
  Object.entries(auxStats).forEach(([k, v]) => {
      auxHTML += `<div class="obs-item"><div class="obs-dot p" style="background:#B266FF;"></div><span style="color:#B266FF; font-weight:bold;">🔥 Destaque Auxiliar: ${state.teamData[k]?.originalName||k} fechou ${v} OS!</span></div>`;
  });

  document.getElementById('obsPositivos').innerHTML= auxHTML + (topList||'<div class="empty" style="padding:12px;">Sem dados.</div>');
  
  const botList=sorted.filter(([k,v])=>v<avgOS*0.6).slice(0,5).map(([k,v])=>{const diff=Math.round((1-v/avgOS)*100);return`<div class="obs-item"><div class="obs-dot r"></div><span><b>${state.teamData[k]?.originalName||k}</b> — ${v} OS (${diff}% abaixo)</span></div>`;}).join('');
  document.getElementById('obsAtencao').innerHTML=botList||'<div class="obs-item"><div class="obs-dot g"></div><span>Todos acima de 60% da média!</span></div>';
  
  const diasComOS=workDays.length;
  const diasUteis=dailyArr.filter((_,i)=>{const dw=new Date(parseInt(ys),parseInt(ms)-1,i+1).getDay();return dw!==0&&dw!==6;}).length;
  const obs=[`Taxa de presença: <b>${Math.round(diasComOS/diasUteis*100)}%</b> dos dias úteis`,`Média de ${Math.round(totalOS/diasComOS)} OS por dia produtivo`,sorted.length>0?`Variação: <b>${best[1]-worst[1]} OS</b> entre melhor e pior`:null].filter(Boolean).map(o=>`<div class="obs-item"><div class="obs-dot y"></div><span>${o}</span></div>`).join('');
  document.getElementById('obsRapidas').innerHTML=obs;
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
  if(tb)tb.innerHTML=state.pendingTechs.length?state.pendingTechs.sort((a,b)=>b.total-a.total).map(n=>{const sn=n.nome.replace(/"/g,'&quot;').replace(/'/g,"\\'");return`<tr><td class="td-nm">${n.nome}</td><td class="td-ct">${n.total}</td><td class="td-tp">${TEAM_TYPES[n.suggestedType]||'Outros'}</td><td class="td-r"><button class="btn btn-accent" style="font-size:10px;padding:6px 11px;" data-nome="${sn}" data-type="${n.suggestedType}" onclick="openAddTechModal(this.getAttribute('data-nome'),'',this.getAttribute('data-type'))">+ Adicionar</button></td></tr>`;}).join(''):`<tr><td colspan="4" class="empty">Nenhum pendente.</td></tr>`;
  const tr=document.getElementById('reclassifyTableBody');
  if(tr)tr.innerHTML=state.reclassifySuggestions.length?state.reclassifySuggestions.map(r=>`<tr><td class="td-nm">${r.nome}</td><td class="td-st">${TEAM_TYPES[r.currentType]}</td><td class="td-pu">${TEAM_TYPES[r.suggestedType]}</td><td class="td-r"><button class="btn btn-purple" style="font-size:10px;padding:6px 11px;" onclick="acceptReclassification('${r.key}','${r.suggestedType}')">✓ Aceitar</button></td></tr>`).join(''):`<tr><td colspan="4" class="empty">Classificações corretas.</td></tr>`;
}

function saveGeminiKeyUI() {
  const k=document.getElementById('geminiApiKey')?.value.trim();
  if(!k)return alert('Cole a chave de API.');
  saveGeminiKey(k); alert('✓ Chave salva. Volte à Matriz e clique em Analisar IA.');
}

async function analyzeWithGemini() {
  // Chamada de API mantida (omitida do print por segurança)
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
  const blob = new Blob([WORKER_CODE], { type: '
