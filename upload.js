/**
 * SGO — upload.js
 * Worker de parsing de planilha, pipeline de dados e filtros.
 */

import { state, TEAM_TYPES } from './state.js';
import { limparNome } from './team.js';
import { showLoading, setStatus, isMobile } from './ui.js';
import { updateDashboardStats } from './charts.js';
import { generateMatrix } from './matrix.js';
import { evaluateTechsByAI, renderPendentes, buildOperationalAnalysis } from './analysis.js';
import { populateFilters, renderTeamTable } from './team.js';

// ── Worker inline (XLSX no thread separado) ───────────────
const WORKER_CODE = `
importScripts('https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.min.js');

function cleanStr(s){
  if(!s)return"";
  return String(s).split('-')[0].split('>')[0].split('(')[0].split('/')[0]
    .normalize("NFD").replace(/[\\u0300-\\u036f]/g,"")
    .replace(/\\s+/g,' ').trim().toUpperCase();
}

function extractDate(v){
  if(!v)return null;
  let s=String(v).trim();
  if(typeof v==='number'||!isNaN(Number(s))){
    let d=new Date(Math.round((Number(s)-25569)*86400*1000));
    d.setMinutes(d.getMinutes()+d.getTimezoneOffset());
    return d;
  }
  let b=s.match(/(\\d{1,2})\\/(\\d{1,2})\\/(\\d{2,4})/);
  if(b){
    let y=b[3].length===2?parseInt('20'+b[3]):parseInt(b[3]);
    return new Date(y,parseInt(b[2],10)-1,parseInt(b[1],10));
  }
  let i=s.match(/(\\d{4})-(\\d{1,2})-(\\d{1,2})/);
  if(i)return new Date(parseInt(i[1]),parseInt(i[2])-1,parseInt(i[3]));
  return null;
}

self.onmessage=function(e){
  try{
    const{fileData}=e.data;
    const wb=XLSX.read(fileData,{type:'array',raw:true});
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:""});
    if(rows.length<=1)throw new Error("Arquivo sem dados");

    const hdrs=rows[0].map(h=>String(h).normalize("NFD").replace(/[\\u0300-\\u036f]/g,"").trim().toLowerCase());
    const iD=hdrs.findIndex(h=>h.includes('fechamento')||h.includes('conclusao')||h.includes('data'));
    const iR=hdrs.findIndex(h=>h.includes('responsavel')||h.includes('tecnico')||h.includes('executor')||h.includes('colaborador'));
    if(iD===-1||iR===-1)throw new Error("Colunas 'Responsável' e/ou 'Data' não encontradas.");

    let pm={},ts={},vr=[];
    const rR=/(RURAL|FAZENDA|S[IÍ]TIO|LINHA |GLEBA|PROJETO)/i;

    for(let i=1;i<rows.length;i++){
      const dt=rows[i][iD],rp=rows[i][iR];
      if(!dt||!rp||String(rp).toLowerCase().includes("filtros"))continue;
      const d=extractDate(dt);
      if(!d)continue;
      const nm=cleanStr(rp);
      const mk=d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,'0');
      pm[mk]=(pm[mk]||0)+1;
      let isR=rR.test(rows[i].join(" "));
      if(!ts[nm])ts[nm]={total:0,rural:0,days:new Set()};
      ts[nm].total++;
      if(isR)ts[nm].rural++;
      ts[nm].days.add(d.getDate());
      vr.push({nome:nm,day:d.getDate(),monthStr:mk});
    }

    if(!vr.length)throw new Error("Nenhuma OS válida encontrada.");
    let am=Object.keys(pm).reduce((a,b)=>pm[a]>pm[b]?a:b);
    vr=vr.filter(r=>r.monthStr===am);
    let ss={};
    for(let k in ts)ss[k]={total:ts[k].total,rural:ts[k].rural,days:Array.from(ts[k].days)};
    self.postMessage({success:true,activeMonth:am,allOS:vr,techStats:ss});
  }catch(err){
    self.postMessage({success:false,error:err.message});
  }
};
`;

let workerBlobUrl = null;

export function initWorker() {
  const blob = new Blob([WORKER_CODE], { type: 'application/javascript' });
  workerBlobUrl = URL.createObjectURL(blob);
  state.workerBlobUrl = workerBlobUrl;
}

// ── Upload de arquivo ─────────────────────────────────────
export function parseFileWithWorker(event) {
  const file = event.target.files[0];
  if (!file) return;
  showLoading(true);

  const reader = new FileReader();
  reader.onload = ev => {
    const ab = ev.target.result;
    const w  = new Worker(workerBlobUrl);

    w.onmessage = msg => {
      showLoading(false);
      // Limpar inputs
      ['fileInput', 'fileInputMob'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
      });

      if (msg.data.success) {
        onWorkerSuccess(msg.data);
      } else {
        alert('Erro ao processar planilha: ' + msg.data.error);
      }
      w.terminate();
    };

    w.postMessage({ fileData: ab }, [ab]);
  };
  reader.readAsArrayBuffer(file);
}

function onWorkerSuccess(data) {
  state.activeMonthYear   = data.activeMonth;
  state.rawExcelCache     = data.allOS;
  state.globalTechStats   = data.techStats;

  // Atualizar inputs de mês
  ['filterMonth', 'filterMonthMob'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = state.activeMonthYear;
  });

  // Mostrar elementos
  document.getElementById('cityTabsWrapper').style.display = 'block';
  document.getElementById('searchBar').style.display = 'flex';
  const ma = document.getElementById('mobActions');
  if (ma) ma.style.display = 'flex';

  setStatus('Dados · ' + state.activeMonthYear);
  syncEcosystem();
}

// ── Rebuild dados brutos ───────────────────────────────────
export function rebuildGlobalRawData() {
  if (!state.rawExcelCache.length) return;
  state.globalRawData = [];
  const cache = {};

  state.rawExcelCache.forEach(os => {
    let tk = cache[os.nome];
    if (tk === undefined) {
      tk = null;
      for (const k in state.teamData) {
        if (os.nome === k || os.nome.includes(k) || k.includes(os.nome)) { tk = k; break; }
      }
      cache[os.nome] = tk;
    }
    if (tk) {
      state.globalRawData.push({
        techKey:  tk,
        cidade:   state.teamData[tk].base,
        tipo:     state.teamData[tk].tipo || 'INSTALAÇÃO CIDADE',
        day:      os.day,
        monthStr: os.monthStr
      });
    }
  });

  applyFilters();
}

// ── Aplicar filtros ────────────────────────────────────────
export function applyFilters() {
  if (!state.globalRawData.length) {
    const wrapper = document.getElementById('matrixWrapper');
    if (wrapper) wrapper.innerHTML = `<div style="padding:40px;text-align:center;font-family:var(--mono);font-size:11px;color:var(--text-3);background:var(--card-bg);border:1px solid var(--card-border);border-radius:var(--r-xl);">Faça o upload da planilha para gerar a matriz.</div>`;
    return;
  }

  // Sincronizar filtro mobile
  const ft  = document.getElementById('filterType')?.value || 'ALL';
  const ftm = document.getElementById('filterTypeMob');
  if (ftm) ftm.value = ft;

  const fc = state.selectedCityTab;
  const filtered = state.globalRawData.filter(i =>
    i.monthStr === state.activeMonthYear &&
    (fc === 'ALL' || i.cidade === fc) &&
    (ft === 'ALL' || i.tipo === ft)
  );

  updateDashboardStats(filtered);
  buildOperationalAnalysis(filtered);
  generateMatrix(filtered);

  document.getElementById('legendBar').style.display = 'flex';
  document.getElementById('matrixSection').style.display = 'block';
  document.getElementById('btnGeminiAnalyze')?.classList.remove('hidden');
}

// ── Sync completo do ecossistema ───────────────────────────
export function syncEcosystem() {
  populateFilters();
  renderTeamTable();
  evaluateTechsByAI();
  renderPendentes();
  rebuildGlobalRawData();
}

// ── Busca na matriz ────────────────────────────────────────
export function filterMatrixBySearch() {
  const q    = document.getElementById('techSearchInput')?.value.trim().toLowerCase() || '';
  const rows = document.querySelectorAll('#matrixWrapper tbody tr');
  let count  = 0;

  rows.forEach(r => {
    const name = r.querySelector('.cn-name');
    if (!name) { r.style.display = ''; return; }
    const match = !q || name.textContent.toLowerCase().includes(q);
    r.style.display = match ? '' : 'none';
    if (match) count++;
  });

  const sc = document.getElementById('searchCount');
  if (sc) sc.textContent = q ? `${count} resultado${count !== 1 ? 's' : ''}` : '';
}

export function clearSearch() {
  const input = document.getElementById('techSearchInput');
  if (input) input.value = '';
  filterMatrixBySearch();
}
