const TEAM_TYPES = {
    "AUXILIAR": "Auxiliar",
    "CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.": "Chefe de equipe/ Instalação cidade.",
    "CHEFE DE EQUIPE/ RURAL": "Chefe de equipe/ Rural",
    "CHEFE DE EQUIPE/ TECNICO 12/36H": "Chefe de equipe/ tecnico 12/36H",
    "SUPORTE MOTO": "Suporte Moto",
    "CHEFE DE EQUIPE/FAZ TUDO": "Chefe de equipe/Faz tudo"
  };
  
  const state = {
    appSettings: { metasDiarias: { "CHEFE DE EQUIPE/ INSTALAÇÃO CIDADE.": 5, "CHEFE DE EQUIPE/ RURAL": 3, "CHEFE DE EQUIPE/ TECNICO 12/36H": 4, "SUPORTE MOTO": 8, "CHEFE DE EQUIPE/FAZ TUDO": 5 } },
    teamData: {}, rawExcelCache: [], globalRawData: [], globalTechStats: {},
    selectedCityTab: "ALL", activeMonthYear: ""
  };
  
  // UI Handlers (Sidebar)
  document.addEventListener('DOMContentLoaded', () => {
    const mt = document.getElementById('menu-toggle');
    if(mt) mt.addEventListener('click', () => document.getElementById('sidebar').classList.toggle('collapsed'));
    
    // Load Settings
    const s = localStorage.getItem('sgo_settings');
    if(s) state.appSettings = JSON.parse(s);
    const t = localStorage.getItem('sgo_team');
    if(t) state.teamData = JSON.parse(t);
  
    loadSettingsToUI();
  });
  
  function switchTab(tab) {
    document.querySelectorAll('.view').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.nav-tab').forEach(el => el.classList.remove('active'));
    document.getElementById(`view-${tab}`)?.classList.add('active');
    document.getElementById(`tab-${tab}`)?.classList.add('active');
  }
  
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
    localStorage.setItem('sgo_settings', JSON.stringify(state.appSettings));
    alert('Metas salvas com sucesso.');
    applyFilters();
  }
  
  // Dashboard Update & City Panorama
  function updateDashboardStats(filtered) {
    document.getElementById('kpiGrid').style.display = 'grid';
    document.getElementById('dashTotalOs').textContent = filtered.length;
    const ts = new Set(filtered.map(i => i.techKey));
    document.getElementById('dashActiveTechs').textContent = ts.size;
    document.getElementById('dashAvgTech').textContent = ts.size > 0 ? Math.round(filtered.length / ts.size) : 0;
  
    const bs = {};
    filtered.forEach(i => {
      if(!bs[i.cidade]) bs[i.cidade] = {total: 0, techs: new Set()};
      bs[i.cidade].total++; bs[i.cidade].techs.add(i.techKey);
    });
    const arr = Object.keys(bs).map(b => ({nome: b, media: bs[b].total / bs[b].techs.size})).sort((a,b) => b.media - a.media);
    
    if(arr.length > 0) {
      document.getElementById('dashTopBase').textContent = arr[0].nome;
      document.getElementById('dashTopBadge').style.display = 'inline-block';
    }
  
    // Ocultar Menor Média se for só 1 cidade
    if (state.selectedCityTab !== 'ALL' || arr.length <= 1) {
        document.getElementById('kpiCriticalBaseCard').style.display = 'none';
    } else {
        document.getElementById('kpiCriticalBaseCard').style.display = 'block';
        document.getElementById('dashCriticalBase').textContent = arr[arr.length - 1].nome;
        document.getElementById('dashCritBadge').style.display = 'inline-block';
    }
  }
  
  // Matriz Logic (Metas de Sabado 50%)
  function vC(v, m) { if (!v || v === 0) return 'vd'; if (v >= m) return 'vg'; if (v >= m - 1) return 'vb'; if (v >= m - 2) return 'vy'; return 'vr'; }
  
  function generateMatrix(filtered) {
      const wrapper = document.getElementById('matrixWrapper');
      wrapper.innerHTML = '';
      if(!filtered.length) return;
      
      const [ys, ms] = state.activeMonthYear.split('-');
      const year = parseInt(ys), mIdx = parseInt(ms) - 1;
      const dIM = new Date(year, mIdx + 1, 0).getDate();
      
      const cm = {};
      filtered.forEach(i => {
          if (!cm[i.cidade]) cm[i.cidade] = {};
          if (!cm[i.cidade][i.techKey]) cm[i.cidade][i.techKey] = { dias: {}, total: 0, tipo: i.tipo };
          cm[i.cidade][i.techKey].dias[i.day] = (cm[i.cidade][i.techKey].dias[i.day] || 0) + 1;
          cm[i.cidade][i.techKey].total++;
      });
  
      Object.keys(cm).forEach(cidade => {
          if(state.selectedCityTab !== 'ALL' && cidade !== state.selectedCityTab) return;
          const techs = cm[cidade];
          let thead = `<tr><th class="cn">Técnico</th>`;
          for(let d=1; d<=dIM; d++) thead += `<th>${d}</th>`;
          thead += `<th>Total</th></tr>`;
  
          let tbody = '';
          Object.entries(techs).forEach(([tk, data]) => {
              const baseMeta = state.appSettings.metasDiarias[data.tipo] || 5;
              const isAux = data.tipo === 'AUXILIAR';
              
              tbody += `<tr><td class="cn">${state.teamData[tk]?.originalName || tk}<br><span style="font-size:9px;color:var(--text-3);">${TEAM_TYPES[data.tipo]||data.tipo}</span></td>`;
              for(let d=1; d<=dIM; d++) {
                  const v = data.dias[d] || 0;
                  const isSaturday = new Date(year, mIdx, d).getDay() === 6;
                  
                  // Lógica Auxiliar=0, Sábado=50%
                  let targetDia = isAux ? 0 : (isSaturday ? Math.ceil(baseMeta * 0.5) : baseMeta);
                  
                  const color = vC(v, targetDia);
                  tbody += `<td class="${color}">${v > 0 ? v : ''}</td>`;
              }
              tbody += `<td>${data.total}</td></tr>`;
          });
          
          wrapper.innerHTML += `<div class="mat-block"><div class="mat-hdr"><div class="mat-city">${cidade}</div></div><div style="overflow-x:auto;"><table class="mat-table"><thead>${thead}</thead><tbody>${tbody}</tbody></table></div></div>`;
      });
  }
  
  // Analise: Identificar Auxiliares
  function buildOperationalAnalysis(filtered) {
      if(!filtered.length) return;
      // ... lógicas padrões
      const auxStats = {};
      filtered.forEach(i => {
          if(state.teamData[i.techKey]?.tipo === 'AUXILIAR') {
              auxStats[i.techKey] = (auxStats[i.techKey] || 0) + 1;
          }
      });
  
      let auxHTML = '';
      Object.entries(auxStats).forEach(([tk, total]) => {
          const nm = state.teamData[tk]?.originalName || tk;
          auxHTML += `<div style="color:var(--accent); font-weight:bold; margin-bottom: 5px;">🔥 Auxiliar Destaque: ${nm} fechou ${total} O.S!</div>`;
      });
  
      document.getElementById('obsPositivos').innerHTML = auxHTML + "<i>Outras análises geradas...</i>";
  }
  
  // Handlers Window Export
  window.App = {
      switchTab, 
      saveGlobalSettings, 
      applyFilters: () => { generateMatrix(state.globalRawData); updateDashboardStats(state.globalRawData); buildOperationalAnalysis(state.globalRawData); },
      // Funções mockadas para garantir funcionamento:
      openAddTechModal: () => document.getElementById('techModal').style.display='flex',
      closeTechModal: () => document.getElementById('techModal').style.display='none',
      saveTechForm: () => {
          const nm = document.getElementById('modTechName').value.toUpperCase();
          state.teamData[nm] = { originalName: nm, base: document.getElementById('modTechCity').value.toUpperCase(), tipo: document.getElementById('modTechType').value };
          localStorage.setItem('sgo_team', JSON.stringify(state.teamData));
          App.closeTechModal(); alert("Salvo.");
      }
  };
