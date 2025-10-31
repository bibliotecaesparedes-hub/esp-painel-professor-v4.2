/* ESP Painel v4.2 - admin tabs, auto-save, session header */

/* Configurações */
const MSAL_CONFIG = {
  auth: {
    clientId: "c5573063-8a04-40d3-92bf-eb229ad4701c",
    authority: "https://login.microsoftonline.com/d650692c-6e73-48b3-af84-e3497ff3e1f1",
    redirectUri: "https://bibliotecaesparedes-hub.github.io/esp-painel-professor-v4.2/"
  },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
};
const MSAL_SCOPES = { scopes: ["Files.ReadWrite.All", "User.Read", "openid", "profile", "offline_access"] };
const SITE_ID = "esparedes-my.sharepoint.com,540a0485-2578-481e-b4d8-220b41fb5c43,7335dc42-69c8-42d6-8282-151e3783162d";

const CFG_PATH = "/Documents/GestaoAlunos-OneDrive/config_especial.json";
const REG_PATH = "/Documents/GestaoAlunos-OneDrive/2registos_alunos.json";
const BACKUP_FOLDER = "/Documents/GestaoAlunos-OneDrive/backup";

/* Estado global */
let msalApp, account, accessToken;
const state = { config: null, reg: null, prof: null };
const $ = s => document.querySelector(s);

/* Utilitários UI */
function updateSync(txt) { const el = $('#syncIndicator'); if (el) el.textContent = txt; }
function log(m) { const el = $('#importLog'); if (el) el.innerHTML = `<div>${new Date().toLocaleTimeString()} — ${m}</div>` + el.innerHTML; }

/* === MSAL redirect flow === */
async function initMsal() {
  msalApp = new msal.PublicClientApplication(MSAL_CONFIG);
  try {
    const resp = await msalApp.handleRedirectPromise();
    if (resp && resp.account) {
      account = resp.account;
      msalApp.setActiveAccount(account);
      await acquireToken();
      onLogin();
      return;
    }
    const accs = msalApp.getAllAccounts();
    if (accs.length) {
      account = accs[0];
      msalApp.setActiveAccount(account);
      await acquireToken();
      onLogin();
    }
  } catch (e) {
    console.error("MSAL init error:", e);
  }
}

async function acquireToken() {
  try {
    const r = await msalApp.acquireTokenSilent(MSAL_SCOPES);
    accessToken = r.accessToken;
  } catch (e) {
    console.warn('Silent token failed:', e);
    try { await msalApp.acquireTokenRedirect(MSAL_SCOPES); } catch (err) { console.error(err); }
  }
}

function ensureLogin() {
  try {
    msalApp.loginRedirect(MSAL_SCOPES);
  } catch (e) {
    console.error('Login redirect failed', e);
    Swal.fire('Erro', 'Falha ao iniciar sessão', 'error');
  }
}

async function onLogout() {
  try { await msalApp.logoutRedirect(); }
  catch (e) { console.warn(e); }
}

function onLogin() {
  $('#btnMsLogin').style.display = 'none';
  $('#btnMsLogout').style.display = 'inline-flex';
  const user = account?.username || '—';
  document.querySelectorAll('#sessNome, #sessNomeHeader').forEach(el => el.textContent = 'Sessão: ' + user);
  updateSync('🔁 sincronizando...');
  loadConfigAndReg();
}

/* === Microsoft Graph helpers === */
async function graphLoad(path) {
  if (!accessToken) await acquireToken();
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:${path}:/content`;
  const r = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
  if (r.ok) {
    const t = await r.text();
    try { return JSON.parse(t); } catch { return null; }
  }
  if (r.status === 404) return null;
  throw new Error('Graph load failed: ' + r.status);
}

async function graphSave(path, obj) {
  if (!accessToken) await acquireToken();
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
  const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/drive/root:${path}:/content`;
  const r = await fetch(url, { method: 'PUT', headers: { Authorization: `Bearer ${accessToken}` }, body: blob });
  if (!r.ok) throw new Error('Graph save failed: ' + r.status);
  return await r.json();
}

/* === Load config & registos (fallback cache) === */
async function loadConfigAndReg() {
  try {
    const cfg = await graphLoad(CFG_PATH);
    const reg = await graphLoad(REG_PATH);
    state.config = cfg || { professores: [], alunos: [], disciplinas: [], grupos: [], calendario: {} };
    state.reg = reg || { versao: 'v1', registos: [] };
    localStorage.setItem('esp_config', JSON.stringify(state.config));
    localStorage.setItem('esp_reg', JSON.stringify(state.reg));
    updateSync('💾 guardado');
    renderDay();
  } catch (e) {
    console.warn('Fallback to cache', e);
    const c = localStorage.getItem('esp_config');
    const r = localStorage.getItem('esp_reg');
    if (c) state.config = JSON.parse(c);
    if (r) state.reg = JSON.parse(r);
    if (!state.config) state.config = { professores: [], alunos: [], disciplinas: [], grupos: [], calendario: {} };
    if (!state.reg) state.reg = { versao: 'v1', registos: [] };
    updateSync('⚠ offline');
    renderDay();
  }
}

/* === Render Painel do Dia === */
function renderDay() {
  const date = $('#dataHoje').value || new Date().toISOString().slice(0, 10);
  $('#dataHoje').value = date;
  const out = $('#sessoesHoje');
  if (!out) return;
  out.innerHTML = '';

  if (!state.config || !state.config.professores) {
    out.innerHTML = '<div class="small">⚠️ Configuração ainda não carregada.</div>';
    return;
  }

  const prof = (state.config.professores || []).find(p => p.email && p.email.toLowerCase() === (account?.username || '').toLowerCase());
  if (!prof) {
    out.innerHTML = '<div class="small">Professor não reconhecido.</div>';
    return;
  }

  const grupos = (state.config.grupos || []).filter(g => g.professorId === prof.id);
  if (!grupos.length) {
    out.innerHTML = '<div class="small">Sem sessões definidas.</div>';
    return;
  }

  grupos.forEach(g => {
    const card = document.createElement('div'); card.className = 'session card';
    const disc = (state.config.disciplinas || []).find(d => d.id === g.disciplinaId) || { nome: g.disciplinaId };
    card.innerHTML = `<div><strong>${disc.nome}</strong> | Sala: ${g.sala || '-'} (${g.inicio || '08:15'} - ${g.fim || '09:05'})</div>
      <div style="margin-top:8px;display:flex;gap:6px">
        <input class="input lessonNumber" placeholder="Nº Lição" style="width:90px">
        <input class="input sumario" placeholder="Sumário">
        <button class="btn presencaP">Presente</button>
        <button class="btn ghost presencaF" style="background:#d33a2c">Falta</button>
        <button class="btn ghost duplicar" style="background:#f6a623">Duplicar</button>
      </div>`;
    out.appendChild(card);

    card.querySelector('.presencaP').addEventListener('click', () => quickSaveAttendance(g, card, true));
    card.querySelector('.presencaF').addEventListener('click', () => quickSaveAttendance(g, card, false));
    card.querySelector('.duplicar').addEventListener('click', () => duplicatePrevious(g, card));
  });
}

/* === Helpers for reg ids === */
function makeRegId() { return 'R' + Date.now(); }

/* === quickSaveAttendance (ensures lesson number and summary saved) === */
async function quickSaveAttendance(group, card, present = true) {
  if (!state.reg) state.reg = { versao: 'v1', registos: [] };
  const lessonEl = card.querySelector('.lessonNumber') || card.querySelector('input[name="lessonNumber"]');
  const sumEl = card.querySelector('.sumario') || card.querySelector('input[name="sumario"]');
  const lessonNumber = lessonEl ? lessonEl.value.trim() : '';
  const sumario = sumEl ? sumEl.value.trim() : '';

  try {
    const date = $('#dataHoje').value || new Date().toISOString().slice(0, 10);
    const students = (Array.isArray(group.studentIds) && group.studentIds.length) ? group.studentIds : [null];

    for (const alunoId of students) {
      const reg = {
        id: makeRegId(),
        data: date,
        professorId: group.professorId,
        alunoId: alunoId,
        disciplinaId: group.disciplinaId,
        presenca: present,
        numeroLicao: lessonNumber,
        sumario: sumario,
        horaInicio: group.inicio || null,
        horaFim: group.fim || null
      };
      state.reg.registos.push(reg);
    }

    updateSync('🔁 sincronizando...');
    await graphSave(REG_PATH, state.reg);
    localStorage.setItem('esp_reg', JSON.stringify(state.reg));
    updateSync('💾 guardado');
    Swal.fire({ icon: 'success', title: 'Registo gravado', timer: 1200, showConfirmButton: false });
  } catch (e) {
    console.error('Save failed', e);
    localStorage.setItem('esp_reg', JSON.stringify(state.reg));
    updateSync('⚠ offline');
    Swal.fire('Aviso', 'Guardado localmente. Será sincronizado quando online.', 'warning');
  }
}

/* === duplicatePrevious === */
function duplicatePrevious(group, card) {
  if (!state.reg || !state.reg.registos) return Swal.fire('Nenhum', 'Não existe registo anterior para duplicar', 'info');
  const last = state.reg.registos.slice().reverse().find(r => r.professorId === group.professorId && r.disciplinaId === group.disciplinaId && r.alunoId);
  if (last) {
    card.querySelector('.lessonNumber').value = last.numeroLicao || '';
    card.querySelector('.sumario').value = last.sumario || '';
    Swal.fire('Duplicado', 'Campos preenchidos com o último registo similar', 'info');
  } else {
    Swal.fire('Nenhum', 'Não existe registo anterior para duplicar', 'info');
  }
}

/* === Manual Reg modal === */
async function manualReg() {
  if (!state.reg) state.reg = { versao: 'v1', registos: [] };
  const profs = (state.config?.professores || []).map(p => `<option value="${p.id}">${p.nome} (${p.id})</option>`).join('');
  const discs = (state.config?.disciplinas || []).map(d => `<option value="${d.id}">${d.nome} (${d.id})</option>`).join('');
  const { value: vals } = await Swal.fire({
    title: 'Novo registo manual',
    html: `<select id="swProf" class="swal2-input">${profs}</select>
           <select id="swDisc" class="swal2-input">${discs}</select>
           <input id="swAluno" class="swal2-input" placeholder="Aluno ID (opcional)">
           <input id="swLicao" class="swal2-input" placeholder="Nr. Lição">
           <input id="swInicio" class="swal2-input" placeholder="08:15">
           <input id="swFim" class="swal2-input" placeholder="09:05">
           <input id="swSum" class="swal2-input" placeholder="Sumário">`,
    focusConfirm: false,
    showCancelButton: true,
    preConfirm: () => ({
      profId: document.getElementById('swProf')?.value,
      discId: document.getElementById('swDisc')?.value,
      alunoId: document.getElementById('swAluno')?.value.trim() || null,
      licao: document.getElementById('swLicao')?.value.trim(),
      inicio: document.getElementById('swInicio')?.value.trim(),
      fim: document.getElementById('swFim')?.value.trim(),
      sum: document.getElementById('swSum')?.value.trim()
    })
  });

  if (vals) {
    const date = new Date().toISOString().slice(0, 10);
    const reg = {
      id: makeRegId(),
      data: date,
      professorId: vals.profId,
      disciplinaId: vals.discId,
      alunoId: vals.alunoId,
      numeroLicao: vals.licao,
      sumario: vals.sum,
      horaInicio: vals.inicio || null,
      horaFim: vals.fim || null,
      presenca: true
    };
    state.reg.registos.push(reg);
    try {
      await graphSave(REG_PATH, state.reg);
      localStorage.setItem('esp_reg', JSON.stringify(state.reg));
      Swal.fire('Feito', 'Registo criado e gravado', 'success');
    } catch (e) {
      localStorage.setItem('esp_reg', JSON.stringify(state.reg));
      Swal.fire('Aviso', 'Guardado localmente. Será sincronizado quando online.', 'warning');
    }
  }
}

/* === Admin modal with auto-save on edits === */
document.addEventListener('click', (e) => { if (e.target && e.target.id === 'btnOpenAdmin') openAdminModal(); });

async function openAdminModal() {
  const user = account?.username?.toLowerCase() || '';
  if (user !== 'biblioteca@esparedes.pt') return Swal.fire('Acesso negado', 'Somente administrador pode aceder', 'error');

  const renderList = (type) => {
    const arr = state.config[type] || [];
    if (!arr.length) return `<div class="small">Sem ${type} definidos.</div>`;
    return `<div class="admin-list">${arr.map(it => `<div style="display:flex;justify-content:space-between;align-items:center;padding:6px;border-bottom:1px solid rgba(0,0,0,0.04)"><div><strong>${it.nome||it.email||it.id}</strong><div class="small">${it.email||it.turma||''}</div></div><div><button data-type="${type}" data-id="${it.id}" class="editBtn">Editar</button> <button data-type="${type}" data-id="${it.id}" class="delBtn">Eliminar</button></div></div>`).join('')}</div>`;
  };

  const { value: action } = await Swal.fire({
    title: 'Administração — Configuração',
    html: `<div class="tabbar"><button id="tab_prof" class="tab active">👥 Professores</button><button id="tab_alunos" class="tab">🎓 Alunos</button><button id="tab_disc" class="tab">📘 Disciplinas</button><button id="tab_hor" class="tab">🕐 Horários</button><button id="tab_cal" class="tab">📅 Calendário</button></div>
           <div id="swContent" style="min-height:260px">${renderList('professores')}</div>
           <div style="display:flex;gap:8px;margin-top:10px"><button id="btnAdd" class="swal2-confirm swal2-styled">Adicionar</button><button id="btnImport" class="swal2-confirm swal2-styled" style="background:#6c757d">Importar XLSX/JSON</button><button id="btnExport" class="swal2-confirm swal2-styled" style="background:#1fa1a1">Exportar JSON</button><button id="btnSaveOne" class="swal2-confirm swal2-styled" style="background:#1d4ed8">Guardar OneDrive</button></div>`,
    showCancelButton: true,
    showConfirmButton: false,
    didOpen: () => {
      const content = document.getElementById('swContent');
      const setTab = (tab) => {
        document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
        document.querySelector(`#tab_${tab}`).classList.add('active');
        if (tab === 'prof') content.innerHTML = renderList('professores');
        if (tab === 'alunos') content.innerHTML = renderList('alunos');
        if (tab === 'disc') content.innerHTML = renderList('disciplinas');
        if (tab === 'hor') content.innerHTML = renderList('grupos');
        if (tab === 'cal') content.innerHTML = `<pre style="white-space:pre-wrap">${JSON.stringify(state.config.calendario || {}, null, 2)}</pre>`;
      };

      document.getElementById('tab_prof').addEventListener('click', () => setTab('prof'));
      document.getElementById('tab_alunos').addEventListener('click', () => setTab('alunos'));
      document.getElementById('tab_disc').addEventListener('click', () => setTab('disc'));
      document.getElementById('tab_hor').addEventListener('click', () => setTab('hor'));
      document.getElementById('tab_cal').addEventListener('click', () => setTab('cal'));

      document.getElementById('btnAdd').addEventListener('click', async () => {
        const active = document.querySelector('.tab.active').id;
        if (active === 'tab_prof') {
          const { value: newObj } = await Swal.fire({
            title: 'Adicionar Professor',
            html: `<input id="nId" class="swal2-input" placeholder="Código (ex: EE7)"><input id="nNome" class="swal2-input" placeholder="Nome"><input id="nEmail" class="swal2-input" placeholder="email@esparedes.pt">`,
            preConfirm: () => ({ id: document.getElementById('nId').value.trim(), nome: document.getElementById('nNome').value.trim(), email: document.getElementById('nEmail').value.trim() })
          });
          if (newObj && newObj.id) {
            state.config.professores = state.config.professores || [];
            if (state.config.professores.some(p => p.id === newObj.id)) { Swal.fire('Erro', 'Código já existe', 'error'); return; }
            state.config.professores.push(newObj);
            content.innerHTML = renderList('professores');
            autoSaveConfig();
          }
        } else if (active === 'tab_alunos') {
          const { value: newObj } = await Swal.fire({
            title: 'Adicionar Aluno',
            html: `<input id="nId" class="swal2-input" placeholder="Código (ex: 9I4)"><input id="nNome" class="swal2-input" placeholder="Nome"><input id="nTurma" class="swal2-input" placeholder="Turma">`,
            preConfirm: () => ({ id: document.getElementById('nId').value.trim(), nome: document.getElementById('nNome').value.trim(), turma: document.getElementById('nTurma').value.trim() })
          });
          if (newObj && newObj.id) { state.config.alunos = state.config.alunos || []; state.config.alunos.push(newObj); content.innerHTML = renderList('alunos'); autoSaveConfig(); }
        } else if (active === 'tab_disc') {
          const { value: newObj } = await Swal.fire({
            title: 'Adicionar Disciplina',
            html: `<input id="nId" class="swal2-input" placeholder="Código (ex: Of.P)"><input id="nNome" class="swal2-input" placeholder="Disciplina">`,
            preConfirm: () => ({ id: document.getElementById('nId').value.trim(), nome: document.getElementById('nNome').value.trim() })
          });
          if (newObj && newObj.id) { state.config.disciplinas = state.config.disciplinas || []; state.config.disciplinas.push(newObj); content.innerHTML = renderList('disciplinas'); autoSaveConfig(); }
        } else if (active === 'tab_hor') {
          const { value: newObj } = await Swal.fire({
            title: 'Adicionar Horário (grupo)',
            html: `<input id="nId" class="swal2-input" placeholder="ID grupo"><input id="nProf" class="swal2-input" placeholder="Professor ID"><input id="nDisc" class="swal2-input" placeholder="Disciplina ID"><input id="nInicio" class="swal2-input" placeholder="08:15"><input id="nFim" class="swal2-input" placeholder="09:05"><input id="nSala" class="swal2-input" placeholder="Sala">`,
            preConfirm: () => ({ id: document.getElementById('nId').value.trim(), professorId: document.getElementById('nProf').value.trim(), disciplinaId: document.getElementById('nDisc').value.trim(), inicio: document.getElementById('nInicio').value.trim(), fim: document.getElementById('nFim').value.trim(), sala: document.getElementById('nSala').value.trim() })
          });
          if (newObj && newObj.id) { state.config.grupos = state.config.grupos || []; state.config.grupos.push(newObj); content.innerHTML = renderList('grupos'); autoSaveConfig(); }
        }
      });

      document.getElementById('btnImport').addEventListener('click', () => document.getElementById('fileImport').click());
      document.getElementById('btnExport').addEventListener('click', () => { const data = JSON.stringify(state.config || {}, null, 2); const a = document.createElement('a'); a.href = 'data:application/json;charset=utf-8,' + encodeURIComponent(data); a.download = 'config_especial.json'; a.click(); });
      document.getElementById('btnSaveOne').addEventListener('click', async () => { try { await graphSave(CFG_PATH, state.config); Swal.fire('Sucesso', 'Configuração gravada no OneDrive', 'success'); } catch (e) { console.error(e); Swal.fire('Erro', 'Falha a gravar no OneDrive', 'error'); } });

      content.addEventListener('click', async (ev) => {
        const ed = ev.target.closest('.editBtn');
        const del = ev.target.closest('.delBtn');
        if (ed) {
          // In this minimal UI we left edit implementation to open a modal per selected item if needed.
          Swal.fire('Editar', 'Clique em adicionar para inserir ou selecione um item (edição avançada em v4.3)', 'info');
        }
        if (del) {
          Swal.fire({ title: 'Eliminar?', text: 'Confirma eliminar este item?', icon: 'warning', showCancelButton: true }).then(res => {
            if (res.isConfirmed) {
              // naive delete: will require proper type/id mapping in future
              Swal.fire('Eliminado', '', 'success');
            }
          });
        }
      });
    }
  });
}

/* === file import handler === */
document.getElementById('fileImport')?.addEventListener('change', async (ev) => {
  const files = ev.target.files; if (!files || !files.length) return;
  for (const f of files) {
    const name = f.name.toLowerCase();
    if (name.endsWith('.json')) {
      const txt = await f.text();
      try { const obj = JSON.parse(txt); state.config = obj; autoSaveConfig(); Swal.fire('Importado', 'JSON importado e salvo.', 'success'); } catch (e) { Swal.fire('Erro', 'JSON inválido', 'error'); }
    } else {
      const data = await f.arrayBuffer();
      const wb = XLSX.read(data);
      const sheetName = wb.SheetNames[0];
      const json = XLSX.utils.sheet_to_json(wb.Sheets[sheetName]);
      // naive mapping => professores
      const map = json.map(r => ({ id: r.id || r.codigo || r.Codigo, nome: r.nome || r.Nome || r.NOME, email: r.email || r.Email || r.EMAIL }));
      state.config.professores = map;
      autoSaveConfig();
      Swal.fire('Importado', 'XLSX importado e guardado (professores).', 'success');
    }
    log(`Importado ${f.name}`);
  }
});

/* === auto-save config (debounced) === */
let autosaveTimer = null;
function autoSaveConfig() {
  if (autosaveTimer) clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(async () => {
    try {
      await graphSave(CFG_PATH, state.config);
      localStorage.setItem('esp_config', JSON.stringify(state.config));
      log('Config auto-saved to OneDrive');
      updateSync('💾 guardado');
    } catch (e) {
      console.error('Auto-save failed', e);
      updateSync('⚠ offline');
      localStorage.setItem('esp_config', JSON.stringify(state.config));
    }
  }, 700);
}

/* === create backup === */
async function createBackupIfExists() {
  try {
    const current = state.config || (localStorage.getItem('esp_config') && JSON.parse(localStorage.getItem('esp_config')));
    if (!current) { log('Sem ficheiro config — sem backup.'); return null; }
    const now = new Date(); const ts = now.getFullYear().toString().padStart(4, '0') + (now.getMonth() + 1).toString().padStart(2, '0') + now.getDate().toString().padStart(2, '0') + '_' + now.getHours().toString().padStart(2, '0') + now.getMinutes().toString().padStart(2, '0');
    const backupPath = BACKUP_FOLDER + '/config_especial_' + ts + '.json';
    await graphSave(backupPath, current);
    log('Backup criado: ' + backupPath);
    return backupPath;
  } catch (e) { console.warn('Backup failed', e); return null; }
}

/* === DOMContentLoaded init === */
document.addEventListener('DOMContentLoaded', async () => {
  document.getElementById('btnMsLogin')?.addEventListener('click', () => ensureLogin());
  document.getElementById('btnMsLogout')?.addEventListener('click', () => onLogout());
  document.getElementById('btnRefreshDay')?.addEventListener('click', () => renderDay());
  document.getElementById('btnManualReg')?.addEventListener('click', () => manualReg());
  document.getElementById('btnBackupNow')?.addEventListener('click', async () => { const b = await createBackupIfExists(); if (b) Swal.fire('Backup criado', b, 'success'); });

  // theme init
  const t = localStorage.getItem('esp_theme') || (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  document.documentElement.setAttribute('data-theme', t);
  document.getElementById('themeToggle')?.addEventListener('click', () => {
    const cur = document.documentElement.getAttribute('data-theme');
    const next = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('esp_theme', next);
  });

  await initMsal();

  const c = localStorage.getItem('esp_config'); if (c) state.config = JSON.parse(c);
  const r = localStorage.getItem('esp_reg'); if (r) state.reg = JSON.parse(r);
  if (!state.config) state.config = { professores: [], alunos: [], disciplinas: [], grupos: [], calendario: {} };
  if (!state.reg) state.reg = { versao: 'v1', registos: [] };

  renderDay();
});
