console.log('[DEBUG] Script starting...');
const ID_ALIASES = {
  chartAreaContainer: ['chartCob'],
  chartBalanceContainer: ['chartSaldo'],
  lblKpiConVac: ['lblConVac'],
  lblKpiSinVac: ['lblSinVac'],
  wizFilterBP: ['wizBP'],
  wizFilterGerencia: ['wizGer'],
  wizFilterArea: ['wizArea'],
  configModal: ['cfgModal'],
  cfgColaDir: ['cfgDir'],
  btnCloseConfigModal: ['btnCloseCfg'],
  btnSaveConfig: ['btnSaveCfg'],
  btnCloseKpiModal: ['btnCloseKpi'],
  previewModal: ['prevModal'],
  btnClosePreviewModal: ['btnClosePrev'],
  btnColab_Buscar: ['btnColabBuscar'],
  btnColab_Enviar: ['btnColabEnviar'],
  btnColab_Restaurar: ['btnColabRest'],
  btnGlobalSendNow: ['fabSend'],
  btnGlobalCancelQueue: ['fabCancel'],
  massGuardHint: ['massHint'],
  msgPreviewBox: ['msgPrev'],
  w3DestCount: ['w3Dest'],
  w3DestNames: ['w3Names'],
  w3TplName: ['w3Tpl'],
  w3MsgSnippet: ['w3Snip'],
  chkConfirmSend: ['chkConfirm'],
  colab_estado_box: ['colabBox'],
  colab_mensaje_status: ['colabStatus'],
  colabChTeams: ['colabTeams'],
  colabChSmtp: ['colabSmtp'],
  colabModoPrueba: ['colabPrueba']
};

const $ = (id) => {
  const direct = document.getElementById(id);
  if (direct) return direct;
  const aliases = ID_ALIASES[id] || [];
  for (const alias of aliases) {
    const node = document.getElementById(alias);
    if (node) return node;
  }
  return null;
};

function bind(id, eventName, handler) {
  const node = $(id);
  if (!node) return false;
  node.addEventListener(eventName, handler);
  return true;
}

const state = {
  supMap: {},
  jefesRef: [],
  personByName: {},
  personByMat: {},
  personOptions: [],
  supervisorNames: [],
  supervisoresData: [],
  objectiveRows: [],
  objectiveRowsPromise: null,
  editableRows: [],
  testerEmail: '',
  selectedSupervisores: new Set(),
  selectedPersonas: new Set(),
  wizardRecipientMode: 'personas',
  massArmed: false,
  massSendInFlight: false,
  pendingQueueFile: '',
  confirmacionesResumen: { total: 0, confirmadas: 0, pendientes: 0, avancePct: 0, loaded: false },
  trimestreLabel: 'Q?-????',
  trimestreModo: 'auto',
  lastDataError: ''
};

function applyTrimestreLabels(label) {
  const tri = (label || 'Q?-????').trim();
  // Las tarjetas 2 y 4 ahora son "Ya cumplieron meta" y "Avance de meta" (texto fijo).
  if ($('lblCoberturaArea')) $('lblCoberturaArea').textContent = `Cobertura de vacaciones ${tri} por área`;
}

async function loadTrimestreControl() {
  const data = await jfetch('/api/trimestre/estado');
  const ctrl = data.trimestre_control || {};
  state.trimestreLabel = ctrl.trimestre || state.trimestreLabel || 'Q?-????';
  state.trimestreModo = ctrl.modo || state.trimestreModo || 'auto';
  if ($('triModo')) $('triModo').value = state.trimestreModo;
  if ($('triManual')) $('triManual').value = state.trimestreLabel;
  applyTrimestreLabels(state.trimestreLabel);
}

async function applyTrimestreControl() {
  const modo = ($('triModo') ? $('triModo').value : 'auto') || 'auto';
  const trimestre = ($('triManual') ? $('triManual').value : '').trim();
  await jfetch('/api/trimestre/configurar', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ modo, trimestre })
  });
  await loadInit();
  await loadTrimestreControl();
  setStatus(`Control trimestral actualizado: ${state.trimestreLabel} (${state.trimestreModo}).`);
  notify(`Trimestre aplicado: ${state.trimestreLabel} (${state.trimestreModo})`, 'ok');
}

function normTxt(v) {
  return (v || '')
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function buildSearchText(nombre, matricula) {
  const n = normTxt(nombre).replace(/[,]+/g, ' ');
  const m = normTxt(matricula || '');
  const toks = n.split(/\s+/).filter(Boolean);
  const reversed = toks.slice().reverse().join(' ');
  return `${n} ${reversed} ${m}`.trim();
}

function matchByTokens(searchText, query) {
  const q = normTxt(query);
  if (!q) return true;
  if (searchText.includes(q)) return true;
  const t = q.split(/\s+/).filter(Boolean);
  return t.length > 0 && t.every(x => searchText.includes(x));
}

function refreshNombreSuggestions(query = '') {
  const q = normTxt(query);
  const fromPersons = state.personOptions
    .filter(p => matchByTokens(p.searchText || '', q))
    .slice(0, 80)
    .map(p => p.nombre);

  const fromSups = (state.supervisorNames || [])
    .filter(n => matchByTokens(buildSearchText(n, ''), q))
    .slice(0, 40);

  const _normKey = n => n.replace(/[,\.]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
  const _nameMap = new Map();
  [...fromSups, ...fromPersons].forEach(n => {
    const k = _normKey(n);
    if (!_nameMap.has(k)) _nameMap.set(k, n);
  });
  const merged = Array.from(_nameMap.values())
    .sort((a, b) => a.localeCompare(b, 'es'))
    .slice(0, 120);

  $('lstJefes').innerHTML = merged.map(n => `<option value="${n}"></option>`).join('');
}

function resolveNombreEntry(raw) {
  const value = (raw || '').trim();
  if (!value) return null;

  const byExactPerson = state.personByName[value.toLowerCase()] || state.personByMat[value];
  if (byExactPerson) {
    return { nombre: byExactPerson.nombre || value, email: byExactPerson.email || '', source: 'person-exact' };
  }

  const byExactSupervisor = state.supMap[value.toLowerCase()];
  if (byExactSupervisor) {
    return { nombre: byExactSupervisor.nombre || value, email: byExactSupervisor.email || '', source: 'supervisor-exact', supervisor: byExactSupervisor };
  }

  const q = normTxt(value);
  if (!q) return null;

  const personMatches = state.personOptions.filter(p => matchByTokens(p.searchText || '', value));
  if (personMatches.length === 1) {
    return { nombre: personMatches[0].nombre || value, email: personMatches[0].email || '', source: 'person-single' };
  }

  const supervisorMatches = (state.supervisorNames || [])
    .map(name => state.supMap[(name || '').toLowerCase()])
    .filter(Boolean)
    .filter(item => matchByTokens(buildSearchText(item.nombre || '', ''), value));
  if (supervisorMatches.length === 1) {
    return { nombre: supervisorMatches[0].nombre || value, email: supervisorMatches[0].email || '', source: 'supervisor-single', supervisor: supervisorMatches[0] };
  }

  return null;
}

function setMassArmed(armed) {
  state.massArmed = !!armed;
  if ($('btnAllSend')) $('btnAllSend').disabled = !state.massArmed;
  const hint = $('massGuardHint');
  if (hint) {
    hint.textContent = state.massArmed
      ? 'Envio masivo habilitado para esta sesion. Puedes cancelar mientras este en cola.'
      : 'Envio masivo bloqueado por seguridad. Primero acepta el envio masivo.';
  }
}

function normalizeSupervisorName(v) {
  return (v || '').toString().trim().toUpperCase();
}

function syncEditJefesFromSupervisores() {
  const names = (state.supervisoresData || [])
    .map(it => (it && it.nombre ? String(it.nombre).trim() : ''))
    .filter(Boolean);
  const uniq = Array.from(new Set(names)).sort((a, b) => a.localeCompare(b, 'es'));
  const list = $('lstEditJefe');
  if (list) list.innerHTML = uniq.map(j => `<option value="${j}"></option>`).join('');
}

function getSelectedSupervisores() {
  return Array.from(state.selectedSupervisores || []);
}

function getSelectedPersonas() {
  return Array.from(state.selectedPersonas || []);
}

function getWizardRecipientMode() {
  return state.wizardRecipientMode || 'personas';
}

function _personKey(row) {
  return ((row && row.matricula) || (row && row.nombre) || '').toString().trim().toUpperCase();
}

function getSelectedPersonRows() {
  const selected = new Set(getSelectedPersonas());
  return (state.objectiveRows || []).filter(row => selected.has(_personKey(row)));
}

function getVisiblePersonCards() {
  return Array.from(document.querySelectorAll('.persona-card')).filter(card => card && card.style.display !== 'none');
}

function getVisibleSelectedPersonRows() {
  const visibleKeys = getVisiblePersonCards()
    .map(card => (card.dataset.personKey || '').trim().toUpperCase())
    .filter(Boolean);
  const visibleSet = new Set(visibleKeys);
  return (state.objectiveRows || []).filter(row => visibleSet.has(_personKey(row)) && state.selectedPersonas.has(_personKey(row)));
}

function getEffectiveSelectedPersonRows() {
  return hasActiveJefeFilters() ? getVisibleSelectedPersonRows() : getSelectedPersonRows();
}

function hasActiveJefeFilters() {
  const bp = ($('wizFilterBP') && $('wizFilterBP').value) || '';
  const ger = ($('wizFilterGerencia') && $('wizFilterGerencia').value) || '';
  const area = ($('wizFilterArea') && $('wizFilterArea').value) || '';
  const q = (($('wizSearchJefe') && $('wizSearchJefe').value) || '').trim();
  return !!(bp || ger || area || q);
}

function getActiveWizardHrbpFilter() {
  return String((($('wizFilterBP') && $('wizFilterBP').value) || '')).trim();
}

function getVisibleJefeCards() {
  return Array.from(document.querySelectorAll('.jefe-card')).filter(card => {
    return card && card.style.display !== 'none';
  });
}

function getVisibleSelectedSupervisores() {
  return getVisibleJefeCards()
    .map(card => card.querySelector('.jefe-check[data-jefe]'))
    .filter(chk => chk && chk.checked)
    .map(chk => normalizeSupervisorName(chk.getAttribute('data-jefe')))
    .filter(Boolean);
}

function getEffectiveSelectedSupervisores() {
  if (getWizardRecipientMode() === 'personas') {
    const uniq = new Set();
    getEffectiveSelectedPersonRows().forEach(row => {
      const sup = normalizeSupervisorName(row.supervisor || '');
      if (sup) uniq.add(sup);
    });
    return Array.from(uniq);
  }
  return hasActiveJefeFilters() ? getVisibleSelectedSupervisores() : getSelectedSupervisores();
}

function updateMassGuardHint() {
  const hint = $('massGuardHint');
  if (!hint) return;
  const selected = getEffectiveSelectedSupervisores();
  const total = selected.length;
  const scopeBase = getWizardRecipientMode() === 'personas' ? 'jefes resueltos desde personas' : 'jefes';
  const scope = hasActiveJefeFilters() ? `${scopeBase} visibles con el filtro actual` : `${scopeBase} seleccionados`;
  hint.textContent = state.massArmed
    ? `Envio masivo habilitado. Destinatarios ${scope}: ${total}`
    : `Envio masivo bloqueado. Destinatarios ${scope}: ${total}`;
}

function getSelectedMassChannels() {
  return {
    enviar_teams: true,
    enviar_smtp: true,   // Activado por defecto — el servidor fuerza modo prueba si vacaciones_test_email está configurado
    encolar_pa: true,
  };
}

function getSelectedIndividualChannels() {
  return {
    enviar_teams: true,
    enviar_smtp: true,   // Activado por defecto — modo prueba activo via pa_config.json
    encolar_para_pa: true,
  };
}

function updateSummaryCards() {
  const mode = getWizardRecipientMode();
  // Usar solo los que aún no cumplieron su meta (mismo filtro que renderPersonasObjetivo)
  const totalPersonasObjetivo = (state.objectiveRows || []).filter(r => {
    if (!(r.persona_objetivo || r.objetivo || r.total_dias)) return false;
    const gozados = Number(r.dias_gozados || 0);
    const objetivo = Number(r.objetivo || 0);
    if (objetivo > 0 && gozados >= objetivo) return false;
    return true;
  }).length;
  if (mode === 'personas') {
    const selectedPersonas = getSelectedPersonas().length;
    const impactedJefes = getEffectiveSelectedSupervisores().length;
    if ($('sumLeaders')) $('sumLeaders').textContent = totalPersonasObjetivo || 0;
    if ($('sumSelected')) $('sumSelected').textContent = selectedPersonas || 0;
    if ($('sumQueue')) $('sumQueue').textContent = impactedJefes || 0;
    if ($('sumLeadersLabel')) $('sumLeadersLabel').innerHTML = 'Personas<br>objetivo';
    if ($('sumSelectedLabel')) $('sumSelectedLabel').innerHTML = 'Personas<br>seleccionadas';
    if ($('sumQueueLabel')) $('sumQueueLabel').innerHTML = 'Jefes<br>a notificar';
  } else {
    const visibleLeaders = getVisibleJefeCards().length;
    const selected = getEffectiveSelectedSupervisores().length;
    if ($('sumLeaders')) $('sumLeaders').textContent = visibleLeaders || state.supervisoresData.length || 0;
    if ($('sumSelected')) $('sumSelected').textContent = selected || 0;
    if ($('sumQueue')) $('sumQueue').textContent = totalPersonasObjetivo || 0;
    if ($('sumLeadersLabel')) $('sumLeadersLabel').innerHTML = 'Jefes<br>detectados';
    if ($('sumSelectedLabel')) $('sumSelectedLabel').innerHTML = 'Seleccionados<br>para enviar';
    if ($('sumQueueLabel')) $('sumQueueLabel').innerHTML = 'Personas<br>objetivo';
  }
  updateMassGuardHint();
}

async function refreshConfirmacionesResumen() {
  try {
    const data = await jfetch('/api/confirmaciones-vacaciones/resumen');
    state.confirmacionesResumen = {
      total: Number(data.total || 0),
      confirmadas: Number(data.confirmadas || 0),
      pendientes: Number(data.pendientes || 0),
      avancePct: Number(data.avance_pct || 0),
      loaded: true,
    };
  } catch (e) {
    log('refreshConfirmacionesResumen error:', e.message);
    state.confirmacionesResumen = {
      ...(state.confirmacionesResumen || {}),
      loaded: false,
    };
  }
  updateSummaryCards();
}

function renderDestinatarios() {
  const destRoot = $('destList');
  if (!destRoot) {
    updateSummaryCards();
    return;
  }
  const q = normTxt($('wizSearchJefe') ? $('wizSearchJefe').value : '');
  const bpFilter = $('wizFilterBP') ? $('wizFilterBP').value : '';

  const visibles = (state.supervisoresData || []).filter(it => {
    const txt = `${it.nombre || ''} ${(it.email || '')}`;
    const matchText = matchByTokens(txt, q);
    const hrbps = Array.isArray(it.hrbps) ? it.hrbps : (it.hrbp ? [it.hrbp] : []);
    const matchBP = bpFilter ? hrbps.includes(bpFilter) : true;
    return matchText && matchBP;
  });
  const html = visibles.map(it => {
    const key = normalizeSupervisorName(it.nombre);
    const checked = state.selectedSupervisores.has(key) ? 'checked' : '';
    const total = Number(it.total_colaboradores || 0);
    return `<label style="display:flex;align-items:flex-start;gap:8px;margin:4px 0"><input type="checkbox" data-sup="${key}" ${checked} /><span style="font-size:13px"><b>${it.nombre || '-'}</b> &lt;${it.email || '-'}&gt; <span style="color:#64748b">(${total} colab.)</span></span></label>`;
  }).join('');
  destRoot.innerHTML = html || '<div class="small">Sin resultados.</div>';
  destRoot.querySelectorAll('input[type="checkbox"][data-sup]').forEach(chk => {
    chk.addEventListener('change', () => {
      const k = chk.getAttribute('data-sup') || '';
      if (!k) return;
      if (chk.checked) state.selectedSupervisores.add(k);
      else state.selectedSupervisores.delete(k);
      updateSummaryCards();
    });
  });
  updateSummaryCards();
}

function selectVisibleDestinatarios(mark) {
  const destRoot = $('destList');
  if (!destRoot) {
    updateSummaryCards();
    return;
  }
  const checks = destRoot.querySelectorAll('input[type="checkbox"][data-sup]');
  checks.forEach(chk => {
    const k = chk.getAttribute('data-sup') || '';
    chk.checked = !!mark;
    if (!k) return;
    if (mark) state.selectedSupervisores.add(k);
    else state.selectedSupervisores.delete(k);
  });
  updateSummaryCards();
}

async function refreshQueueState() {
  const data = await jfetch('/api/cola-pa/ultima-pendiente');
  const p = data.pendiente;
  const fab = $('btnGlobalCancelQueue');
  const fabSend = $('btnGlobalSendNow');
  const btnNow = $('btnSendNow');

  if (!p) {
    state.pendingQueueFile = '';
    $('massQueueState').textContent = 'Cola PA: sin pendientes.';
    if (fab) fab.classList.add('hidden');
    if (fabSend) fabSend.classList.add('hidden');
    if (btnNow) btnNow.classList.add('hidden');
    updateSummaryCards();
    return;
  }
  state.pendingQueueFile = p.archivo || '';
  const seg = p.segundos_restantes || 0;
  const min = Math.floor(seg / 60);
  const s = seg % 60;
  const timerStr = min > 0 ? `${min}m ${s}s` : `${s}s`;

  $('massQueueState').innerHTML = `<b style="color:#0369a1">Cola PA: ${p.archivo}</b> | Libera en <span style="color:#b91c1c;font-weight:900">${timerStr}</span>`;
  if (fab) fab.classList.remove('hidden');
  if (fabSend) fabSend.classList.remove('hidden');
  if (btnNow) btnNow.classList.remove('hidden');
  updateSummaryCards();
}

async function enviarTodoAhora() {
  if (!state.pendingQueueFile) return;
  const fabSend = document.getElementById('fabSend');
  let originalHtml = '';
  if (fabSend) {
    originalHtml = fabSend.innerHTML;
    fabSend.innerHTML = '<span class="spin" style="width:16px;height:16px;border-width:2px;margin-right:8px;border-top-color:#fff;border-right-color:rgba(255,255,255,.3);border-bottom-color:rgba(255,255,255,.3);border-left-color:rgba(255,255,255,.3);vertical-align:middle"></span> Procesando...';
    fabSend.style.pointerEvents = 'none';
    fabSend.style.opacity = '0.8';
  }
  
  try {
    await jfetch('/api/cola-pa/liberar-todo', { method: 'POST' });
    await refreshQueueState();
    notify('Envío forzado iniciado. Los mensajes se están procesando.', 'ok');
  } catch (err) {
    notify('Error: ' + err.message, 'err');
  } finally {
    if (fabSend) {
      fabSend.innerHTML = originalHtml;
      fabSend.style.pointerEvents = 'auto';
      fabSend.style.opacity = '1';
    }
  }
}

// ─── Modo Prueba / Modo Real ──────────────────────────────────────────────────
let _modoEnvioTestEmail = '';

async function loadModoEnvio() {
  try {
    const d = await jfetch('/api/tester-config');
    _modoEnvioTestEmail = d.email_tester || '';
    _actualizarBtnModo();
  } catch (e) { /* silencioso si el servidor no responde aún */ }
}

function _actualizarBtnModo() {
  const btn = $('btnModoEnvio');
  const lbl = $('lblModoEnvio');
  if (!btn) return;
  const esPrueba = !!_modoEnvioTestEmail;
  if (esPrueba) {
    btn.innerHTML = '🔒 Prueba';
    btn.style.cssText = 'font-size:10px;padding:2px 9px;border-radius:99px;border:1.5px solid #d97706;background:#fef3c7;color:#92400e;cursor:pointer;font-weight:600;line-height:1.6;transition:all .15s';
    if (lbl) lbl.textContent = `correos van solo a ${_modoEnvioTestEmail}`;
  } else {
    btn.innerHTML = '✅ Real';
    btn.style.cssText = 'font-size:10px;padding:2px 9px;border-radius:99px;border:1.5px solid #16a34a;background:#dcfce7;color:#166534;cursor:pointer;font-weight:600;line-height:1.6;transition:all .15s';
    if (lbl) lbl.textContent = 'correos van a supervisores reales';
  }
}

async function toggleModoEnvio() {
  const btn = $('btnModoEnvio');
  if (btn) btn.disabled = true;
  try {
    const esPrueba = !!_modoEnvioTestEmail;
    const nuevoEmail = esPrueba ? '' : 'jlopezp@usil.edu.pe';
    const confirmMsg = esPrueba
      ? '¿Activar Modo Real? Los correos irán a los supervisores reales.'
      : '¿Volver a Modo Prueba? Los correos solo llegarán a ti.';
    if (!window.confirm(confirmMsg)) return;
    const d = await jfetch('/api/tester-config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email_tester: nuevoEmail })
    });
    _modoEnvioTestEmail = d.email_tester || '';
    _actualizarBtnModo();
    const msg = _modoEnvioTestEmail
      ? `🔒 Modo Prueba activado — correos solo a ${_modoEnvioTestEmail}`
      : '✅ Modo Real activado — los correos irán a los supervisores reales';
    notify(msg, _modoEnvioTestEmail ? 'warn' : 'ok');
  } catch (e) {
    notify('No se pudo cambiar el modo: ' + e.message, 'err');
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function cancelarUltimoPendiente() {
  await refreshQueueState();
  if (!state.pendingQueueFile) {
    setStatus('No hay archivo pendiente por cancelar.');
    notify('No hay envío pendiente para cancelar.', 'warn');
    return;
  }
  const ok = window.confirm('Se cancelaran todos los envíos pendientes para evitar que se manden mensajes. ¿Deseas continuar?');
  if (!ok) return;
  const data = await jfetch('/api/cola-pa/cancelar-todos', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ motivo: 'cancelado global desde boton flotante' })
  });
  await refreshQueueState();
  const total = Number(data.total_cancelados || 0);
  setStatus(`Envío cancelado. Pendientes anulados: ${total}.`);
  notify(`Envío cancelado correctamente. Se anularon ${total} pendiente(s).`, 'ok');
}

function setKpiActive(filterValue) {
  ['kpiTotal', 'kpiEleg', 'kpiSinIniciar', 'kpiCasi', 'kpiAlerta', 'kpiParciales', 'kpiSinMeta'].forEach(id => {
    const el = $(id);
    if (!el) return;
    el.classList.toggle('active', (el.dataset.filter || '') === filterValue);
  });
}

function applyKpiFilter(filterValue) {
  const filtro = $('fEleg');
  if (filtro) filtro.value = filterValue;
  setKpiActive(filterValue);
  loadRanking().catch(e => log(e.message));
}

function _nombreKpi(filterValue) {
  if (filterValue === 'elegibles') return 'Elegibles';
  if (filterValue === 'con_vacaciones') return `Con Vacaciones ${state.trimestreLabel || 'Trimestre'}`;
  if (filterValue === 'con_saldo') return 'Con Saldo Vacacional';
  if (filterValue === 'sin_vacaciones') return `Sin Vacaciones ${state.trimestreLabel || 'Trimestre'}`;
  if (filterValue === 'con_meta') return 'Colaboradores objetivo';
  if (filterValue === 'cumplieron') return 'Cumplieron';
  if (filterValue === 'sin_iniciar') return 'Sin Iniciar (0 días)';
  if (filterValue === 'parciales') return 'Parciales (en proceso)';
  if (filterValue === 'casi_listos') return 'Casi listos (75-99%)';
  if (filterValue === 'proyeccion') return 'Predicción de Cierre';
  if (filterValue === 'ya_iniciaron') return 'Ya Iniciaron';
  if (filterValue === 'sin_meta_con_vac') return 'Sin Meta con Vacaciones';
  return 'Total Colaboradores';
}

function _renderRowsHtml(rows) {
  return (rows || []).map(r => `
    <tr data-mat="${r.matricula}">
      <td>${r.matricula || '-'}</td>
      <td class="editable" data-field="nombre">${r.nombre || '-'}</td>
      <td>${r.gerencia || r.departamento || '-'}</td>
      <td class="editable" data-field="area">${r.area || '-'}</td>
      <td class="editable" data-field="puesto">${r.puesto || '-'}</td>
      <td>${r.hrbp || '-'}</td>
      <td class="editable num" data-field="dias_gozados_q1" style="font-weight:700">${Number(r.dias_gozados || 0).toFixed(1)}</td>
      <td class="editable num" data-field="saldo_total_dias" style="font-weight:800;color:var(--brand)">${Number(r.total_dias || 0).toFixed(1)}</td>
      <td>${r.estado || '-'}</td>
    </tr>`).join('');
}

function openKpiModal(filterValue, rows) {
  $('kpiModalTitle').textContent = `Detalle KPI: ${_nombreKpi(filterValue)} (${(rows || []).length})`;
  $('tbKpiModal').innerHTML = _renderRowsHtml(rows);
  $('kpiModal').classList.remove('hidden');
  $('kpiModal').style.display = 'flex';
}

function closeKpiModal() {
  $('kpiModal').classList.add('hidden');
  $('kpiModal').style.display = '';
}

function openPreviewModal() {
  $('previewModal').classList.remove('hidden');
  $('previewModal').style.display = 'flex';
}

function closePreviewModal() {
  $('previewModal').classList.add('hidden');
  $('previewModal').style.display = '';
}

async function openKpiModalByFilter(filterValue) {
  applyKpiFilter(filterValue);
  $('kpiModalTitle').textContent = `Detalle KPI: ${_nombreKpi(filterValue)} (cargando...)`;
  $('tbKpiModal').innerHTML = '<tr><td colspan="7" style="text-align:center;color:#64748b">Cargando datos...</td></tr>';
  $('kpiModal').classList.remove('hidden');
  $('kpiModal').style.display = 'flex';
  try {
    let rows;
    if (filterValue === 'con_meta' || filterValue === 'cumplieron' || filterValue === 'sin_iniciar' || filterValue === 'parciales' || filterValue === 'casi_listos' || filterValue === 'ya_iniciaron' || filterValue === 'sin_meta_con_vac') {
      const j = await jfetch('/api/vacaciones/meta_detalle?seg=' + encodeURIComponent(filterValue));
      rows = (j && j.registros) || [];
    } else {
      rows = await fetchRankingRows(filterValue);
    }
    $('tbKpiModal').innerHTML = _renderRowsHtml(rows);
    $('kpiModalTitle').textContent = `Detalle KPI: ${_nombreKpi(filterValue)} (${rows.length})`;
    attachInlineEditListeners($('tbKpiModal'));
  } catch (e) {
    $('kpiModalTitle').textContent = `Detalle KPI: ${_nombreKpi(filterValue)} (error)`;
    $('tbKpiModal').innerHTML = `<tr><td colspan="7" style="text-align:center;color:#9f1239">${(e && e.message) || 'No se pudo cargar el detalle.'}</td></tr>`;
  }
}

function setDashboardFallback(reason = '') {
  if ($('kTotal')) $('kTotal').textContent = '0';
  if ($('kSinMeta')) $('kSinMeta').textContent = '0';
  if ($('kEleg')) $('kEleg').textContent = '0%';
  if ($('kParciales')) $('kParciales').textContent = '0%';
  if ($('kSinIniciar')) $('kSinIniciar').textContent = '0%';
  if ($('kCasi')) $('kCasi').textContent = '0%';
  if ($('kAlerta')) $('kAlerta').textContent = '0%';
  if ($('updatedAt')) {
    const base = `Actualizado: ${new Date().toLocaleString()} | Trimestre: ${state.trimestreLabel} (${state.trimestreModo})`;
    $('updatedAt').textContent = reason ? `${base} | Estado: Error de datos` : base;
  }
  if ($('chartAreaContainer')) {
    $('chartAreaContainer').innerHTML = '<div class="small" style="padding:12px;text-align:center;color:#9f1239">Sin datos para cobertura por área.</div>';
  }
  if ($('chartBalanceContainer')) {
    $('chartBalanceContainer').innerHTML = '<div class="small" style="padding:12px;text-align:center;color:#9f1239">Sin datos para saldo acumulado.</div>';
  }
}

// ── KPIs de meta (conteos de colaboradores + avance en tiempo real) ──────────
function computeMetaKpis(rows) {
  let conMeta = 0, cumplieron = 0, sinIniciar = 0, casiListos = 0, parciales = 0;
  for (const r of (rows || [])) {
    const obj = Number(r.objetivo || 0);
    if (obj <= 0) continue;
    conMeta++;
    const reg = Number(r.dias_gozados || 0);
    if (reg >= obj) { cumplieron++; continue; }
    if (reg === 0) { sinIniciar++; continue; }
    parciales++;
    const pct = obj > 0 ? reg / obj : 0;
    if (pct >= 0.75 && pct < 1.0) casiListos++;
  }
  return { conMeta, cumplieron, sinIniciar, casiListos, parciales };
}

function renderAvanceKpi(kp) {
  const av = (kp && kp.avance != null && !isNaN(kp.avance)) ? Number(kp.avance) : null;
  const num = $('kAlerta'), bar = $('kAvanceBar'), foot = $('kAvanceFoot');
  if (num) num.textContent = (av == null) ? '—' : (av * 100).toFixed(1) + '%';
  if (bar) {
    requestAnimationFrame(() => {
      const pct = (av == null) ? 0 : Math.max(0, Math.min(100, av * 100));
      bar.style.width = pct.toFixed(1) + '%';
      // Color-code: red < 70%, yellow 70-90%, green > 90%
      bar.classList.remove('bar-red', 'bar-yellow', 'bar-green');
      if (pct < 70) bar.classList.add('bar-red');
      else if (pct < 90) bar.classList.add('bar-yellow');
      else bar.classList.add('bar-green');
    });
  }
  // Also color the number
  if (num && av != null) {
    const pct = av * 100;
    if (pct < 70) num.style.color = '#c0392b';
    else if (pct < 90) num.style.color = '#e08a1e';
    else num.style.color = '#1f9d55';
  }
  if (foot && kp && kp.registrado_total != null && kp.meta_total != null) {
    foot.textContent = Math.round(kp.registrado_total).toLocaleString('es-PE') + ' / ' +
                       Math.round(kp.meta_total).toLocaleString('es-PE') + ' días';
  } else if (foot) {
    foot.textContent = 'en tiempo real';
  }
}

async function renderKpiCards(retries = 3, delayMs = 3000) {
  // Una sola fuente: /api/vacaciones/kpis (avance + conteos del MISMO BASE GENERAL).
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const j = await jfetch('/api/vacaciones/kpis');
      const kp = (j && j.kpis) || {};
      const fmt  = (v) => (v == null ? '0' : Number(v).toLocaleString('es-PE'));
      const fmtD = (v) => (v == null ? '0' : Math.round(Number(v)).toLocaleString('es-PE'));

      const total       = Number(kp.con_meta       || 0);
      const cumplieron  = Number(kp.cumplieron     || 0);
      const sinIniciar  = Number(kp.sin_iniciar    || 0);
      const parciales   = Number(kp.parciales      || 0) || Math.max(0, total - cumplieron - sinIniciar);
      const sinMetaN    = Number(kp.sin_meta_con_vac || 0);

      // — CON META —
      if ($('kTotal')) $('kTotal').textContent = fmt(total);
      const diasMetaTotal = Number(kp.dias_meta_total || 0);
      if ($('kTotalFoot')) $('kTotalFoot').textContent = fmtD(diasMetaTotal) + ' días de meta';

      // — SIN META CON VAC —
      if ($('kSinMeta')) $('kSinMeta').textContent = fmt(sinMetaN);
      const diasSinMeta = Number(kp.dias_gozados_sin_meta || 0);
      if ($('kSinMetaFoot')) $('kSinMetaFoot').textContent = fmtD(diasSinMeta) + ' días gozados';

      // — CUMPLIERON —
      if ($('kEleg')) {
        const pct = total > 0 ? (cumplieron / total * 100) : 0;
        $('kEleg').textContent = pct.toFixed(1) + '%';
      }
      const diasCumpMeta = Number(kp.dias_meta_cumplieron    || 0);
      const diasCumpGoz  = Number(kp.dias_gozados_cumplieron || 0);
      if ($('kElegFoot')) $('kElegFoot').textContent =
        fmt(cumplieron) + ' con meta · ' + fmt(sinMetaN) + ' sin meta';
      if ($('kElegDays')) $('kElegDays').textContent =
        fmtD(diasCumpGoz) + ' / ' + fmtD(diasCumpMeta) + ' d.';

      // — PARCIALES —
      if ($('kParciales')) {
        const pct = total > 0 ? (parciales / total * 100) : 0;
        $('kParciales').textContent = pct.toFixed(1) + '%';
      }
      const diasParcMeta = Number(kp.dias_meta_parciales    || 0);
      const diasParcGoz  = Number(kp.dias_gozados_parciales || 0);
      if ($('kParcialesFoot')) $('kParcialesFoot').textContent = fmt(parciales) + ' con meta pendiente';
      if ($('kParcialesDays')) $('kParcialesDays').textContent =
        fmtD(diasParcGoz) + ' goz. / ' + fmtD(diasParcMeta) + ' d. meta';

      // — SIN INICIAR —
      if ($('kSinIniciar')) {
        const pct = total > 0 ? (sinIniciar / total * 100) : 0;
        $('kSinIniciar').textContent = pct.toFixed(1) + '%';
      }
      if ($('kSinIniciarFoot')) $('kSinIniciarFoot').textContent = fmt(sinIniciar) + ' · 0 días gozados';
      // Predicción de cierre usando fechas reales de la campaña (inicio trimestre → fecha límite de campaña)
      if ($('kCasi')) {
        // Parsear trimestre desde state (e.g. "Q2-2026")
        const triLabel = (state && state.trimestreLabel) || 'Q2-2026';
        const triMatch = triLabel.match(/Q(\d)-(\d{4})/i);
        const quarter = triMatch ? parseInt(triMatch[1]) : 2;
        const year    = triMatch ? parseInt(triMatch[2]) : new Date().getFullYear();
        const qStartMonth = (quarter - 1) * 3;           // Q1→0, Q2→3, Q3→6, Q4→9
        const qStart = new Date(year, qStartMonth, 1);

        // Fecha de cierre de campaña: pa_config puede sobreescribirla vía state.fechaCierreCampana.
        // Por defecto Q2 va hasta el 31 de agosto (2 meses adicionales de ejecución).
        let qEnd;
        if (state.fechaCierreCampana) {
          qEnd = new Date(state.fechaCierreCampana);
        } else if (quarter === 2) {
          qEnd = new Date(year, 7, 31); // 31 de agosto
        } else if (quarter === 1) {
          qEnd = new Date(year, 4, 31); // 31 de mayo
        } else if (quarter === 3) {
          qEnd = new Date(year, 10, 30); // 30 de noviembre
        } else {
          qEnd = new Date(year, 1, 28);  // 28 de febrero año siguiente
        }

        const now = new Date();
        const totalDays = (qEnd - qStart) / 86400000;
        let elapsedDays = (now - qStart) / 86400000;
        elapsedDays = Math.max(1, Math.min(totalDays, elapsedDays));

        const remainingDays = Math.max(0, totalDays - elapsedDays);
        const currentAvance = kp.avance != null && !isNaN(kp.avance) ? Number(kp.avance) : 0;
        const dailyRate = currentAvance / elapsedDays; // ritmo real observado

        // Tres escenarios con factor de desaceleración:
        // En campañas de RRHH el ritmo cae a medida que avanzan las semanas (casos difíciles quedan al final).
        // Conservador: 50% del ritmo actual · Probable: 68% · Optimista: 85%
        const predConservador = Math.min(100, (currentAvance + dailyRate * 0.50 * remainingDays) * 100);
        const predProbable    = Math.min(100, (currentAvance + dailyRate * 0.68 * remainingDays) * 100);
        const predOptimista   = Math.min(100, (currentAvance + dailyRate * 0.85 * remainingDays) * 100);

        // Mostrar el escenario "Probable" como headline
        const predicted = predProbable;
        const kCasiEl = $('kCasi');
        if (kCasiEl) {
          kCasiEl.textContent = predicted.toFixed(1) + '%';
          if (predicted < 70)      kCasiEl.style.color = '#c0392b';
          else if (predicted < 90) kCasiEl.style.color = '#e08a1e';
          else                     kCasiEl.style.color = '#1f9d55';
        }

        // Rango visual conservador–optimista debajo del headline
        const kPredRange = $('kPredRange');
        if (kPredRange) {
          const cColor = predConservador < 70 ? '#c0392b' : predConservador < 90 ? '#e08a1e' : '#1f9d55';
          const oColor = predOptimista  < 70 ? '#c0392b' : predOptimista  < 90 ? '#e08a1e' : '#1f9d55';
          kPredRange.innerHTML =
            `<span style="color:${cColor}">${predConservador.toFixed(0)}%</span>` +
            `<span class="pred-range-sep">—</span>` +
            `<span style="color:${oColor}">${predOptimista.toFixed(0)}%</span>` +
            `<span class="pred-range-label">conserv. / optim.</span>`;
        }

        const kPredWrap = $('kPredictionWrap');
        if (kPredWrap) {
          const daysLeft = Math.round(remainingDays);
          const fechaCierreStr = qEnd.toLocaleDateString('es-PE', { day: 'numeric', month: 'long', year: 'numeric' });
          kPredWrap.title = `Cierre de campaña: ${fechaCierreStr} · Quedan ${daysLeft} d. · Ritmo actual: ${(dailyRate * 100).toFixed(2)}%/día · Escenario probable asume desaceleración del 32%`;
        }
      }
      renderAvanceKpi(kp);
      return; // éxito
    } catch (e) {
      if (attempt < retries) {
        // Servidor aún procesando el Excel; esperar y reintentar
        await new Promise(r => setTimeout(r, delayMs));
      } else {
        // Todos los intentos fallaron: limpiar spinners con 0s para que no queden infinitos
        setDashboardFallback('Error al cargar KPIs');
        renderAvanceKpi(null);
      }
    }
  }
}

async function loadObjectiveRows(force = false) {
  if (!force && state.objectiveRows && state.objectiveRows.length > 0) return state.objectiveRows;
  if (state.objectiveRowsPromise && !force) return state.objectiveRowsPromise;

  state.objectiveRowsPromise = (async () => {
    try {
      const data = await jfetch('/api/objetivos?limite=5000');
      state.objectiveRows = data.registros || [];
      state.lastDataError = '';
      if (getWizardRecipientMode() === 'personas') renderWizardRecipients();
      return state.objectiveRows;
    } catch (e) {
      console.error('Error loading objectives:', e);
      const msg = `No se pudieron cargar los objetivos: ${(e && e.message) || 'Error desconocido'}`;
      state.lastDataError = msg;
      setStatus(msg, true);
      setDashboardFallback(msg);
      return [];
    } finally {
      state.objectiveRowsPromise = null;
    }
  })();
  return state.objectiveRowsPromise;
}

async function loadEditableRowsAutocomplete() {
  try {
    const data = await jfetch('/api/colaboradores_editor?limite=300');
    const items = data.items || [];
    const names = new Set();
    const areas = new Set();
    for (const item of items) {
      if (item.nombre) names.add(item.nombre);
      if (item.area) areas.add(item.area);
    }
    $('lstEditQ').innerHTML = Array.from(names).map(n => `<option value="${n}"></option>`).join('');
    $('lstEditArea').innerHTML = Array.from(areas).sort().map(a => `<option value="${a}"></option>`).join('');
    // El filtro de jefe se alinea a los lideres del bloque de campaña.
    syncEditJefesFromSupervisores();
  } catch (e) {
    log('Error loading autocomplete:', e);
  }
}

function renderEditableRows() {
  const rows = state.editableRows || [];
  if (!rows.length) {
    $('tbEditRows').innerHTML = '<tr><td colspan="11" style="text-align:center;color:#64748b">Sin resultados para los filtros actuales.</td></tr>';
    return;
  }
  $('tbEditRows').innerHTML = rows.map(r => `
    <tr data-mat="${r.matricula}">
      <td>${r.matricula || '-'}</td>
      <td><input class="edit-input" data-field="nombre" value="${(r.nombre || '').replace(/"/g, '&quot;')}" /></td>
      <td><input class="edit-input" data-field="email" value="${(r.email || '').replace(/"/g, '&quot;')}" /></td>
      <td>
        <input class="edit-input" data-field="area" value="${(r.area || '').replace(/"/g, '&quot;')}" style="margin-bottom:6px" />
        <input class="edit-input" data-field="departamento" value="${(r.departamento || '').replace(/"/g, '&quot;')}" placeholder="Departamento" />
      </td>
      <td>
        <input class="edit-input" data-field="jefe" value="${(r.jefe || '').replace(/"/g, '&quot;')}" style="margin-bottom:6px" />
        <input class="edit-input" data-field="puesto" value="${(r.puesto || '').replace(/"/g, '&quot;')}" placeholder="Puesto" />
      </td>
      <td><input class="edit-input edit-numero" type="number" step="0.1" min="0" data-field="dias_gozados_q1" value="${Number(r.dias_gozados_q1 || 0).toFixed(1)}" style="width:60px" title="Gozado Q1" /></td>
      <td><input class="edit-input edit-numero" type="number" step="0.1" min="0" data-field="vac_pendiente" value="${Number(r.vac_pendiente || 0).toFixed(1)}" style="width:60px" title="Pendiente" /></td>
      <td><input class="edit-input edit-numero" type="number" step="0.1" min="0" data-field="vac_vencidas" value="${Number(r.vac_vencidas || 0).toFixed(1)}" style="width:60px" title="Vencidas" /></td>
      <td><input class="edit-input edit-numero" type="number" step="0.1" min="0" data-field="vac_truncas" value="${Number(r.vac_truncas || 0).toFixed(1)}" style="width:60px" title="Truncas" /></td>
      <td><input class="edit-input edit-numero" type="number" step="0.1" min="0" data-field="saldo_total_dias" value="${Number(r.saldo_total_dias || 0).toFixed(1)}" style="width:60px;font-weight:bold" title="Saldo Total" /></td>
      <td><button class="btn" data-save-row="${r.matricula}">Guardar</button></td>
    </tr>
  `).join('');
  $('tbEditRows').querySelectorAll('[data-save-row]').forEach(btn => {
    btn.addEventListener('click', () => saveEditableRow(btn.getAttribute('data-save-row')));
  });
}

async function loadEditableRows() {
  $('editMeta').textContent = 'Cargando colaboradores editables...';
  const qs = new URLSearchParams();
  if ($('editQ').value.trim()) qs.set('q', $('editQ').value.trim());
  if ($('editArea').value.trim()) qs.set('area', $('editArea').value.trim());
  if ($('editJefe').value.trim()) qs.set('jefe', $('editJefe').value.trim());
  qs.set('limite', '300');
  const data = await jfetch('/api/colaboradores_editor?' + qs.toString());
  state.editableRows = data.items || [];
  $('editMeta').textContent = `Resultados: ${state.editableRows.length} de ${data.total || state.editableRows.length}`;
  renderEditableRows();
}

async function saveEditableRow(matricula, updatedFields = {}) {
  const payload = {
    matricula,
    ...updatedFields
  };

  try {
    await jfetch('/api/colaboradores_editor', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    // Actualizar localmente state.objectiveRows para no recargar todo el tiempo
    const idx = (state.objectiveRows || []).findIndex(r => r.matricula === matricula);
    if (idx !== -1) {
      state.objectiveRows[idx] = { ...state.objectiveRows[idx], ...updatedFields };
    }

    notify(`Cambios guardados para ${matricula}.`, 'ok', 2000);
  } catch (e) {
    notify(`Error al guardar: ${e.message}`, 'err');
  }
}

function attachInlineEditListeners(parent = document) {
  parent.querySelectorAll('.editable').forEach(cell => {
    cell.addEventListener('click', function (e) {
      if (this.querySelector('input')) return;

      const field = this.dataset.field;
      const tr = this.closest('tr');
      const mat = tr.dataset.mat;
      const originalText = this.textContent.trim();

      const input = document.createElement('input');
      input.className = 'edit-cell-input';
      input.value = originalText === '-' ? '' : originalText;
      if (this.classList.contains('num')) {
        input.type = 'number';
        input.step = '0.1';
      }

      this.innerHTML = '';
      this.appendChild(input);
      input.focus();

      const finishEdit = async () => {
        const newVal = input.value.trim();
        const finalVal = newVal || '-';
        this.innerHTML = finalVal;

        if (newVal !== originalText && newVal !== '') {
          const payload = {};
          payload[field] = newVal;
          await saveEditableRow(mat, payload);
          // Recargar ligeramente los componentes visuales
          loadVacacionesCharts().catch(() => { });
        }
      };

      input.addEventListener('blur', finishEdit);
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          input.blur();
        } else if (e.key === 'Escape') {
          this.innerHTML = originalText;
        }
      });
    });
  });
}

async function fetchRankingRows(filterValue) {
  const rows = await loadObjectiveRows();
  const query = (($('fBuscar') && $('fBuscar').value) || '').trim().toLowerCase();
  let out = [...rows];

  if (filterValue === 'con_vacaciones') {
    out = out.filter(r => !!r.tiene_vacaciones_q1);
  } else if (filterValue === 'con_saldo') {
    out = out.filter(r => !!r.tiene_saldo_vacacional);
  } else if (filterValue === 'sin_vacaciones') {
    out = out.filter(r => !r.tiene_vacaciones_q1 && !!r.tiene_saldo_vacacional);
  } else if (filterValue === 'con_meta') {
    out = out.filter(r => Number(r.objetivo || 0) > 0);
  } else if (filterValue === 'cumplieron') {
    out = out.filter(r => Number(r.objetivo || 0) > 0 && Number(r.dias_gozados || 0) >= Number(r.objetivo || 0));
  } else if (filterValue === 'sin_iniciar') {
    out = out.filter(r => Number(r.objetivo || 0) > 0 && Number(r.dias_gozados || 0) === 0);
  } else if (filterValue === 'parciales') {
    out = out.filter(r => {
      const obj = Number(r.objetivo || 0), reg = Number(r.dias_gozados || 0);
      return obj > 0 && reg > 0 && reg < obj;
    });
  } else if (filterValue === 'casi_listos') {
    out = out.filter(r => {
      const obj = Number(r.objetivo || 0), reg = Number(r.dias_gozados || 0);
      if (obj <= 0 || reg >= obj) return false;
      const pct = reg / obj;
      return pct >= 0.75 && pct < 1.0;
    });
  }

  if (query) {
    out = out.filter(r => [r.nombre, r.area, r.puesto, r.departamento, r.hrbp, r.seccion]
      .filter(Boolean)
      .some(v => v.toString().toLowerCase().includes(query)));
  }

  out.sort((a, b) => (Number(b.total_dias || 0) - Number(a.total_dias || 0)) || String(a.nombre || '').localeCompare(String(b.nombre || ''), 'es'));
  return out.slice(0, 500);
}

function renderCoverageChart(rows) {
  const agg = {};
  for (const r of rows || []) {
    const name = (r.area || r.departamento || 'SIN AREA').toString().trim();
    if (!agg[name]) agg[name] = { total: 0, conVacaciones: 0 };
    agg[name].total += 1;
    if (r.tiene_vacaciones_q1) agg[name].conVacaciones += 1;
  }
  const items = Object.entries(agg)
    .map(([name, vals]) => ({
      name,
      total: vals.total,
      conVacaciones: vals.conVacaciones,
      ratio: vals.total ? (vals.conVacaciones / vals.total) * 100 : 0,
    }))
    .sort((a, b) => b.ratio - a.ratio)
    .slice(0, 10);

  const container = $('chartAreaContainer');
  if (!container) return;
  if (!items.length) {
    container.innerHTML = '<div class="small" style="padding:12px;text-align:center;color:var(--muted)">Sin datos para cobertura por área.</div>';
    return;
  }
  container.innerHTML = `<div class="v-stack" style="gap:10px">
    ${items.map(it => {
    const color = it.ratio > 70 ? '#10b981' : (it.ratio > 40 ? '#f59e0b' : '#ef4444');
    return `
        <div class="bar-row" onclick="showCoverageModal('${it.name}')" style="cursor:pointer" title="Ver detalle de ${it.name}">
          <div class="bar-info">
            <span class="bar-label">${it.name}</span>
            <span class="bar-meta">${it.conVacaciones}/${it.total} colaboradores</span>
          </div>
          <div class="bar-track">
            <div class="bar-fill" style="width:${it.ratio}%; background-color:${color}"></div>
            <span class="bar-value">${it.ratio.toFixed(1)}%</span>
          </div>
        </div>
      `;
  }).join('')}
  </div>`;
}

function renderBalanceChart(rows) {
  const agg = {};
  for (const r of rows || []) {
    const name = (r.area || r.departamento || 'SIN AREA').toString().trim();
    if (!agg[name]) agg[name] = { saldo: 0, porProgramar: 0 };
    agg[name].saldo += Number(r.total_dias || 0);
    agg[name].porProgramar += Number(r.por_programar || 0);
  }
  const items = Object.entries(agg)
    .map(([name, vals]) => ({ name, saldo: vals.saldo, porProgramar: vals.porProgramar }))
    .sort((a, b) => b.saldo - a.saldo)
    .slice(0, 10);

  const container = $('chartBalanceContainer');
  if (!container) return;
  if (!items.length) {
    container.innerHTML = '<div class="small" style="padding:12px;text-align:center;color:var(--muted)">Sin datos para saldo acumulado.</div>';
    return;
  }
  const max = Math.max(...items.map(it => it.saldo), 1);
  container.innerHTML = `<div class="v-stack" style="gap:10px">
    ${items.map(it => {
    const width = (it.saldo / max) * 100;
    return `
        <div class="bar-row" onclick="showBalanceModal('${it.name}')" style="cursor:pointer" title="Ver detalle de ${it.name}">
          <div class="bar-info">
            <span class="bar-label">${it.name}</span>
            <span class="bar-meta">${it.saldo.toFixed(0)} días saldo</span>
          </div>
          <div class="bar-track">
            <div class="bar-fill" style="width:${width}%; background-color:#3b82f6"></div>
            <span class="bar-value">${it.saldo.toFixed(0)}</span>
          </div>
        </div>
      `;
  }).join('')}
  </div>`;
}

async function loadVacacionesCharts() {
  const rows = await loadObjectiveRows();
  renderCoverageChart(rows);
  renderBalanceChart(rows);
}

function log(msg, data) {
  if (data !== undefined) {
    console.info('[VAC]', msg, data);
  } else {
    console.info('[VAC]', msg);
  }
}

function setStatus(msg, isError = false) {
  const box = $('uiStatus');
  if (!box) return;
  box.textContent = msg || '';
  box.classList.remove('hidden');
  box.classList.toggle('err', !!isError);
}

function notify(msg, type = 'ok', timeoutMs = 4200) {
  const stack = $('toastStack');
  if (!stack || !msg) return;
  const item = document.createElement('div');
  item.className = `toast ${type}`;
  item.textContent = msg;
  stack.appendChild(item);
  window.setTimeout(() => {
    item.remove();
  }, timeoutMs);
}

function showCoverageModal(areaName) {
  const modal = $('kpiModal');
  $('kpiModalTitle').textContent = `Detalle: ${areaName} - Con Vacaciones`;
  $('tbKpiModal').innerHTML = '<tr><td colspan="9" style="text-align:center">Cargando...</td></tr>';
  modal.classList.remove('hidden');
  modal.style.display = 'flex';
  const rows = (state.objectiveRows || []).filter(r => (r.area || r.departamento || 'SIN AREA').trim() === areaName && r.tiene_vacaciones_q1);
  $('tbKpiModal').innerHTML = _renderRowsHtml(rows);
}

function showBalanceModal(areaName) {
  const modal = $('kpiModal');
  $('kpiModalTitle').textContent = `Detalle: ${areaName} - Saldo Vacacional`;
  $('tbKpiModal').innerHTML = '<tr><td colspan="9" style="text-align:center">Cargando...</td></tr>';
  modal.classList.remove('hidden');
  modal.style.display = 'flex';
  const rows = (state.objectiveRows || []).filter(r => (r.area || r.departamento || 'SIN AREA').trim() === areaName && (r.total_dias || 0) > 0);
  $('tbKpiModal').innerHTML = _renderRowsHtml(rows);
}

function setThemeFromSystem() {
  document.body.dataset.theme = 'light';
}
setThemeFromSystem();

async function jfetch(url, opts = {}) {
  const resp = await fetch(url, opts);
  const txt = await resp.text();
  let data = {};
  try {
    data = txt ? JSON.parse(txt) : {};
  } catch {
    // Algunas respuestas pueden incluir NaN/Infinity; se sanea para parseo JS.
    try {
      const sane = (txt || '')
        .replace(/\uFEFF/g, '')
        .replace(/:\s*NaN(?=[,}\]])/g, ': null')
        .replace(/:\s*Infinity(?=[,}\]])/g, ': null')
        .replace(/:\s*-Infinity(?=[,}\]])/g, ': null');
      data = sane ? JSON.parse(sane) : {};
    } catch {
      throw new Error((txt || '').slice(0, 180));
    }
  }
  if (!resp.ok || data.ok === false) throw new Error(data.error || `HTTP ${resp.status}`);
  return data;
}

function pickNombreCorreo() {
  const nombre = $('inpNombre').value.trim();
  let correo = $('inpCorreo').value.trim();
  if (!correo && nombre) {
    const entry = resolveNombreEntry(nombre);
    correo = (entry && entry.email) || '';
    if (correo) {
      $('inpCorreo').value = correo;
    }
  }
  if (!nombre || !correo || !correo.includes('@')) {
    throw new Error('Completa nombre y correo valido para previsualizar');
  }
  return { nombre, correo };
}

async function loadInit() {
  try {
    console.log('[DEBUG] loadInit() iniciando...');
    const data = await jfetch('/api/init');
    console.log('[DEBUG] Datos recibidos de /api/init:', data);
    const rows = await loadObjectiveRows();
    const triCtrl = data.trimestre_control || {};
    state.trimestreLabel = triCtrl.trimestre || (data.fuente_datos || {}).trimestre_vigente || 'Q?-????';
    state.trimestreModo = triCtrl.modo || (data.fuente_datos || {}).trimestre_modo || 'auto';
    applyTrimestreLabels(state.trimestreLabel);

    // Usar valores del resumen del backend
    const resumen = data.resumen || {};
    console.log('[DEBUG] Resumen:', resumen);
    console.log('[DEBUG] Actualizando KPIs con valores:', {
      total: resumen.total,
      con_vacaciones: resumen.con_vacaciones,
      con_saldo: resumen.con_saldo,
      sin_vacaciones: resumen.sin_vacaciones
    });
    renderKpiCards().catch(() => {});
    console.log('[DEBUG] KPIs (meta) actualizados');

    $('updatedAt').textContent = `Actualizado: ${new Date().toLocaleString()} | Trimestre: ${state.trimestreLabel} (${state.trimestreModo})`;
    const fuente = data.fuente_datos || {};
    if (fuente.archivo && !state.lastDataError) {
      const origen = fuente.origen ? ` (${fuente.origen})` : '';
      const triTxt = fuente.trimestre_vigente ? ` | Trimestre activo: ${fuente.trimestre_vigente}` : '';
      setStatus(`Sistema listo. Fuente activa: ${fuente.archivo}${origen}${triTxt}`);
    }
  } catch (e) {
    const msg = `Error cargando métricas: ${(e && e.message) || 'Error desconocido'}`;
    state.lastDataError = msg;
    setDashboardFallback(msg);
    setStatus(msg, true);
  }
}

async function loadRanking() {
  const elegVal = ($('fEleg') && $('fEleg').value) || 'todos';
  const rows = await fetchRankingRows(elegVal);
  if ($('tbRank')) {
    $('tbRank').innerHTML = _renderRowsHtml(rows);
    attachInlineEditListeners($('tbRank'));
  }

  const jefes = new Set(state.jefesRef.map(x => x.toUpperCase()));
  rows.forEach(r => {
    const sup = (r.hrbp || r.supervisor || '').trim();
    if (sup) jefes.add(sup.toUpperCase());
  });
  if ($('lstBuscar')) {
    $('lstBuscar').innerHTML = Array.from(jefes)
      .sort((a, b) => a.localeCompare(b, 'es'))
      .map(n => `<option value="${n}"></option>`)
      .join('');
  }
}

// Fetch con timeout para no bloquear indefinidamente (OneDrive puede tardar)
async function jfetchWithTimeout(url, opts = {}, timeoutMs = 15000) {
  const ctrl = new AbortController();
  const tid = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    return await jfetch(url, { ...opts, signal: ctrl.signal });
  } finally {
    clearTimeout(tid);
  }
}

async function loadSupervisores() {
  let data;
  const maxRetries = 3;
  const container = $('jefesArbol');
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      if (container && attempt > 1) {
        container.innerHTML = `<span class="spin"></span> Cargando jefes (intento ${attempt}/${maxRetries})...`;
      }
      // Timeout de 15s por intento para no quedar colgado si OneDrive tarda
      data = await jfetchWithTimeout('/api/jefes_equipo_arbol', {}, 15000);
      break; // success
    } catch (e) {
      const isTimeout = e && (e.name === 'AbortError' || (e.message || '').includes('abort'));
      const motivo = isTimeout ? 'El servidor tardó demasiado' : (e.message || 'Error de red');
      if (attempt < maxRetries) {
        const waitSec = 4 * attempt;
        console.log(`[VAC] loadSupervisores intento ${attempt} falló (${motivo}), reintentando en ${waitSec}s...`);
        if (container) {
          container.innerHTML = `<span class="spin"></span> Procesando datos... reintentando en ${waitSec}s`;
        }
        await new Promise(r => setTimeout(r, waitSec * 1000));
      } else {
        // Último intento fallido: mostrar botón de reintentar en lugar de error estático
        console.error('[VAC] loadSupervisores: todos los intentos fallaron');
        if (container) {
          container.innerHTML = `
            <div style="padding:14px;color:#5b6f87;text-align:center">
              <div style="font-size:13px;margin-bottom:10px">⚠ El servidor tardó en responder. Puedes intentarlo de nuevo.</div>
              <button class="btn btn-brand" onclick="loadSupervisores().then(()=>renderWizardRecipients())">↺ Reintentar</button>
            </div>`;
        }
        return;
      }
    }
  }
  const items = data.arbol || data.items || [];
  state.supervisoresData = items;

  const bps = new Set();
  const gers = new Set();
  const areas = new Set();
  state.supervisoresData.forEach(s => {
    if (s.es_vista_ejecutiva) return;
    const hrbps = Array.isArray(s.hrbps) ? s.hrbps : (s.hrbp ? [s.hrbp] : []);
    hrbps.forEach(bp => { if (bp) bps.add(bp); });
    if (s.gerencias) s.gerencias.forEach(g => gers.add(g));
    if (s.subgerencias) s.subgerencias.forEach(g => gers.add(g));
    if (s.areas) s.areas.forEach(a => areas.add(a.nombre || a));
  });

  if ($('wizFilterBP')) {
    const currentVal = $('wizFilterBP').value;
    $('wizFilterBP').innerHTML = '<option value="">Todos los BP</option>' +
      Array.from(bps).sort().map(bp => `<option value="${bp}">${bp}</option>`).join('');
    $('wizFilterBP').value = currentVal;
  }
  if ($('wizFilterGerencia')) {
    const currentVal = $('wizFilterGerencia').value;
    $('wizFilterGerencia').innerHTML = '<option value="">Todas las Gerencias</option>' +
      Array.from(gers).sort().map(g => `<option value="${g}">${g}</option>`).join('');
    $('wizFilterGerencia').value = currentVal;
  }
  if ($('wizFilterArea')) {
    const currentVal = $('wizFilterArea').value;
    $('wizFilterArea').innerHTML = '<option value="">Todas las Áreas</option>' +
      Array.from(areas).sort().map(a => `<option value="${a}">${a}</option>`).join('');
    $('wizFilterArea').value = currentVal;
  }

  state.selectedSupervisores = new Set(); // Default unselected as requested
  state.supMap = {};
  const jefesRef = [
    'CHANG CHANG, GABRIEL ANDRES',
    'LOLI PICON, CARMEN',
    'TORRES MANYARI, CARLOS JESUS',
    'REYES SANEZ, CESAR AUGUSTO',
    'VASQUEZ ROSAS, ROSARIO RAQUEL AURORA',
    'SUAREZ VERA, NANCY MERCEDES VICTORIA',
    'HIGA KANASHIRO, DEBORAH MELISSA',
    'REYES GUMEZ, LESLEY NOEMY',
    'SALAZAR FERNANDEZ, MARIA DE FATIMA'
  ];
  state.jefesRef = jefesRef;
  const opciones = [];
  if ($('lstJefes')) $('lstJefes').innerHTML = items.map(s => {
    state.supMap[(s.nombre || '').toLowerCase()] = s;
    opciones.push((s.nombre || '').toUpperCase());
    return `<option value="${s.nombre}"></option>`;
  }).join('');
  state.supervisorNames = items.map(s => (s.nombre || '').trim()).filter(Boolean);
  syncEditJefesFromSupervisores();

  const buscadorSet = new Set([...opciones, ...jefesRef.map(x => x.toUpperCase())]);
  if ($('lstBuscar')) {
    $('lstBuscar').innerHTML = Array.from(buscadorSet)
      .sort((a, b) => a.localeCompare(b, 'es'))
      .map(n => `<option value="${n}"></option>`)
      .join('');
  }
  renderDestinatarios($('inpDestFilter') ? $('inpDestFilter').value : '');
  refreshNombreSuggestions('');
  updateSummaryCards();
  log(`Jefes detectados: ${items.length}`);

  renderWizardRecipients();
}

function renderJefesArbol(arbol) {
  const container = $('jefesArbol');
  if (!container) return;

  let html = '';
  for (const jefe of arbol) {
    const jefe_id = 'jefe_' + jefe.nombre.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    const checked = state.selectedSupervisores.has(normalizeSupervisorName(jefe.nombre)) ? 'checked' : '';
    const gers = ((jefe.gerencias || []).concat(jefe.subgerencias || [])).join('|').toLowerCase();
    const areas = (jefe.areas || []).map(a => a.nombre || a).join('|').toLowerCase();

    html += `
      <div class="jefe-block" data-jefe-name="${jefe.nombre.toLowerCase()}" data-hrbp="${(jefe.hrbp || '').toLowerCase()}" data-gerencias="${gers}" data-areas="${areas}" style="margin-bottom:8px;border-left:3px solid #0f6ea5;padding-left:8px">
        <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px">
          <input type="checkbox" id="${jefe_id}_cb" ${checked} data-jefe="${jefe.nombre}" onchange="toggleJefe(this)" />
          <button class="expand-btn" id="${jefe_id}_btn" onclick="toggleJefeDetails('${jefe_id}')" style="background:none;border:none;cursor:pointer;padding:0;color:#0f6ea5;font-weight:bold;font-size:12px">▶</button>
          <span style="font-weight:700;color:#1b2a3d;font-size:14px">${jefe.nombre}</span>
          <span style="font-size:11px;color:#64748b">(${jefe.total_colaboradores} colab.)</span>
        </div>
        <div id="${jefe_id}_details" style="display:none;margin-top:6px;padding-left:12px;border-left:2px solid #e2e8f0">
    `;

    for (const area of jefe.areas || []) {
      const area_id = jefe_id + '_area_' + area.nombre.replace(/[^a-z0-9]/gi, '_').toLowerCase();
      html += `
        <div style="margin-bottom:6px;padding:6px;background:#f8fafc;border-radius:6px;border-left:2px solid #94a3b8">
          <div style="display:flex;align-items:center;gap:4px;margin-bottom:4px">
            <button class="expand-btn-area" id="${area_id}_btn" onclick="toggleAreaDetails('${area_id}')" style="background:none;border:none;cursor:pointer;padding:0;color:#475569;font-size:11px">▶</button>
            <span style="font-weight:600;color:#334155;font-size:13px">${area.nombre}</span>
            <span style="font-size:10px;color:#94a3b8">${area.total} personas</span>
          </div>
          <div id="${area_id}_details" style="display:none;padding:6px;background:#fff;border-radius:4px;margin-top:4px;font-size:12px">
      `;

      for (const p of area.personas || []) {
        html += `
          <div style="padding:4px;margin-bottom:2px;border-left:2px solid #cbd5e1;padding-left:8px;line-height:1.4">
            <div style="font-weight:600;color:#1b2a3d">${p.nombre}</div>
            <div style="font-size:11px;color:#64748b">
              ${p.gozados_q1}/${p.total_vacaciones} días | ${p.porcentaje_avance}% | Saldo: ${p.saldo}
            </div>
          </div>
        `;
      }

      html += `
          </div>
        </div>
      `;
    }

    html += `
        </div>
      </div>
    `;
  }

  container.innerHTML = html;
}

function toggleJefeDetails(jefe_id) {
  const btn = $(jefe_id + '_btn');
  const details = $(jefe_id + '_details');
  if (details) {
    details.style.display = details.style.display === 'none' ? 'block' : 'none';
    if (btn) btn.textContent = details.style.display === 'none' ? '▶' : '▼';
  }
}

function toggleAreaDetails(area_id) {
  const btn = $(area_id + '_btn');
  const details = $(area_id + '_details');
  if (details) {
    details.style.display = details.style.display === 'none' ? 'block' : 'none';
    if (btn) btn.textContent = details.style.display === 'none' ? '▶' : '▼';
  }
}

function toggleJefe(checkbox) {
  const jefe_nombre = checkbox.dataset.jefe;
  if (checkbox.checked) {
    state.selectedSupervisores.add(normalizeSupervisorName(jefe_nombre));
  } else {
    state.selectedSupervisores.delete(normalizeSupervisorName(jefe_nombre));
  }
  updateSummaryCards();
}

async function loadPersonas() {
  let data;
  try { data = await jfetch('/api/personas_autocomplete?limite=5000'); } catch(e) { log('loadPersonas error:', e.message); return; }
  if (!data) return;
  const items = data.items || [];
  state.personByName = {};
  state.personByMat = {};
  state.personOptions = [];

  for (const p of items) {
    const nombre = (p.nombre || '').trim();
    const email = (p.email || '').trim();
    const matricula = (p.matricula || '').trim();
    if (!nombre) continue;

    const key = nombre.toLowerCase();
    if (!state.personByName[key] || (!state.personByName[key].email && email)) {
      state.personByName[key] = { nombre, email, matricula };
    }
    if (matricula) state.personByMat[matricula] = { nombre, email, matricula };
    state.personOptions.push({ nombre, email, matricula, searchText: buildSearchText(nombre, matricula) });
  }

  refreshNombreSuggestions('');
  log(`Personas para autocompletar: ${items.length}`);
}

function renderPreview(data) {
  const frame = $('previewFrame');
  const teamsNode = $('previewTeams');
  const subjNode = $('previewSubj');
  const toNode = $('previewTo');
  const metaNode = $('previewMeta');
  const titleNode = $('previewModalTitle');
  if (frame) {
    frame.srcdoc = data.mensaje_html || '<html><body style="padding:18px;font-family:Segoe UI,Arial,sans-serif">Sin previsualizacion disponible.</body></html>';
  }
  if (teamsNode) {
    teamsNode.textContent = (data.mensaje_teams || 'Sin previsualización de Teams.').toString();
  }
  const camp = data.campania || {};
  const tramo = camp.trimestre ? ` | ${camp.trimestre}` : '';
  if (subjNode) subjNode.textContent = data.asunto || 'Alerta Vacaciones';
  if (toNode) toNode.textContent = `${data.nombre_jefe || '-'} <${data.email_jefe || '-'}>`;
  if (metaNode) metaNode.textContent = `Total: ${data.total_colaboradores || 0} | Retraso: ${data.en_retraso || 0} | Proximos: ${data.proximos || 0} | Sin cumplir: ${data.sin_cumplir || 0}${tramo}`;
  if (titleNode) titleNode.textContent = `Previsualizacion multicanal: ${data.nombre_jefe || 'Jefe'}`;
  openPreviewModal();
}

async function enviarGlobal(flags) {
  if (state.massSendInFlight) {
    throw new Error('Ya hay un envio masivo en proceso. Espera a que termine.');
  }
  if (!state.massArmed) {
    throw new Error('Envio masivo bloqueado. Debes aceptar primero el envio masivo.');
  }
  if (!flags.enviar_teams && !flags.enviar_smtp && !flags.encolar_pa) {
    throw new Error('Selecciona al menos un canal para la campaña.');
  }
  const selectedSupervisores = getEffectiveSelectedSupervisores();
  if (!selectedSupervisores.length) {
    throw new Error('No hay jefes seleccionados dentro de la vista actual para enviar.');
  }
  const btn = $('btnAllSend');
  state.massSendInFlight = true;
  if (btn) btn.disabled = true;
  try {
    setStatus('Enviando campaña masiva...');
    notify('Enviando campaña masiva, por favor espera...', 'warn', 2800);
      const currentFields = readCurrentTemplateFields();
    const data = await jfetch('/api/test-notif-teams', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          dias_proximos: 30,
          selected_supervisores: selectedSupervisores,
          hrbp_filtro: getActiveWizardHrbpFilter(),
          asunto: currentFields.asunto,
          mensaje: currentFields.mensaje,
          aviso: currentFields.aviso,
          recomendacion: currentFields.reco,
          ...flags
        })
    });
    log('Envio global', data);
    await Promise.allSettled([refreshQueueState(), refreshConfirmacionesResumen()]);
    const delay = (((data || {}).json_meta || {}).retraso_segundos || 60);
    const file = data.json_archivo || 'archivo pendiente';
    const paOkM = flags.encolar_pa && data && (data.json_guardado || data.json_archivo);
    const teamsOkM = flags.enviar_teams && data && data.teams_enviado;
    if (paOkM) {
      setStatus(`✅ Campaña enviada. Correos en camino vía Power Automate (~${delay}s).${teamsOkM ? ' Teams: OK.' : ''}`);
    } else if (teamsOkM) {
      setStatus(`✅ Campaña enviada vía Teams.`);
    } else {
      setStatus(`⚠️ Campaña enviada con advertencias. Revisa la cola PA.`);
    }
    notify(`Campaña enviada: ${file}. Seguridad activa (${delay}s).`, 'ok');
  } finally {
    state.massSendInFlight = false;
    wizardEvalSendBtn();
  }
}

function hydrateByName() {
  const raw = $('inpNombre').value.trim();
  const entry = resolveNombreEntry(raw);
  if (entry && entry.nombre && raw !== entry.nombre) {
    $('inpNombre').value = entry.nombre;
  }
  if (entry && entry.email) {
    $('inpCorreo').value = entry.email;
  }

  const s = entry && entry.supervisor ? entry.supervisor : state.supMap[raw.toLowerCase()];
  if (!s) return;
  if (!($('inpCorreo').value || '').trim()) {
    $('inpCorreo').value = s.email || '';
  }
  if (!$('txtMensaje').value) $('txtMensaje').value = s.mensaje || '';
  if (!$('txtAviso').value) $('txtAviso').value = s.aviso || '';
  if (!$('txtReco').value) $('txtReco').value = s.recomendacion || '';
}

function hydrateByTyping() {
  const raw = $('inpNombre').value.trim();
  refreshNombreSuggestions(raw);

  if (!raw) {
    $('inpCorreo').value = '';
    return;
  }

  const entry = resolveNombreEntry(raw);
  if (entry && entry.email) {
    $('inpCorreo').value = entry.email;
  }
}

async function previewInd() {
  const { nombre, correo } = pickNombreCorreo();
  const payload = {
    nombre,
    email: correo,
    asunto: $('txtAsunto') ? $('txtAsunto').value : '',
    mensaje: $('txtMensaje').value,
    aviso: $('txtAviso').value,
    recomendacion: $('txtReco').value,
    modo_prueba: $('chkModoPrueba') ? $('chkModoPrueba').checked : false
  };
  const data = await jfetch('/api/preview-supervisor', {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload)
  });
  renderPreview(data);
  setStatus(`Previsualizacion generada para ${data.nombre_jefe || nombre}. Total colaboradores: ${data.total_colaboradores || 0}.`);
  log('Preview individual listo', {
    nombre_jefe: data.nombre_jefe,
    email_jefe: data.email_jefe,
    email_destino_real: data.email_destino_real,
    modo_prueba: data.modo_prueba,
    total_colaboradores: data.total_colaboradores,
    en_retraso: data.en_retraso,
    proximos: data.proximos,
    sin_cumplir: data.sin_cumplir
  });
}

async function enviarInd(flags) {
  const { nombre, correo } = pickNombreCorreo();
  if (!flags.enviar_teams && !flags.enviar_smtp && !flags.encolar_para_pa) {
    throw new Error('Selecciona al menos un canal para el envío individual.');
  }
  setStatus(`Enviando mensaje a ${nombre}...`);
  notify(`Enviando mensaje a ${nombre}...`, 'warn', 2600);
  const payload = {
    nombre,
    email: correo,
    asunto: $('txtAsunto') ? $('txtAsunto').value : '',
    mensaje: $('txtMensaje').value,
    aviso: $('txtAviso').value,
    recomendacion: $('txtReco').value,
    modo_prueba: $('chkModoPrueba') ? $('chkModoPrueba').checked : false,
    ...flags
  };
  const data = await jfetch('/api/enviar-a-supervisor', {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload)
  });
  if (data && data.ok) {
    await previewInd();
  }
  const delay = Number(data.retraso_seguridad_segundos || (((data || {}).json_meta || {}).retraso_segundos || 60));
  const infoFile = data.archivo || 'sin archivo';
  const paOkI = flags.encolar_para_pa && data.json_guardado;
  const teamsOkI = flags.enviar_teams && data.teams_enviado;
  const smtpOkI = flags.enviar_smtp && data.smtp_enviado;
  const pruebaTxt = data.modo_prueba ? ' (modo prueba)' : '';
  if (paOkI) {
    setStatus(`✅ Correo en camino para ${nombre}${pruebaTxt}. Power Automate lo procesará en ~${delay}s.${teamsOkI ? ' Teams: OK.' : ''}`);
  } else if (teamsOkI || smtpOkI) {
    setStatus(`✅ Envío ejecutado para ${nombre}${pruebaTxt}.`);
  } else {
    const teamsErrI = flags.enviar_teams && !data.teams_enviado ? ` Teams: ${data.teams_error || 'sin webhook'}` : '';
    setStatus(`⚠️ Envío para ${nombre} con advertencias.${teamsErrI}${pruebaTxt}`);
  }
  notify(`Mensaje enviado para ${nombre}. Se procesa en ${delay}s (${infoFile}).`, 'ok');
  await Promise.allSettled([refreshQueueState(), refreshConfirmacionesResumen()]);
  log('Envio individual', data);
}

async function boot() {
  console.log('[DEBUG] boot() called');

  // Disparar KPIs en paralelo con /api/init para que el cache del servidor
  // se caliente y los contadores aparezcan tan pronto como sea posible.
  const kpiPromise = renderKpiCards().catch(() => {});

  // CARGAR KPIS CON REINTENTOS - el servidor puede estar precargando datos
  let initData = null;
  const maxInitRetries = 3;
  for (let attempt = 1; attempt <= maxInitRetries; attempt++) {
    try {
      if (attempt > 1) {
        if ($('updatedAt')) $('updatedAt').textContent = `Conectando al servidor (intento ${attempt}/${maxInitRetries})...`;
      }
      console.log(`[DEBUG] Cargando KPIs (intento ${attempt})...`);
      initData = await jfetch('/api/init');
      loadModoEnvio().catch(() => {});
      break;
    } catch (e) {
      console.warn(`[BOOT] /api/init intento ${attempt} falló:`, e.message);
      if (attempt < maxInitRetries) {
        const waitSec = 3 * attempt;
        if ($('updatedAt')) $('updatedAt').textContent = `Servidor cargando datos... reintentando en ${waitSec}s`;
        await new Promise(r => setTimeout(r, waitSec * 1000));
      } else {
        console.error('[ERROR] Cargando KPIs:', e);
        const msg = 'Error cargando KPIs: ' + e.message;
        state.lastDataError = msg;
        setDashboardFallback(msg);
        setStatus(msg, true);
      }
    }
  }

  if (initData) {
    const resumen = initData.resumen || {};
    console.log('[DEBUG] Resumen recibido:', resumen);

    // Esperar a que los KPIs terminen (incluye sus reintentos internos)
    await kpiPromise;
    console.log('[DEBUG] KPIs (meta) actualizados en DOM');

    // Actualizar etiquetas de trimestre
    const triCtrl = initData.trimestre_control || {};
    state.trimestreLabel = triCtrl.trimestre || (initData.fuente_datos || {}).trimestre_vigente || 'Q?-????';
    state.trimestreModo = triCtrl.modo || (initData.fuente_datos || {}).trimestre_modo || 'auto';
    applyTrimestreLabels(state.trimestreLabel);
    if ($('updatedAt')) {
      $('updatedAt').textContent = `Actualizado: ${new Date().toLocaleString()} | Trimestre: ${state.trimestreLabel} (${state.trimestreModo})`;
    }

    const fuente = initData.fuente_datos || {};
    if (fuente.archivo) {
      const origen = fuente.origen ? ` (${fuente.origen})` : '';
      const triTxt = fuente.trimestre_vigente ? ` | Trimestre activo: ${fuente.trimestre_vigente}` : '';
      setStatus(`Sistema listo. Fuente activa: ${fuente.archivo}${origen}${triTxt}`);
    }
  }

  // Fase 1: datos rápidos de la pantalla inicial; personas se carga en segundo plano.
  await Promise.allSettled([loadObjectiveRows()]);
  window.setTimeout(() => {
    loadPersonas().catch(e => log(e.message));
  }, 0);
  // Fase 3: ranking inicial (jefesRef aún vacío, se actualiza cuando lleguen supervisores)
  loadRanking().catch(e => log(e.message));
  // Supervisores en background — lento (7-8s); refresca ranking + cards cuando termina
  loadSupervisores().catch(() => {}).then(() => {
    loadRanking().catch(e => log(e.message));
    updateSummaryCards();
  });
  await loadTrimestreControl().catch(() => { });
  refreshColabSuggestions('');
  window.setTimeout(() => {
    loadEditableRowsAutocomplete().catch(e => log(e.message));
  }, 120);
  setMassArmed(false);
  await Promise.allSettled([refreshQueueState(), refreshConfirmacionesResumen()]);
  updateSummaryCards();
  if (!state.lastDataError) {
    setStatus('Sistema listo y conectado a fuentes de datos.');
  }
  log('Front cargado con base Talento y Cultura');

  // Refresco periódico de KPIs cada 60s para mantener predicción y avance actualizados
  setInterval(() => {
    renderKpiCards(1, 0).catch(() => {});  // 1 intento, sin espera extra
  }, 60000);
}

// --- Configuración de Carpeta de Red ---
async function loadConfigCola() {
  const data = await jfetch('/api/config/cola_dir');
  updateConfigColaView(data);
}

function getConfigColaInput() {
  return $('cfgDir') || $('cfgColaDir');
}

function getConfigColaModal() {
  return $('cfgModal') || $('configModal');
}

function openConfigColaModal() {
  const modal = getConfigColaModal();
  if (!modal) return;
  modal.classList.remove('hidden');
  modal.style.display = 'flex';
}

function closeConfigColaModal() {
  const modal = getConfigColaModal();
  if (!modal) return;
  modal.classList.add('hidden');
  modal.style.display = '';
}

async function saveConfigCola() {
  const input = getConfigColaInput();
  const newDir = String((input && input.value) || '').trim();
  if (!newDir) return notify('Escribe una ruta valida', 'err');
  const data = await jfetch('/api/config/cola_dir', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ dir: newDir })
  });
  if (data.ok) {
    notify('Configuración guardada.', 'ok');
    updateConfigColaView(data, 'Carpeta actualizada correctamente.');
  }
}

async function openConfigColaDir(ev) {
  if (ev) {
    ev.preventDefault();
  }
  const input = getConfigColaInput();
  const dir = String((input && input.value) || ($('cfgCurrentDir') && $('cfgCurrentDir').textContent) || '').trim();
  if (!dir) {
    notify('No hay carpeta configurada para abrir.', 'warn');
    return;
  }
  const data = await jfetch('/api/config/cola_dir/open', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ dir })
  });
  if (data && data.ok) {
    notify('Carpeta local abierta.', 'ok');
  }
}

async function pickConfigColaDir() {
  const input = getConfigColaInput();
  const currentDir = String((input && input.value) || ($('cfgCurrentDir') && $('cfgCurrentDir').textContent) || '').trim();
  const data = await jfetch('/api/config/cola_dir/pick', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ dir: currentDir })
  });
  if (data && data.cancelled) {
    return;
  }
  if (input && data && data.dir) {
    input.value = data.dir;
    input.title = data.dir;
  }
  if ($('cfgStatus')) {
    $('cfgStatus').textContent = 'Carpeta seleccionada. Pulsa Guardar Configuración para aplicarla.';
  }
}

function updateConfigColaView(data, msgOk = '') {
  const dir = String((data || {}).dir || '').trim();
  const exists = Boolean((data || {}).exists);
  const input = getConfigColaInput();
  if (input) {
    input.value = dir;
    input.title = dir;
  }
  if ($('cfgCurrentDir')) {
    $('cfgCurrentDir').textContent = dir || 'No configurada';
    $('cfgCurrentDir').title = dir || 'No configurada';
  }
  if ($('cfgOpenLink')) {
    if (dir) {
      $('cfgOpenLink').href = '#';
      $('cfgOpenLink').classList.remove('disabled');
    } else {
      $('cfgOpenLink').href = '#';
      $('cfgOpenLink').classList.add('disabled');
    }
  }
  if ($('cfgStatus')) {
    const estado = dir ? (exists ? 'Carpeta detectada y lista para usar.' : 'La ruta se guardó, pero la carpeta aún no existe.') : 'Configura una carpeta para usar la cola compartida.';
    $('cfgStatus').textContent = msgOk || estado;
  }
}

bind('btnConfigCola', 'click', () => {
  openConfigColaModal();
  loadConfigCola()
    .then(() => pickConfigColaDir())
    .catch(e => {
      setStatus(e.message, true);
      notify(e.message || 'No se pudo abrir el selector de carpeta.', 'err');
    });
});
bind('btnCloseCfg', 'click', () => {
  closeConfigColaModal();
});
bind('btnSaveCfg', 'click', saveConfigCola);
bind('btnPickCfgDir', 'click', () => pickConfigColaDir().catch(e => {
  setStatus(e.message, true);
  notify(e.message || 'No se pudo seleccionar la carpeta.', 'err');
}));
bind('cfgOpenLink', 'click', (ev) => openConfigColaDir(ev).catch(e => {
  setStatus(e.message, true);
  notify(e.message || 'No se pudo abrir la carpeta.', 'err');
}));

bind('btnPrev', 'click', () => previewInd().catch(e => setStatus(e.message, true)));
bind('btnIndSend', 'click', () => enviarInd(getSelectedIndividualChannels()).catch(e => setStatus(e.message, true)));

bind('btnArmMass', 'click', () => {
  const total = getEffectiveSelectedSupervisores().length;
  if (!total) {
    setStatus('Selecciona al menos un jefe para envio masivo.', true);
    return;
  }
  const ok = window.confirm(`Aceptas habilitar envio masivo para ${total} jefes seleccionados?`);
  if (!ok) return;
  setMassArmed(true);
  setStatus('Envio masivo habilitado para esta sesion.');
});
bind('btnCheckQueue', 'click', () => refreshQueueState().then(() => setStatus('Estado de cola actualizado.')).catch(e => setStatus(e.message, true)));
bind('btnSendNow', 'click', enviarTodoAhora);
bind('btnGlobalSendNow', 'click', enviarTodoAhora);
bind('btnCancelQueue', 'click', () => cancelarUltimoPendiente().catch(e => setStatus(e.message, true)));
bind('btnGlobalCancelQueue', 'click', () => cancelarUltimoPendiente().catch(e => {
  setStatus(e.message, true);
  notify(e.message || 'No se pudo cancelar.', 'err');
}));

window.setInterval(refreshQueueState, 1000);
bind('btnSearchEdit', 'click', () => loadEditableRows().catch(e => setStatus(e.message, true)));
if ($('inpDestFilter')) $('inpDestFilter').addEventListener('input', () => renderDestinatarios($('inpDestFilter').value));
if ($('btnSelAll')) $('btnSelAll').addEventListener('click', () => selectVisibleDestinatarios(true));
if ($('btnSelNone')) $('btnSelNone').addEventListener('click', () => selectVisibleDestinatarios(false));

bind('btnToggleAdv', 'click', () => $('advPanel').classList.toggle('hidden'));
bind('btnTriApply', 'click', () => applyTrimestreControl().catch(e => setStatus(e.message, true)));

bind('inpUpload', 'change', async function () {
  const file = this.files[0];
  if (!file) return;
  const fd = new FormData();
  fd.append('excel', file, file.name);
  setStatus('Subiendo archivo...');
  try {
    const r = await fetch('/api/subir_excel', { method: 'POST', body: fd });
    const d = await r.json();
    if (d.ok) {
      const detalleTipo = d.tipo_carga ? ` (${d.tipo_carga}${d.trimestre ? ` ${d.trimestre}` : ''})` : '';
      notify(`Archivo subido: ${d.archivo}${detalleTipo}`, 'ok');
      setStatus(`Datos actualizados${detalleTipo}. Recargando...`);
      await loadInit();
      await loadSupervisores();
      await loadPersonas();
      await loadRanking().catch(() => { });
    } else {
      notify(d.error || 'Error al subir archivo', 'err');
      setStatus(d.error || 'Error', true);
    }
  } catch (e) {
    notify(e.message, 'err');
  } finally {
    this.value = '';
  }
});

bind('fBuscar', 'input', () => loadRanking().catch(e => log(e.message)));
bind('fEleg', 'change', () => {
  loadRanking().catch(e => log(e.message));
  updateSummaryCards();
});
bind('inpNombre', 'change', hydrateByName);
bind('inpNombre', 'input', hydrateByTyping);
bind('kpiTotal', 'click', () => openKpiModalByFilter('con_meta'));
bind('kpiEleg', 'click', () => openKpiModalByFilter('cumplieron'));
bind('kpiSinIniciar', 'click', () => openKpiModalByFilter('sin_iniciar'));
bind('kpiParciales', 'click', () => openKpiModalByFilter('parciales'));
bind('kpiSinMeta', 'click', () => openKpiModalByFilter('sin_meta_con_vac'));
bind('btnCloseKpiModal', 'click', closeKpiModal);
bind('kpiModal', 'click', (ev) => {
  if (ev.target === $('kpiModal')) closeKpiModal();
});
bind('btnClosePreviewModal', 'click', closePreviewModal);
bind('previewModal', 'click', (ev) => {
  if (ev.target === $('previewModal')) closePreviewModal();
});

// ====== COLABORADORES ======
function refreshColabSuggestions(query) {
  const q = normTxt(query || '');
  const items = state.personOptions
    .filter(p => !q || matchByTokens(p.searchText || '', q))
    .slice(0, 120);
  const dl = $('lstColab');
  if (dl) {
    dl.innerHTML = items.map(p =>
      `<option value="${(p.nombre || '').replace(/"/g, '&quot;')}" data-mat="${p.matricula || ''}"></option>`
    ).join('');
  }
}

async function buscarColaborador() {
  const buscar = $('colab_buscar').value.trim();
  if (!buscar) return notify('Ingresa matrícula o nombre', 'warn');

  // Intentar resolver matrícula desde datalist si el texto coincide con un nombre
  let query = buscar;
  const pExacta = state.personByName[buscar.toLowerCase()];
  if (pExacta && pExacta.matricula) query = pExacta.matricula;

  try {
    const encodedQuery = encodeURIComponent(query);
    const resp = await fetch(`/api/colaborador/${encodedQuery}`);
    if (!resp.ok) throw new Error((await resp.json()).error || 'No encontrado');
    const data = await resp.json();
    const p = data.persona;
    const pl = data.plantilla;

    $('colab_nombre').textContent = p.nombre || '-';
    $('colab_bp').textContent = p.hrbp || '-';
    $('colab_pendientes').textContent = (p.pendientes || 0).toFixed(1) + ' días';
    $('colab_vencidas').textContent = (p.vencidas || 0).toFixed(1) + ' días';

    $('colab_asunto').value = pl.asunto;
    $('colab_cuerpo').value = pl.cuerpo;

    $('colab_estado_box').classList.remove('hidden');
    $('colab_buscar').dataset.matricula = p.matricula || query;
    $('colab_mensaje_status').style.display = 'none';

    notify(`${p.nombre} cargado`, 'ok');
  } catch (e) {
    notify(e.message || 'Error al buscar', 'err');
  }
}

async function enviarAColaborador() {
  const mat = $('colab_buscar').dataset.matricula;
  if (!mat) return notify('Primero busca un colaborador', 'warn');

  const enviarTeams = !!$('colabChTeams').checked;
  const enviarEmail = !!$('colabChSmtp').checked;
  const modoPrueba = !!$('colabModoPrueba').checked;

  if (!enviarTeams && !enviarEmail) {
    return notify('Selecciona al menos un canal de envío', 'warn');
  }

  const btnEnviar = $('btnColab_Enviar');
  btnEnviar.disabled = true;
  btnEnviar.textContent = 'Enviando...';

  try {
    const resp = await fetch('/api/enviar-a-colaborador', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        matricula: mat,
        asunto: $('colab_asunto').value,
        cuerpo: $('colab_cuerpo').value,
        enviar_teams: enviarTeams,
        enviar_email: enviarEmail,
        modo_prueba: modoPrueba
      })
    });
    if (!resp.ok) throw new Error((await resp.json()).error || 'Error en envío');
    const data = await resp.json();

    const statusEl = $('colab_mensaje_status');
    statusEl.textContent = data.mensaje || `Procesado para ${data.nombre}.`;
    statusEl.style.display = 'block';
    statusEl.style.borderLeftColor = '#16a34a';
    statusEl.style.background = '#f0fdf4';
    statusEl.style.color = '#166534';

    notify(data.mensaje || 'Enviado correctamente', 'ok', 5000);
  } catch (e) {
    const statusEl = $('colab_mensaje_status');
    statusEl.textContent = e.message || 'Error al enviar';
    statusEl.style.display = 'block';
    statusEl.style.borderLeftColor = '#d9480f';
    statusEl.style.background = '#fff5f5';
    statusEl.style.color = '#9f1239';
    notify(e.message || 'Error al enviar', 'err');
  } finally {
    btnEnviar.disabled = false;
    btnEnviar.textContent = 'Enviar mensaje';
  }
}

function restaurarPlantilla() {
  const mat = $('colab_buscar').dataset.matricula;
  if (!mat) return;
  buscarColaborador();
}

bind('btnColab_Buscar', 'click', buscarColaborador);
bind('btnColab_Enviar', 'click', enviarAColaborador);
bind('btnColab_Restaurar', 'click', restaurarPlantilla);
bind('colab_buscar', 'keypress', (e) => {
  if (e.key === 'Enter') buscarColaborador();
});
bind('colab_buscar', 'input', () => refreshColabSuggestions($('colab_buscar').value));

// Ejecutar boot cuando el DOM esté listo
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    boot().catch(e => {
      console.error('[BOOT ERROR]', e);
      notify(e.message || 'Error al iniciar el sistema.', 'err');
    });
  });
} else {
  boot().catch(e => {
    console.error('[BOOT ERROR]', e);
    notify(e.message || 'Error al iniciar el sistema.', 'err');
  });
}

// ====================================================
// WIZARD UX — Funciones nuevas (no reemplazan ninguna)
// ====================================================

// --- Navegación entre pasos ---
let _wizardStep = 1;
function wizardGo(step) {
  _wizardStep = step;
  [1, 2, 3].forEach(n => {
    const panel = document.getElementById('wpanel' + n);
    const btn = document.getElementById('wstep' + n) || document.getElementById('step' + n);
    if (panel) panel.classList.toggle('visible', n === step);
    if (btn) {
      btn.classList.toggle('active', n === step);
      btn.classList.toggle('done', n < step);
    }
  });
  if (step === 3) wizardRefreshResumen();
}

// --- Templates de mensaje ---
const _TEMPLATES = {
  recordatorio: {
    nombre: '📋 Programación 2T-2026 (Oficial)',
    asunto: 'Meta Vacaciones USIL {trimestre}',
    mensaje: `Hola, {nombre}:

Desde la Gerencia de Talento y Cultura, seguimos impulsando la eficiencia institucional y el bienestar de nuestros equipos. Para lograrlo, la gestión oportuna del descanso es fundamental.

Hemos compartido contigo el Reporte de Vacaciones de las personas a tu cargo, donde se detalla la meta de días a programar para este trimestre. Para visualizar este reporte, te solicitamos revises tu bandeja de entrada de tu cuenta de correo institucional.

📌 <b>En el reporte encontrarás:</b>

<ul>
<li style="margin-bottom:8px"><b>Meta de Vacaciones para el {trimestre}:</b> El reporte indica una meta específica para cada colaborador. Si algún miembro de tu equipo no tiene una meta establecida, eso no es impedimento para que pueda programar días de descanso si fuera el caso.</li>
<li style="margin-bottom:8px"><b>Proyección de Saldo:</b> Es importante precisar que el saldo de vacaciones mostrado en el cuadro enviado por correo electrónico es una proyección de los días ganados por el colaborador calculada hasta el 31 de julio de 2026.</li>
<li style="margin-bottom:8px"><b>Vacaciones Corporativas:</b> El periodo del 20 de julio al 02 de agosto ha sido definido como prioridad para la Institución, debido a la suspensión de actividades académicas.</li>
<li style="margin-bottom:8px"><b>Cumplimiento:</b> Si la meta trimestral de un colaborador es superior a los 14 días del periodo de vacaciones corporativas, es necesario programar los días restantes dentro del presente trimestre para alcanzar el objetivo asignado.</li>
</ul>

<span style="color:#b91c1c"><b>Importante:</b> Agradeceremos completar esta programación y las aprobaciones correspondientes en el sistema <b>lo más pronto posible</b>.</span>

Atentamente,
<b><span style="background-color:#fef08a">{hrbp_nombre}</span></b>
Subgerencia de Talento y Cultura
<i>"Formar personas con valores para que dejen su huella en el mundo"</i>`,
    aviso: '',
    reco: ''
  },
  urgente: {
    nombre: '⚠️ Vacaciones por vencer',
    asunto: 'URGENTE: Vacaciones por vencer en tu equipo',
    mensaje: `Hola, {nombre}:

⚠️ ATENCIÓN: Tienes colaboradores con días de vacaciones próximos a vencer este trimestre. Es necesario programarlos a la brevedad para evitar que se pierdan o generen contingencias laborales.

Te adjuntamos el detalle de tu equipo con los saldos pendientes. Por favor, coordina con cada colaborador y registra las fechas en Adryan a la brevedad posible.

Atentamente,
{hrbp}
Subgerencia de Talento y Cultura`,
    aviso: 'Los días de vacaciones vencidos pueden generar observaciones en auditoría de RRHH.',
    reco: 'Prioriza la programación de los colaboradores con mayor saldo pendiente y días más próximos a vencer.'
  },
  positivo: {
    nombre: '🌟 Felicitación y seguimiento',
    asunto: 'Excelente avance en programación de vacaciones',
    mensaje: `Hola, {nombre}:

🌟 ¡Excelente avance! Tu equipo está cumpliendo de manera destacada con el plan de vacaciones del trimestre. Desde la Gerencia de Talento y Cultura, reconocemos y valoramos este compromiso.

Te compartimos el reporte actualizado de tu equipo. Te pedimos continuar impulsando la programación de los colaboradores que aún tienen días pendientes.

Atentamente,
<b><span style="background-color:#fef08a">{hrbp_nombre}</span></b>
Subgerencia de Talento y Cultura
"Formar personas con valores para que dejen su huella en el mundo"`,
    aviso: 'Quedan algunos colaboradores con días pendientes de programar en el trimestre.',
    reco: 'Mantén el buen avance coordinando los casos restantes lo más pronto posible.'
  },
  personalizado: {
    nombre: '✏️ Mensaje personalizado',
    asunto: '',
    mensaje: '',
    aviso: '',
    reco: ''
  }
};

let _currentTemplate = 'recordatorio';
let _customTemplateDraft = {
  asunto: '',
  mensaje: '',
  aviso: '',
  reco: ''
};

function readCurrentTemplateFields() {
  return {
    asunto: ((document.getElementById('txtAsunto') || {}).value || '').trim(),
    mensaje: ((document.getElementById('txtMensaje') || {}).value || ''),
    aviso: ((document.getElementById('txtAviso') || {}).value || ''),
    reco: ((document.getElementById('txtReco') || {}).value || '')
  };
}

function hasTemplateContent(data) {
  if (!data) return false;
  return Boolean(String(data.asunto || '').trim() || String(data.mensaje || '').trim() || String(data.aviso || '').trim() || String(data.reco || '').trim());
}

function decodeHtmlEntities(value) {
  const node = document.createElement('textarea');
  node.innerHTML = String(value || '');
  return node.value;
}

function htmlTemplateToEditableText(value) {
  const raw = String(value || '');
  if (!/[<>]/.test(raw)) return raw;
  const normalized = raw
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>\s*<p[^>]*>/gi, '\n\n')
    .replace(/<p[^>]*>/gi, '')
    .replace(/<\/p>/gi, '')
    .replace(/<li[^>]*>/gi, '• ')
    .replace(/<\/li>/gi, '\n')
    .replace(/<\/?(?:ul|ol)[^>]*>/gi, '\n')
    .replace(/<\/?(?:div|span)[^>]*>/gi, '')
    .replace(/<\/?(?:b|strong|i|em|u)[^>]*>/gi, '')
    .replace(/<[^>]+>/g, '');
  return decodeHtmlEntities(normalized)
    .replace(/\r/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function makeEditableDraft(data) {
  const next = data || {};
  return {
    asunto: String(next.asunto || ''),
    mensaje: htmlTemplateToEditableText(next.mensaje || ''),
    aviso: htmlTemplateToEditableText(next.aviso || ''),
    reco: htmlTemplateToEditableText(next.reco || '')
  };
}

function writeCurrentTemplateFields(data) {
  const next = data || {};
  const asu = document.getElementById('txtAsunto');
  const msg = document.getElementById('txtMensaje');
  const avi = document.getElementById('txtAviso');
  const rec = document.getElementById('txtReco');
  if (asu) asu.value = next.asunto || '';
  if (msg) msg.value = next.mensaje || '';
  if (avi) avi.value = next.aviso || '';
  if (rec) rec.value = next.reco || '';
}

function setCustomEditorVisible(show) {
  const editor = $('customMsgEditor');
  if (!editor) return;
  editor.classList.toggle('hidden', !show);
  editor.setAttribute('aria-hidden', show ? 'false' : 'true');
}

function syncCustomTemplateDraft() {
  _customTemplateDraft = readCurrentTemplateFields();
}

function selectTemplate(card, key) {
  const previousFields = readCurrentTemplateFields();
  if (_currentTemplate === 'personalizado') {
    _customTemplateDraft = previousFields;
  }
  document.querySelectorAll('.template-card, .tpl').forEach(c => c.classList.remove('active'));
  if (!card) return;
  card.classList.add('active');
  _currentTemplate = key;
  const tpl = _TEMPLATES[key] || _TEMPLATES.recordatorio;
  if (key === 'personalizado') {
    if (!hasTemplateContent(_customTemplateDraft)) {
      _customTemplateDraft = hasTemplateContent(previousFields)
        ? makeEditableDraft(previousFields)
        : makeEditableDraft({
            asunto: _TEMPLATES.recordatorio.asunto || '',
            mensaje: _TEMPLATES.recordatorio.mensaje || '',
            aviso: _TEMPLATES.recordatorio.aviso || '',
            reco: _TEMPLATES.recordatorio.reco || ''
          });
    }
    writeCurrentTemplateFields(_customTemplateDraft);
  } else {
    writeCurrentTemplateFields({
      asunto: tpl.asunto || '',
      mensaje: tpl.mensaje || '',
      aviso: tpl.aviso || '',
      reco: tpl.reco || ''
    });
  }
  setCustomEditorVisible(key === 'personalizado');
  updateMsgPreview();
  const prev = $('msgPreviewBox');
  if (prev) prev.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  if (key === 'personalizado') {
    const msgField = document.getElementById('txtMensaje');
    if (msgField) msgField.focus();
  }
}

function escapePreviewHtml(value) {
  return (value || '')
    .toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getPreviewFieldElement(fieldName) {
  const fieldMap = {
    mensaje: 'txtMensaje',
    aviso: 'txtAviso',
    reco: 'txtReco'
  };
  return document.getElementById(fieldMap[fieldName] || '');
}

function readPreviewEditableValue(node) {
  if (!node) return '';
  const fieldName = node.getAttribute('data-edit-field');
  if (fieldName === 'mensaje') {
    return (node.innerHTML || '').trim();
  }
  return ((node.innerText) || '').replace(/\r/g, '').replace(/\u00a0/g, ' ');
}

function attachPreviewInlineEditing(box) {
  if (!box || box.dataset.inlineEditBound === '1') return;

  box.addEventListener('dblclick', (event) => {
    const editable = event.target.closest('[data-edit-field]');
    if (!editable) return;
    editable.setAttribute('contenteditable', 'true');
    editable.classList.add('editing');
    editable.focus();
    const range = document.createRange();
    range.selectNodeContents(editable);
    range.collapse(false);
    const selection = window.getSelection();
    if (selection) {
      selection.removeAllRanges();
      selection.addRange(range);
    }
  });

  box.addEventListener('input', (event) => {
    const editable = event.target.closest('[data-edit-field]');
    if (!editable) return;
    const field = getPreviewFieldElement(editable.getAttribute('data-edit-field'));
    if (!field) return;
    field.value = readPreviewEditableValue(editable);
  });

  box.addEventListener('blur', (event) => {
    const editable = event.target.closest('[data-edit-field]');
    if (!editable) return;
    const field = getPreviewFieldElement(editable.getAttribute('data-edit-field'));
    const nextValue = readPreviewEditableValue(editable);
    if (field) field.value = nextValue;
    editable.removeAttribute('contenteditable');
    editable.classList.remove('editing');
    updateMsgPreview();
  }, true);

  box.addEventListener('keydown', (event) => {
    const editable = event.target.closest('[data-edit-field]');
    if (!editable) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      editable.blur();
    }
  });

  box.dataset.inlineEditBound = '1';
}

function updateMsgPreview() {
  const box = $('msgPreviewBox');
  if (!box) return;
  let msg = (document.getElementById('txtMensaje') || {}).value || '';
  let avi = (document.getElementById('txtAviso') || {}).value || '';
  let rec = (document.getElementById('txtReco') || {}).value || '';
  if (_currentTemplate === 'personalizado') {
    syncCustomTemplateDraft();
  }

  // Placeholder replacement for realistic preview
  const firstJefe = getEffectiveSelectedSupervisores()[0] || '';
  const firstPersona = (state.selectedPersonasObjetivoRows && state.selectedPersonasObjetivoRows[0]) || null;
  const typedNombre = String((($('inpNombre') && $('inpNombre').value) || '')).trim();
  const previewNombre = String((firstPersona && firstPersona.nombre) || firstJefe || typedNombre || 'Líder').trim();
  const hrbpName = String((firstPersona && firstPersona.hrbp) || 'People Analytics - USIL').trim();
  const rawTrimestre = String(state.trimestreLabel || '').trim();
  const updatedAtText = String((($('updatedAt') && $('updatedAt').textContent) || '')).trim();
  const updatedAtMatch = updatedAtText.match(/Trimestre:\s*([^|]+)/i);
  const trimestreActual = String(
    (rawTrimestre && rawTrimestre !== 'Q?-????' && rawTrimestre !== 'Trimestre vigente' && rawTrimestre) ||
    (updatedAtMatch && updatedAtMatch[1] && updatedAtMatch[1].trim()) ||
    (($('triManual') && $('triManual').value) || '').trim() ||
    'trimestre vigente'
  ).trim();
  const fechaLimite = 'la fecha acordada';

  const replaceAll = (t) => t
    .replace(/{nombre}/g, previewNombre)
    .replace(/{hrbp_nombre}/g, hrbpName)
    .replace(/{hrbp}/g, hrbpName)
    .replace(/{trimestre}/g, trimestreActual)
    .replace(/{fecha_limite}/g, fechaLimite);
  msg = replaceAll(msg);
  avi = replaceAll(avi);
  rec = replaceAll(rec);

  const safeAvi = escapePreviewHtml(avi);
  const safeRec = escapePreviewHtml(rec);

  if (!String(msg).trim() && !String(avi).trim() && !String(rec).trim()) {
    box.innerHTML = `
      <div class="preview-empty-state">
        <div class="preview-empty-title">Empieza a escribir tu mensaje personalizado</div>
        <div class="preview-empty-hint">Usa el editor de abajo para redactar el cuerpo, el aviso y la recomendación. La vista previa aparecerá aquí en tiempo real.</div>
      </div>`;
    return;
  }

  const content = `
    <div class="preview-editable preview-editable-main" data-edit-field="mensaje" style="font-family:inherit;">${msg}</div>
    ${avi ? `<div class="preview-editable preview-editable-note" data-edit-field="aviso" style="margin-top:12px; padding:8px; background:#fff7ed; border-left:4px solid #f97316; font-size:13px; white-space: pre-wrap;">📌 <b>Aviso:</b> ${safeAvi}</div>` : ''}
    ${rec ? `<div class="preview-editable preview-editable-note" data-edit-field="reco" style="margin-top:12px; padding:8px; background:#f0f9ff; border-left:4px solid #3b82f6; font-size:13px; white-space: pre-wrap;">💡 <b>Recomendación:</b> ${safeRec}</div>` : ''}
  `;
  box.innerHTML = content || '<span style="color:#94a3b8 italic">Selecciona una plantilla para ver cómo se verá el mensaje.</span>';
  attachPreviewInlineEditing(box);
}

// --- Resumen del paso 3 ---
function wizardRefreshResumen() {
  const mode = getWizardRecipientMode();
  const selectedSupervisores = getEffectiveSelectedSupervisores();
  const selectedPersonRows = getSelectedPersonRows();
  const count = mode === 'personas' ? selectedPersonRows.length : selectedSupervisores.length;
  const el3c = $('w3DestCount');
  const el3n = $('w3DestNames');
  const el3t = $('w3TplName');
  const el3s = $('w3MsgSnippet');
  const elBtn = $('w3BtnCount');
  if (el3c) el3c.textContent = mode === 'personas'
    ? count + ' persona' + (count !== 1 ? 's' : '')
    : count + ' jefe' + (count !== 1 ? 's' : '');
  if (el3n) {
    const names = mode === 'personas'
      ? selectedPersonRows.slice(0, 3).map(r => r.nombre || '-')
      : selectedSupervisores.slice(0, 3);
    el3n.textContent = names.join(', ') + (count > 3 ? ` y ${count - 3} más` : '');
  }
  const tpl = _TEMPLATES[_currentTemplate] || _TEMPLATES.recordatorio;
  const currentFields = readCurrentTemplateFields();
  const resumenMensaje = (currentFields.mensaje || tpl.mensaje || '').replace(/\s+/g, ' ').trim();
  if (el3t) el3t.textContent = tpl.nombre;
  if (el3s) el3s.textContent = resumenMensaje
    ? resumenMensaje.slice(0, 80).replace(/{nombre}/g, '[Jefe]') + (resumenMensaje.length > 80 ? '…' : '')
    : '—';
  if (elBtn) elBtn.textContent = count;
  wizardEvalSendBtn();
}

function wizardEvalSendBtn() {
  const chk = $('chkConfirmSend');
  const btn = $('btnAllSend');
  const count = getEffectiveSelectedSupervisores().length;
  const ok = chk && chk.checked && count > 0 && !state.massSendInFlight;
  if (btn) btn.disabled = !ok;
}

function populateWizardFiltersFromRows(rows) {
  const bps = new Set();
  const gers = new Set();
  const areas = new Set();
  (rows || []).forEach(item => {
    if (item.hrbp) bps.add(item.hrbp);
    if (item.gerencia || item.departamento) gers.add(item.gerencia || item.departamento);
    if (item.area) areas.add(item.area);
  });
  if ($('wizFilterBP')) {
    const currentVal = $('wizFilterBP').value;
    $('wizFilterBP').innerHTML = '<option value="">Todos los BP</option>' + Array.from(bps).sort().map(bp => `<option value="${bp}">${bp}</option>`).join('');
    $('wizFilterBP').value = currentVal;
  }
  if ($('wizFilterGerencia')) {
    const currentVal = $('wizFilterGerencia').value;
    $('wizFilterGerencia').innerHTML = '<option value="">Todas las Gerencias</option>' + Array.from(gers).sort().map(g => `<option value="${g}">${g}</option>`).join('');
    $('wizFilterGerencia').value = currentVal;
  }
  if ($('wizFilterArea')) {
    const currentVal = $('wizFilterArea').value;
    $('wizFilterArea').innerHTML = '<option value="">Todas las Áreas</option>' + Array.from(areas).sort().map(a => `<option value="${a}">${a}</option>`).join('');
    $('wizFilterArea').value = currentVal;
  }
}

function renderPersonasObjetivo() {
  const container = $('jefesArbol');
  if (!container) return;
  const rows = (state.objectiveRows || []).filter(r => {
    if (!(r.persona_objetivo || r.objetivo || r.total_dias)) return false;
    // Excluir los que ya cumplieron su meta
    const gozados = Number(r.dias_gozados || 0);
    const objetivo = Number(r.objetivo || 0);
    if (objetivo > 0 && gozados >= objetivo) return false;
    return true;
  });
  populateWizardFiltersFromRows(rows);
  if (!rows.length) {
    container.innerHTML = '<div class="wizard-empty">No se encontraron personas objetivo para mostrar.</div>';
    return;
  }
  const sorted = [...rows].sort((a, b) => (Number(b.por_programar || 0) - Number(a.por_programar || 0)) || String(a.nombre || '').localeCompare(String(b.nombre || ''), 'es'));
  container.innerHTML = sorted.map(row => {
    const key = _personKey(row);
    const keyAttr = key.replace(/"/g, '&quot;');
    const keyDom = key.replace(/[^A-Z0-9_-]/gi, '_');
    const selected = state.selectedPersonas.has(key);
    const supervisor = row.supervisor || 'Sin jefe';
    const hrbp = row.hrbp || 'Sin BP';
    const gerencia = row.gerencia || row.departamento || 'Sin gerencia';
    const area = row.area || row.seccion || 'Sin área';
    const objetivo = Number(row.objetivo || 0).toFixed(0);
    const gozados = Number(row.dias_gozados || 0).toFixed(0);
    const porProgramar = Number(row.por_programar || 0).toFixed(0);
    const saldo = Number(row.total_dias || 0).toFixed(0);
    return `
      <div class="jefe-card persona-card${selected ? ' selected' : ''}"
           data-person-key="${keyAttr}"
           data-jefe-name="${(supervisor || '').replace(/"/g, '&quot;')}"
           data-hrbps="${(hrbp || '').toLowerCase().replace(/"/g, '&quot;')}"
           data-gerencias="${(gerencia || '').toLowerCase().replace(/"/g, '&quot;')}"
           data-areas="${(area || '').toLowerCase().replace(/"/g, '&quot;')}">
        <div class="jefe-card-header" data-person-key="${keyAttr}" onclick="togglePersonaCardByKey(this.dataset.personKey)">
          <input type="checkbox" class="jefe-check" data-person-key="${keyAttr}" ${selected ? 'checked' : ''}
            onclick="event.stopPropagation();togglePersonaCard(this)" />
          <div class="jefe-avatar">${(row.nombre || '?').split(/\s+/).filter(Boolean).slice(0, 2).map(p => p[0]).join('').toUpperCase().slice(0, 2)}</div>
          <div class="jefe-info">
            <div class="jefe-name">${row.nombre || '-'}</div>
            <div class="jefe-meta"><span>${area}</span><span style="color:var(--brand)">Jefe: ${supervisor}</span></div>
          </div>
          <span class="jefe-badge">${porProgramar} por prog.</span>
          <button class="jefe-expand-btn" data-person-dom="${keyDom}" onclick="event.stopPropagation();togglePersonaCardBody(this.dataset.personDom)" title="Ver detalle">▶</button>
        </div>
        <div class="jefe-card-body" id="persona_${keyDom}_body">
          <div class="jefe-origin">
            <span class="jefe-origin-pill">BP: ${hrbp}</span>
            <span class="jefe-origin-pill">Gerencia: ${gerencia}</span>
            <span class="jefe-origin-pill">Saldo: ${saldo} días</span>
          </div>
          <div class="area-group">
            <div class="area-group-title">Estado vacacional</div>
            <div class="area-chip-list">
              <span class="area-chip">Objetivo: <b>${objetivo}</b></span>
              <span class="area-chip">Gozado: <b>${gozados}</b></span>
              <span class="area-chip">Por programar: <b>${porProgramar}</b></span>
              <span class="area-chip">Estado: <b>${row.estado || '-'}</b></span>
            </div>
          </div>
        </div>
      </div>`;
  }).join('');
  wizardFilterJefes();
}

function renderWizardModeUi() {
  const mode = getWizardRecipientMode();
  const titles = { personas: '👥 PERSONAS QUE FALTAN TOMAR VACACIONES', sin_meta: '📋 PERSONAS SIN META ASIGNADA', jefes: '👥 JEFATURAS GENERALES' };
  const helpers = {
    personas: 'Solo se muestran personas que aún no han cumplido su meta. El envío sale al jefe directo resuelto desde el Excel objetivo.',
    sin_meta: 'Personas sin meta de vacaciones asignada en el trimestre. El envío sale al jefe directo de cada una.',
    jefes: 'Cada jefatura funciona como un toggle: un clic la activa, otro la desactiva. Usa el chevron para ver sus áreas.'
  };
  const placeholders = { personas: 'Buscar persona...', sin_meta: 'Buscar persona sin meta...', jefes: 'Buscar jefatura...' };
  if ($('wizardRecipientTitle')) $('wizardRecipientTitle').textContent = titles[mode] || titles.jefes;
  if ($('wizardRecipientHelper')) $('wizardRecipientHelper').textContent = helpers[mode] || helpers.jefes;
  if ($('wizSearchJefe')) $('wizSearchJefe').placeholder = placeholders[mode] || placeholders.jefes;
}

function renderPersonasSinMeta() {
  const container = $('jefesArbol');
  if (!container) return;
  // Muestra personas cuya meta es 0 o nula pero sí tienen saldo de vacaciones
  const rows = (state.objectiveRows || []).filter(r => {
    const meta = Number(r.objetivo || 0);
    const saldo = Number(r.total_dias || 0);
    return meta === 0 && saldo > 0;
  });
  populateWizardFiltersFromRows(rows);
  if (!rows.length) {
    container.innerHTML = '<div class="wizard-empty">No hay personas sin meta con saldo de vacaciones.</div>';
    return;
  }
  const sorted = [...rows].sort((a, b) => Number(b.total_dias || 0) - Number(a.total_dias || 0) || String(a.nombre || '').localeCompare(String(b.nombre || ''), 'es'));
  container.innerHTML = sorted.map(row => {
    const key = _personKey(row);
    const keyAttr = key.replace(/"/g, '&quot;');
    const keyDom = key.replace(/[^A-Z0-9_-]/gi, '_');
    const selected = state.selectedPersonas.has(key);
    const supervisor = row.supervisor || 'Sin jefe';
    const hrbp = row.hrbp || 'Sin BP';
    const gerencia = row.gerencia || row.departamento || 'Sin gerencia';
    const area = row.area || row.seccion || 'Sin área';
    const saldo = Number(row.total_dias || 0).toFixed(0);
    return `
      <div class="jefe-card persona-card${selected ? ' selected' : ''}"
           data-person-key="${keyAttr}"
           data-jefe-name="${(supervisor || '').replace(/"/g, '&quot;')}"
           data-hrbps="${(hrbp || '').toLowerCase().replace(/"/g, '&quot;')}"
           data-gerencias="${(gerencia || '').toLowerCase().replace(/"/g, '&quot;')}"
           data-areas="${(area || '').toLowerCase().replace(/"/g, '&quot;')}">
        <div class="jefe-card-header" data-person-key="${keyAttr}" onclick="togglePersonaCardByKey(this.dataset.personKey)">
          <input type="checkbox" class="jefe-check" data-person-key="${keyAttr}" ${selected ? 'checked' : ''}
            onclick="event.stopPropagation();togglePersonaCard(this)" />
          <div class="jefe-avatar">${(row.nombre || '?').split(/\s+/).filter(Boolean).slice(0,2).map(p=>p[0]).join('').toUpperCase().slice(0,2)}</div>
          <div class="jefe-info">
            <div class="jefe-name">${row.nombre || '-'}</div>
            <div class="jefe-meta"><span>${area}</span><span style="color:var(--brand)">Jefe: ${supervisor}</span></div>
          </div>
          <span class="jefe-badge" style="background:#fef3c7;color:#92400e">${saldo} días saldo</span>
          <button class="jefe-expand-btn" data-person-dom="${keyDom}" onclick="event.stopPropagation();togglePersonaCardBody(this.dataset.personDom)" title="Ver detalle">▶</button>
        </div>
        <div class="jefe-card-body" id="persona_${keyDom}_body">
          <div class="jefe-origin">
            <span class="jefe-origin-pill">BP: ${hrbp}</span>
            <span class="jefe-origin-pill">Gerencia: ${gerencia}</span>
            <span class="jefe-origin-pill" style="background:#fef3c7;color:#92400e">Sin meta asignada</span>
          </div>
        </div>
      </div>`;
  }).join('');
  wizardFilterJefes();
}

function renderWizardRecipients() {
  renderWizardModeUi();
  const mode = getWizardRecipientMode();
  if (mode === 'personas') {
    renderPersonasObjetivo();
  } else if (mode === 'sin_meta') {
    renderPersonasSinMeta();
  } else {
    renderJefesArbol(state.supervisoresData || []);
  }
  updateSummaryCards();
  wizardRefreshResumen();
}

function wizardChangeRecipientMode() {
  const modeNode = $('wizRecipientMode');
  state.wizardRecipientMode = ((modeNode && modeNode.value) || 'personas').toLowerCase();
  if (!state.objectiveRows || !state.objectiveRows.length) {
    const container = $('jefesArbol');
    if (container) container.innerHTML = '<span class="spin"></span> Cargando base...';
    loadObjectiveRows().catch(e => log(e.message)).then(() => renderWizardRecipients());
    return;
  }
  renderWizardRecipients();
}

// --- Botón de envío del wizard (puente al setMassArmed + enviarGlobal existentes) ---
function wizardLanzar() {
  if (state.massSendInFlight) return;
  const chk = document.getElementById('chkConfirmSend');
  if (!chk || !chk.checked) return;
  
  const mode = getWizardRecipientMode();
  const selectedSupervisores = getEffectiveSelectedSupervisores();
  const selectedPersonRows = getEffectiveSelectedPersonRows();
  const isSingleJefeSend = mode === 'jefes' && selectedSupervisores.length === 1;
  const isSinglePersonaSend = mode === 'personas' && selectedPersonRows.length === 1;
  
  const btn = document.getElementById('btnAllSend');
  let originalBtnHtml = '';
  if (btn) {
    originalBtnHtml = btn.innerHTML;
    btn.innerHTML = '<span class="spin" style="width:18px;height:18px;border-width:2px;margin-right:8px;border-top-color:#fff;border-right-color:rgba(255,255,255,.3);border-bottom-color:rgba(255,255,255,.3);border-left-color:rgba(255,255,255,.3);"></span> Enviando...';
    btn.disabled = true;
    btn.style.opacity = '0.9';
  }

  // Armar internamente (evita el confirm nativo)
  setMassArmed(true);
  const runSend = isSinglePersonaSend
    ? enviarPersonaObjetivoSeleccionada(selectedPersonRows[0], getSelectedMassChannels())
    : (isSingleJefeSend
      ? enviarSupervisorSeleccionado(selectedSupervisores[0], getSelectedMassChannels())
      : enviarGlobal(getSelectedMassChannels()));
      
  // Llamar al envío existente
  runSend
    .then(() => {
      // Resetear guardia
      if (btn) {
        btn.innerHTML = originalBtnHtml;
        btn.style.opacity = '';
      }
      resetWizardAfterSend();
      if (chk) chk.checked = false;
      wizardEvalSendBtn();
      const pill = document.getElementById('massStatusPill');
      if (pill) { pill.className = 'status-pill queue'; pill.textContent = '📬 Mensajes en camino'; }
    })
    .catch(e => {
      if (btn) {
        btn.innerHTML = originalBtnHtml;
        btn.style.opacity = '';
      }
      setStatus(e.message, true);
      notify(e.message || 'Error al enviar campaña.', 'err');
      setMassArmed(false);
      if (chk) chk.checked = false;
      wizardEvalSendBtn();
    });
}

async function enviarSupervisorSeleccionado(nombreSupervisor, flags) {
  const key = normalizeSupervisorName(nombreSupervisor);
  const supervisor = state.supMap[(nombreSupervisor || '').toLowerCase()] || state.supMap[(key || '').toLowerCase()];
  if (!supervisor) {
    throw new Error('No se pudo resolver la jefatura seleccionada para envío individual.');
  }
  const payload = {
    nombre: supervisor.nombre || nombreSupervisor,
    email: supervisor.email || '',
    hrbp_objetivo: getActiveWizardHrbpFilter(),
    asunto: $('txtAsunto') ? $('txtAsunto').value : '',
    mensaje: $('txtMensaje') ? $('txtMensaje').value : '',
    aviso: $('txtAviso') ? $('txtAviso').value : '',
    recomendacion: $('txtReco') ? $('txtReco').value : '',
    modo_prueba: $('chkModoPrueba') ? $('chkModoPrueba').checked : false,
    enviar_teams: !!flags.enviar_teams,
    enviar_smtp: !!flags.enviar_smtp,
    encolar_para_pa: !!flags.encolar_pa,
  };
  setStatus(`Enviando mensaje a ${payload.nombre}...`);
  notify(`Enviando mensaje a ${payload.nombre}...`, 'warn', 2600);
  const data = await jfetch('/api/enviar-a-supervisor', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
  await Promise.allSettled([refreshQueueState(), refreshConfirmacionesResumen()]);
  const delay = Number(data.retraso_seguridad_segundos || (((data || {}).json_meta || {}).retraso_segundos || 60));
  const infoFile = data.archivo || 'sin archivo';
  const paOk = payload.encolar_para_pa && data.json_guardado;
  const teamsOk = payload.enviar_teams && data.teams_enviado;
  const smtpOk = payload.enviar_smtp && data.smtp_enviado;
  if (paOk) {
    setStatus(`✅ Correo en camino para ${payload.nombre}. Power Automate lo procesará en ~${delay}s.${teamsOk ? ' Teams: OK.' : ''}`);
  } else if (teamsOk || smtpOk) {
    setStatus(`✅ Envío ejecutado para ${payload.nombre}.`);
  } else {
    const teamsErr = payload.enviar_teams && !data.teams_enviado ? ` Teams: ${data.teams_error || 'sin webhook'}` : '';
    setStatus(`⚠️ Envío para ${payload.nombre} con advertencias.${teamsErr}`);
  }
  notify(`Mensaje enviado para ${payload.nombre}. Se procesa en ${delay}s (${infoFile}).`, 'ok');
  return data;
}

async function enviarPersonaObjetivoSeleccionada(row, flags) {
  const supervisorNombre = String((row && row.supervisor) || '').trim();
  if (!supervisorNombre) {
    throw new Error('La persona seleccionada no tiene una jefatura asociada para el envío.');
  }
  const supervisor = state.supMap[supervisorNombre.toLowerCase()] || {};
  const nombrePersona = String((row && row.nombre) || supervisorNombre).trim();
  const payload = {
    nombre: supervisor.nombre || supervisorNombre,
    email: supervisor.email || '',
    nombre_objetivo: nombrePersona,
    email_objetivo: String((row && (row.correo || row.email)) || '').trim(),
    hrbp_objetivo: String((row && row.hrbp) || '').trim(),
    asunto: $('txtAsunto') ? $('txtAsunto').value : '',
    mensaje: $('txtMensaje') ? $('txtMensaje').value : '',
    aviso: $('txtAviso') ? $('txtAviso').value : '',
    recomendacion: $('txtReco') ? $('txtReco').value : '',
    modo_prueba: $('chkModoPrueba') ? $('chkModoPrueba').checked : false,
    enviar_teams: !!flags.enviar_teams,
    enviar_smtp: !!flags.enviar_smtp,
    encolar_para_pa: !!flags.encolar_pa,
  };
  setStatus(`Enviando mensaje para ${nombrePersona}...`);
  notify(`Enviando mensaje para ${nombrePersona}...`, 'warn', 2600);
  const data = await jfetch('/api/enviar-a-supervisor', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
  await Promise.allSettled([refreshQueueState(), refreshConfirmacionesResumen()]);
  const delay = Number(data.retraso_seguridad_segundos || (((data || {}).json_meta || {}).retraso_segundos || 60));
  const infoFile = data.archivo || 'sin archivo';
  const paOk2 = payload.encolar_para_pa && data.json_guardado;
  const teamsOk2 = payload.enviar_teams && data.teams_enviado;
  const smtpOk2 = payload.enviar_smtp && data.smtp_enviado;
  if (paOk2) {
    setStatus(`✅ Correo en camino para ${nombrePersona}. Power Automate lo procesará en ~${delay}s.${teamsOk2 ? ' Teams: OK.' : ''}`);
  } else if (teamsOk2 || smtpOk2) {
    setStatus(`✅ Envío ejecutado para ${nombrePersona}.`);
  } else {
    const teamsErr2 = payload.enviar_teams && !data.teams_enviado ? ` Teams: ${data.teams_error || 'sin webhook'}` : '';
    setStatus(`⚠️ Envío para ${nombrePersona} con advertencias.${teamsErr2}`);
  }
  notify(`Mensaje enviado para ${nombrePersona}. Se procesa en ${delay}s (${infoFile}).`, 'ok');
  return data;
}

function resetWizardAfterSend() {
  state.selectedSupervisores.clear();
  state.selectedPersonas.clear();
  setMassArmed(false);
  renderWizardRecipients();
  wizardGo(1);
}

// --- Filtro de búsqueda de jefes en paso 1 ---
function wizardFilterJefes() {
  const q = normTxt($('wizSearchJefe') ? $('wizSearchJefe').value : '');
  const bp = ($('wizFilterBP') ? $('wizFilterBP').value : '').toLowerCase();
  const ger = ($('wizFilterGerencia') ? $('wizFilterGerencia').value : '').toLowerCase();
  const area = ($('wizFilterArea') ? $('wizFilterArea').value : '').toLowerCase();

  document.querySelectorAll('.jefe-card, .persona-card').forEach(block => {
    const name = (block.dataset.jefeName || '').toLowerCase();
    const cardBp = (block.dataset.hrbps || '').toLowerCase().split('|').filter(Boolean);
    const cardGers = (block.dataset.gerencias || '').toLowerCase();
    const cardAreas = (block.dataset.areas || '').toLowerCase();

    const matchText = !q || name.includes(q) || matchByTokens(name, q);
    const matchBP = !bp || cardBp.includes(bp);
    const matchGer = !ger || cardGers.includes(ger);
    const matchArea = !area || cardAreas.includes(area);

    block.style.display = (matchText && matchBP && matchGer && matchArea) ? '' : 'none';
  });
  updateSummaryCards();
  wizardRefreshResumen();
}

// --- Seleccionar / deseleccionar todos los jefes ---
function wizardSelAll(mark) {
  if (getWizardRecipientMode() === 'personas') {
    getVisiblePersonCards().forEach(card => {
      const chk = card.querySelector('.jefe-check[data-person-key]');
      if (!chk) return;
      const key = (chk.getAttribute('data-person-key') || '').trim().toUpperCase();
      chk.checked = !!mark;
      if (mark) state.selectedPersonas.add(key);
      else state.selectedPersonas.delete(key);
      card.classList.toggle('selected', !!mark);
    });
    updateSummaryCards();
    wizardRefreshResumen();
    return;
  }
  getVisibleJefeCards().forEach(card => {
    const chk = card.querySelector('.jefe-check[data-jefe]');
    if (!chk) return;
    const jefeNom = chk.getAttribute('data-jefe');
    chk.checked = !!mark;
    const k = normalizeSupervisorName(jefeNom);
    if (mark) state.selectedSupervisores.add(k);
    else state.selectedSupervisores.delete(k);
    if (card) card.classList.toggle('selected', !!mark);
  });
  updateSummaryCards();
  wizardRefreshResumen();
}

// --- Árbol de jefes rediseñado ---
// Sobrescribe renderJefesArbol para usar las nuevas tarjetas visuales
const _origRenderJefesArbol = renderJefesArbol;
renderJefesArbol = function (arbol) {
  const container = document.getElementById('jefesArbol');
  if (!container) return;
  if (!arbol || !arbol.length) {
    container.innerHTML = '<div class="wizard-empty">No se encontraron jefaturas para mostrar.</div>';
    return;
  }
  let html = '';
  for (const jefe of arbol) {
    const jefe_id = 'jefe_' + (jefe.nombre || '').replace(/[^a-z0-9]/gi, '_').toLowerCase();
    const checkbox_id = jefe_id + '_check';
    const isSelected = state.selectedSupervisores.has(normalizeSupervisorName(jefe.nombre));
    const initials = (jefe.nombre || '?').split(/[,\s]+/).filter(Boolean).slice(-2).map(w => w[0]).join('').toUpperCase().slice(0, 2);
    const totalColab = Number(jefe.total_colaboradores || 0);
        const hrbps = Array.isArray(jefe.hrbps) ? jefe.hrbps.filter(Boolean) : (jefe.hrbp ? [jefe.hrbp] : []);
        const hrbp = jefe.hrbp || hrbps[0] || '';
    const esVistaEjecutiva = !!jefe.es_vista_ejecutiva;
    const origenEmail = jefe.origen_email || 'Sin origen identificado';
    const gers = ((jefe.gerencias || []).concat(jefe.subgerencias || [])).join('|').toLowerCase();
    const areas = (jefe.areas || []).map(a => a.nombre || a).join('|').toLowerCase();
        const hrbpsAttr = hrbps.map(v => v.toLowerCase()).join('|');
    const hrbpLabel = esVistaEjecutiva ? '' : (hrbps.length > 1 ? hrbps.join(' | ') : hrbp);
    const principalPill = esVistaEjecutiva ? 'Vista ejecutiva' : `BP: ${hrbpLabel || 'Sin BP'}`;

    html += `
      <div class="jefe-card${isSelected ? ' selected' : ''}" 
           data-jefe-name="${(jefe.nombre || '').replace(/"/g, '&quot;')}" 
          data-hrbp="${hrbp.replace(/"/g, '&quot;')}"
          data-hrbps="${hrbpsAttr.replace(/"/g, '&quot;')}"
           data-gerencias="${gers.replace(/"/g, '&quot;')}"
           data-areas="${areas.replace(/"/g, '&quot;')}">
        <div class="jefe-card-header" onclick="toggleJefeCardById('${checkbox_id}')">
          <input type="checkbox" id="${checkbox_id}" class="jefe-check" data-jefe="${(jefe.nombre || '').replace(/"/g, '&quot;')}"
            ${isSelected ? 'checked' : ''}
            onclick="event.stopPropagation();toggleJefeCard(this)"
          />
          <div class="jefe-avatar">${initials}</div>
          <div class="jefe-info">
            <div class="jefe-name">${jefe.nombre || '-'}</div>
            <div class="jefe-meta"><span>${(jefe.areas || []).length} área(s)</span>${hrbpLabel ? `<span style="color:var(--brand)">${hrbpLabel}</span>` : ''}</div>
          </div>
          <span class="jefe-badge">${totalColab} colab.</span>
          <button class="jefe-expand-btn" onclick="event.stopPropagation();toggleJefeCardBody('${jefe_id}')" title="Ver equipo">▶</button>
        </div>
        <div class="jefe-card-body" id="${jefe_id}_body">
          <div class="jefe-origin">
            <span class="jefe-origin-pill">${principalPill}</span>
            <span class="jefe-origin-pill">Origen: ${origenEmail}</span>
            <span class="jefe-origin-pill">Áreas visibles: ${(jefe.areas || []).length}</span>
          </div>
          ${(() => {
            const jefeNameNorm = normalizeSupervisorName(jefe.nombre);
            const allJefePersonas = (state.objectiveRows || []).filter(r => normalizeSupervisorName(r.supervisor) === jefeNameNorm);
            
            if (!allJefePersonas.length) {
              return (jefe.areas || []).map(area => {
                const areaNameStr = area.nombre || area || '';
                const areaCount = Number(area.total || 0);
                return `<div class="area-group"><div class="area-group-title">${areaNameStr || '-'}</div><div class="area-chip-list" style="margin-top:6px"><span class="area-chip"><b>${areaCount}</b> personas objetivo</span></div></div>`;
              }).join('');
            }
            
            const grouped = {};
            allJefePersonas.forEach(p => {
              const aName = p.area || p.seccion || p.departamento || 'Sin Área';
              if (!grouped[aName]) grouped[aName] = [];
              grouped[aName].push(p);
            });
            
            return Object.keys(grouped).sort().map(aName => {
              const chips = grouped[aName].map(p => {
                const pct = Number(p.objetivo || 0) > 0 ? (Number(p.dias_gozados || 0) / Number(p.objetivo)) * 100 : 0;
                const dotCls = pct >= 100 ? 'chip-ok' : pct >= 50 ? 'chip-warn' : 'chip-bad';
                return `<span class="persona-chip" title="Objetivo: ${p.objetivo} | Gozado: ${p.dias_gozados}"><span class="chip-dot ${dotCls}"></span>${p.nombre || '-'} <span style="color:var(--muted)">(${pct.toFixed(0)}%)</span></span>`;
              }).join('');
              return `<div class="area-group"><div class="area-group-title">${aName}</div><div class="area-chip-list" style="margin-top:6px">${chips}</div></div>`;
            }).join('');
          })()}
        </div>
      </div>`;
  }
  container.innerHTML = html;
  wizardFilterJefes();
};

function toggleJefeCardById(checkboxId) {
  const checkbox = document.getElementById(checkboxId);
  if (!checkbox) return;
  checkbox.checked = !checkbox.checked;
  toggleJefeCard(checkbox);
}

function toggleJefeCard(checkbox) {
  const jefe_nombre = checkbox.getAttribute('data-jefe');
  const card = checkbox.closest('.jefe-card');
  if (checkbox.checked) {
    state.selectedSupervisores.add(normalizeSupervisorName(jefe_nombre));
    if (card) card.classList.add('selected');
  } else {
    state.selectedSupervisores.delete(normalizeSupervisorName(jefe_nombre));
    if (card) card.classList.remove('selected');
  }
  updateSummaryCards();
  wizardRefreshResumen();
}

function togglePersonaCardByKey(personKey) {
  const checkbox = document.querySelector(`.jefe-check[data-person-key="${CSS.escape((personKey || '').toUpperCase())}"]`);
  if (!checkbox) return;
  checkbox.checked = !checkbox.checked;
  togglePersonaCard(checkbox);
}

function togglePersonaCard(checkbox) {
  const key = (checkbox.getAttribute('data-person-key') || '').trim().toUpperCase();
  const card = checkbox.closest('.persona-card');
  if (checkbox.checked) {
    state.selectedPersonas.add(key);
    if (card) card.classList.add('selected');
  } else {
    state.selectedPersonas.delete(key);
    if (card) card.classList.remove('selected');
  }
  updateSummaryCards();
  wizardRefreshResumen();
}

function togglePersonaCardBody(personKey) {
  const body = document.getElementById(`persona_${personKey}_body`);
  const header = body && body.previousElementSibling;
  const btn = header && header.querySelector('.jefe-expand-btn');
  if (!body) return;
  const open = body.classList.toggle('open');
  if (btn) btn.textContent = open ? '▼' : '▶';
}

function toggleJefeCardBody(jefe_id) {
  const body = document.getElementById(jefe_id + '_body');
  const header = body && body.previousElementSibling;
  const btn = header && header.querySelector('.jefe-expand-btn');
  if (!body) return;
  const open = body.classList.toggle('open');
  if (btn) btn.textContent = open ? '▼' : '▶';
}

// --- Toggle visual de botones de canal ---
function toggleChannelBtn(labelEl, checkboxId) {
  const chk = document.getElementById(checkboxId);
  if (!chk) return;
  chk.checked = !chk.checked;
  labelEl.classList.toggle('active', chk.checked);
}

// Inicializar plantilla por defecto al arrancar
window.addEventListener('load', () => {
  selectTemplate(document.querySelector('.template-card.active, .tpl.active') || document.querySelector('.template-card, .tpl'), 'recordatorio');
});

window.wizGo = wizardGo;
window.selTpl = selectTemplate;
window.updPrev = updateMsgPreview;
window.evalSendBtn = wizardEvalSendBtn;
window.wizSelAll = wizardSelAll;
window.wizFilter = wizardFilterJefes;
window.wizLanzar = wizardLanzar;
window.wizardChangeRecipientMode = wizardChangeRecipientMode;
window.togglePersonaCardByKey = togglePersonaCardByKey;
window.togglePersonaCard = togglePersonaCard;
window.togglePersonaCardBody = togglePersonaCardBody;
