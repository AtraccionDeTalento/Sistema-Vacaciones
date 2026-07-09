/* ============================================================================
 * pipeline_vac.js  ·  Actualizar Sistema de Vacaciones + KPIs en vivo
 * Autocontenido: no depende de app_completo.js. Dispara el pipeline en el
 * servidor, hace polling del progreso y anima el avance de meta antes->despues.
 * ==========================================================================*/
(function () {
  'use strict';
  function $(id) { return document.getElementById(id); }
  var pollTimer = null, jobId = null, antes = null, _ultimaOpAdryan = null;

  function fmtPct(v) {
    if (v === null || v === undefined || isNaN(v)) return '—';
    return (v * 100).toFixed(1) + '%';
  }
  function fmtNum(v) {
    if (v === null || v === undefined || isNaN(v)) return '—';
    return Math.round(v).toLocaleString('es-PE');
  }
  function pickAvance(kp) {
    if (!kp) return null;
    var v = (kp.avance_cumplimiento != null) ? kp.avance_cumplimiento : kp.avance;
    return (typeof v === 'number') ? v : null;
  }

  function setBar(pct, err) {
    pct = Math.max(0, Math.min(100, pct || 0));
    var f = $('pipeBar');
    if (f) { f.style.width = pct + '%'; f.classList.toggle('err', !!err); }
    if ($('pipePct')) $('pipePct').textContent = Math.round(pct) + '%';
  }
  function setStep(txt) { if ($('pipeStep')) $('pipeStep').textContent = txt || ''; }

  function setAvance(val) {
    if ($('pipeAvance')) $('pipeAvance').textContent = fmtPct(val);
    if ($('pipeAvanceBar')) $('pipeAvanceBar').style.width = (val ? Math.min(100, val * 100) : 0) + '%';
  }
  function setAvanceDelta(dp) {
    var d = $('pipeAvanceDelta'); if (!d) return;
    if (dp == null || isNaN(dp) || Math.abs(dp) < 1e-9) { d.textContent = ''; d.classList.remove('neg'); return; }
    var pp = dp * 100;
    d.textContent = (pp >= 0 ? '▲ +' : '▼ ') + pp.toFixed(1) + ' pp';
    d.classList.toggle('neg', pp < 0);
  }
  function setReg(val, antesVal, metaVal) {
    if ($('pipeReg')) $('pipeReg').textContent = fmtNum(val);
    if (metaVal != null && $('pipeMeta')) $('pipeMeta').textContent = fmtNum(metaVal);
    var d = $('pipeRegDelta'); if (!d) return;
    if (antesVal != null && val != null && Math.round(val - antesVal) !== 0) {
      var diff = Math.round(val - antesVal);
      d.textContent = (diff >= 0 ? '+' : '') + diff.toLocaleString('es-PE') + ' días';
      d.classList.toggle('neg', diff < 0);
    } else { d.textContent = ''; }
  }
  function appendLog(lineas) {
    if (!lineas || !$('pipeLog')) return;
    var el = $('pipeLog');
    el.textContent = lineas.join('\n');
    el.scrollTop = el.scrollHeight;
  }
  function animateAvance(from, to) {
    if (to == null) { setAvance(from); return; }
    from = (from == null ? to : from);
    var t0 = performance.now(), dur = 1100;
    function frame(now) {
      var k = Math.min(1, (now - t0) / dur);
      var v = from + (to - from) * (k * (2 - k)); // easeOutQuad
      setAvance(v);
      if (k < 1) requestAnimationFrame(frame);
    }
    requestAnimationFrame(frame);
  }

  function openModal() {
    if ($('pipeModal')) $('pipeModal').classList.remove('hidden');
    if ($('pipeReload')) $('pipeReload').classList.add('hidden');
    if ($('pipeResetBtn')) $('pipeResetBtn').classList.add('hidden');
    setBar(0, false); setStep('Listo para actualizar.');
    if ($('pipeLog')) $('pipeLog').textContent = '';
    setBusy(false);
    cargarAdryan();
    fetchKpisInicial();
  }
  function closeModal() {
    if ($('pipeModal')) $('pipeModal').classList.add('hidden');
    if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
  }

  function fetchKpisInicial() {
    setAvance(null); setReg(null, null, null);
    fetch('/api/vacaciones/kpis').then(function (r) { return r.json(); }).then(function (j) {
      if (j && j.ok && j.kpis) {
        antes = j.kpis;
        setAvance(pickAvance(j.kpis));
        setReg(j.kpis.registrado_total, null, j.kpis.meta_total);
      }
    }).catch(function () {});
  }

  function tieneArchivo() {
    var f = $('pipeFile');
    return !!(f && f.files && f.files[0]);
  }
  function setBusy(b) {
    if ($('pipeRun')) $('pipeRun').disabled = b;
    if ($('pipeRunAdryan')) $('pipeRunAdryan').disabled = b;
    if ($('pipeRunUpload')) $('pipeRunUpload').disabled = b || !tieneArchivo();
    if ($('pipeRunMaestro')) $('pipeRunMaestro').disabled = b;
  }
  function startRun(body) {
    setBusy(true);
    if ($('pipeReload')) $('pipeReload').classList.add('hidden');
    setBar(3, false); setStep('Iniciando…');
    if ($('pipeLog')) $('pipeLog').textContent = '';
    jobId = null;
    var opts = { method: 'POST' };
    if (body) opts.body = body;            // FormData (metodo 2)
    fetch('/api/vacaciones/pipeline/run', opts)
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j.ok) { fail(j.error || 'No se pudo iniciar'); setBusy(false); return; }
        jobId = j.job_id;
        pollTimer = setInterval(poll, 1000);
      })
      .catch(function (e) { fail('' + e); setBusy(false); });
  }
  function run() { startRun(null); }       // metodo 1: ultimo de Descargas
  function runAdryan() {                    // metodo 0: el bot descarga de Adryan y luego procesa
    _ultimaOpAdryan = 'adryan';
    var fd = new FormData();
    fd.append('descargar', '1');
    var fi = $('pipeFechaInicio'), ft = $('pipeFechaTermino');
    if (fi && fi.value) fd.append('fecha_inicio', fi.value);
    if (ft && ft.value) fd.append('fecha_termino', ft.value);
    startRun(fd);
  }
  function runUpload() {                    // metodo 2: archivo subido
    if (!tieneArchivo()) return;
    var fd = new FormData();
    fd.append('archivo', $('pipeFile').files[0]);
    startRun(fd);
  }
  function runMaestro() {
    _ultimaOpAdryan = 'maestro';
    setBusy(true);
    if ($('pipeReload')) $('pipeReload').classList.add('hidden');
    setBar(3, false); setStep('Iniciando descarga de maestro…');
    if ($('pipeLog')) $('pipeLog').textContent = '';
    jobId = null;
    fetch('/api/vacaciones/pipeline/maestro', { method: 'POST' })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j.ok) { fail(j.error || 'No se pudo iniciar el maestro'); setBusy(false); return; }
        jobId = j.job_id;
        pollTimer = setInterval(poll, 1000);
      })
      .catch(function (e) { fail('' + e); setBusy(false); });
  }
  function cargarAdryan() {
    fetch('/api/vacaciones/config').then(function (r) { return r.json(); }).then(function (c) {
      var a = $('pipeAdryan'); if (!a) return;
      if (c && c.url_adryan) { a.href = c.url_adryan; a.removeAttribute('data-empty'); }
      else { a.href = '#'; a.setAttribute('data-empty', '1'); }
      if (c && c.fecha_inicio) {
        var fi = $('pipeFechaInicio'); if (fi && !fi.value) fi.value = c.fecha_inicio;
      }
      if (c && c.fecha_termino) {
        var ft = $('pipeFechaTermino'); if (ft && !ft.value) ft.value = c.fecha_termino;
      }
    }).catch(function () {});
  }

  function poll() {
    if (!jobId) return;
    fetch('/api/vacaciones/pipeline/estado/' + jobId)
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j.ok) return;
        if (j.kpis_antes) antes = j.kpis_antes;
        setBar(j.pct, j.estado === 'error');
        setStep(j.paso || '');
        appendLog(j.lineas);
        if (antes && j.estado === 'running') {
          setAvance(pickAvance(antes));
          setReg(antes.registrado_total, null, antes.meta_total);
        }
        if (j.estado === 'done') { clearInterval(pollTimer); pollTimer = null; done(j); }
        else if (j.estado === 'error') {
          clearInterval(pollTimer); pollTimer = null;
          if (j.necesita_password_adryan) {
            setBusy(false);
            mostrarModalPasswordAdryan();
          } else {
            fail((j.lineas && j.lineas.slice(-3).join(' · ')) || j.error || 'Error en el proceso');
          }
        }
      })
      .catch(function () {});
  }

  function done(j) {
    if (!j.kpis_despues) {
      // fallback: el worker pudo no alcanzar a leer KPIs; los pedimos al archivo publicado
      fetch('/api/vacaciones/kpis')
        .then(function (r) { return r.json(); })
        .then(function (k) { if (k && k.ok && k.kpis) j.kpis_despues = k.kpis; })
        .catch(function () {})
        .then(function () { finalizeDone(j); });
      return;
    }
    finalizeDone(j);
  }

  function finalizeDone(j) {
    setBar(100, false);
    setStep('✅ Actualizado en ' + (j.duracion || '?') + 's');
    var a = pickAvance(antes), b = pickAvance(j.kpis_despues);
    animateAvance(a, b);
    setAvanceDelta((a != null && b != null) ? (b - a) : null);
    if (j.kpis_despues) {
      setReg(j.kpis_despues.registrado_total, antes && antes.registrado_total, j.kpis_despues.meta_total);
    }
    setBusy(false);
    if ($('pipeReload')) $('pipeReload').classList.remove('hidden');
  }

  function fail(msg) {
    setBar(100, true);
    setStep('❌ ' + msg);
    setBusy(false);
    // Si el error es por lock pegado, mostrar botón de desbloqueo
    if (msg && msg.indexOf('actualizacion en curso') !== -1) {
      var rs = $('pipeResetBtn');
      if (rs) rs.classList.remove('hidden');
    }
  }

  function forceReset() {
    var rs = $('pipeResetBtn'); if (rs) rs.disabled = true;
    fetch('/api/vacaciones/pipeline/reset', { method: 'POST' })
      .then(function(r) { return r.json(); })
      .then(function(j) {
        if (j.ok) {
          setStep('🔓 Desbloqueado. Ya puedes volver a actualizar.');
          setBar(0, false);
          if (rs) rs.classList.add('hidden');
          setBusy(false);
        } else {
          setStep('❌ No se pudo desbloquear: ' + (j.error || 'error'));
          if (rs) rs.disabled = false;
        }
      })
      .catch(function(e) { setStep('❌ ' + e); if (rs) rs.disabled = false; });
  }

  // ── Modal contraseña Adryan ─────────────────────────────────────────────────
  function mostrarModalPasswordAdryan() {
    var existente = document.getElementById('_modalPwAdryan');
    if (existente) existente.remove();
    var modal = document.createElement('div');
    modal.id = '_modalPwAdryan';
    modal.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:99999;display:flex;align-items:center;justify-content:center';
    modal.innerHTML = [
      '<div style="background:#fff;border-radius:14px;padding:28px 32px;max-width:400px;width:90%;box-shadow:0 8px 40px rgba(0,0,0,.18)">',
      '  <div style="font-size:28px;text-align:center;margin-bottom:8px">🔑</div>',
      '  <h3 style="margin:0 0 6px;font-size:17px;color:#1e293b;text-align:center">Contraseña de Adryan requerida</h3>',
      '  <p style="margin:0 0 16px;font-size:13px;color:#64748b;text-align:center">La contraseña no está guardada en este equipo.<br>Ingrésala una vez y se guardará de forma segura.</p>',
      '  <input id="_pwAdryanInput" type="password" placeholder="Contraseña de Adryan" autocomplete="current-password"',
      '    style="width:100%;padding:10px 12px;border:1.5px solid #cbd5e1;border-radius:8px;font-size:14px;box-sizing:border-box;margin-bottom:14px">',
      '  <div id="_pwAdryanError" style="color:#dc2626;font-size:12px;margin-bottom:10px;display:none"></div>',
      '  <div style="display:flex;gap:10px;justify-content:flex-end">',
      '    <button id="_pwAdryanCancelar" style="padding:8px 16px;border:1.5px solid #cbd5e1;background:#f8fafc;border-radius:8px;cursor:pointer;font-size:13px">Cancelar</button>',
      '    <button id="_pwAdryanGuardar" style="padding:8px 18px;background:#2563eb;color:#fff;border:none;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600">Guardar y reintentar</button>',
      '  </div>',
      '</div>'
    ].join('');
    document.body.appendChild(modal);

    var inp = document.getElementById('_pwAdryanInput');
    var errDiv = document.getElementById('_pwAdryanError');
    inp.focus();

    document.getElementById('_pwAdryanCancelar').onclick = function() { modal.remove(); setStep('❌ Operación cancelada'); };
    document.getElementById('_pwAdryanGuardar').onclick = guardarYReintentar;
    inp.addEventListener('keydown', function(e) { if (e.key === 'Enter') guardarYReintentar(); });

    function guardarYReintentar() {
      var pw = inp.value.trim();
      if (!pw) { errDiv.textContent = 'Ingresa la contraseña'; errDiv.style.display = ''; return; }
      var btn = document.getElementById('_pwAdryanGuardar');
      btn.disabled = true; btn.textContent = 'Guardando...';
      fetch('/api/adryan/guardar-password', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ password: pw })
      })
        .then(function(r) { return r.json(); })
        .then(function(res) {
          if (!res.ok) {
            errDiv.textContent = 'Error: ' + (res.error || 'no se pudo guardar');
            errDiv.style.display = '';
            btn.disabled = false; btn.textContent = 'Guardar y reintentar';
            return;
          }
          modal.remove();
          // Reintentar la operación que falló
          if (_ultimaOpAdryan === 'maestro') runMaestro();
          else runAdryan();
        })
        .catch(function(e) {
          errDiv.textContent = 'Error de red: ' + e;
          errDiv.style.display = '';
          btn.disabled = false; btn.textContent = 'Guardar y reintentar';
        });
    }
  }

  function wire() {
    var b = $('btnActualizarVac'); if (b) b.addEventListener('click', openModal);
    var c = $('pipeClose'); if (c) c.addEventListener('click', closeModal);
    var ra = $('pipeRunAdryan'); if (ra) ra.addEventListener('click', runAdryan);
    var r = $('pipeRun'); if (r) r.addEventListener('click', run);
    var ru = $('pipeRunUpload'); if (ru) ru.addEventListener('click', runUpload);
    var rm = $('pipeRunMaestro'); if (rm) rm.addEventListener('click', runMaestro);
    var rsbtn = $('pipeResetBtn'); if (rsbtn) rsbtn.addEventListener('click', forceReset);
    var fi = $('pipeFile'); if (fi) fi.addEventListener('change', function () {
      var nm = tieneArchivo() ? fi.files[0].name : 'ningún archivo';
      if ($('pipeFileName')) $('pipeFileName').textContent = nm;
      setBusy(false);
    });
    var ad = $('pipeAdryan'); if (ad) ad.addEventListener('click', function (ev) {
      if (ad.getAttribute('data-empty')) {
        ev.preventDefault();
        alert('Configura la URL de Adryan en PIPELINE/motor/config.json → integracion_front.url_adryan');
      }
    });
    var rl = $('pipeReload'); if (rl) rl.addEventListener('click', function () { location.reload(); });
    var m = $('pipeModal'); if (m) m.addEventListener('click', function (ev) { if (ev.target === m) closeModal(); });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', wire);
  else wire();
})();


/* ============================================================================
 * Panel "Avance de la meta de vacaciones" (global + por Business Partner)
 * ==========================================================================*/
(function () {
  'use strict';
  function $(id) { return document.getElementById(id); }
  function fpct(v) { return (v == null || isNaN(v)) ? '—' : (v * 100).toFixed(1) + '%'; }
  function fnum(v) { return (v == null || isNaN(v)) ? '—' : Math.round(v).toLocaleString('es-PE'); }
  function esc(s) { return (s == null ? '' : String(s)).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function color(v) { if (v == null) return '#9aa7b2'; if (v >= 0.66) return '#1f9d55'; if (v >= 0.33) return '#e08a1e'; return '#c0392b'; }

  // BPs incorporados recientemente: el toggle decide si entran al listado y al
  // calculo de "cumplimiento BPs" mostrado en el panel global.
  var BPS_NUEVOS = ['César Reyes', 'Gabriel Chang'];
  var LS_KEY = 'vac_avance_incluir_nuevos';
  function incluirNuevos() {
    var v = null;
    try { v = localStorage.getItem(LS_KEY); } catch (e) {}
    return v === null ? true : v === '1';
  }
  function setIncluirNuevos(b) {
    try { localStorage.setItem(LS_KEY, b ? '1' : '0'); } catch (e) {}
  }

  var _ultimoPayload = null; // { global, por_bp } de la ultima carga, para re-render sin refetch

  function recalcularGlobalBP(por_bp_filtrado) {
    var meta = 0, reg = 0;
    por_bp_filtrado.forEach(function (b) { meta += (b.meta || 0); reg += (b.registro || 0); });
    return meta ? (reg / meta) : null;
  }

  function pintarGlobal(G, por_bp_filtrado) {
    // Avance = el mismo total crudo que Excel (toda la base, sin excluir colegio/cesados/fecha).
    var avanceTotal = G.avance_todo;
    var metaTotal = G.meta_total;
    var regTotal  = G.registrado_total;

    $('avTodo').textContent = fpct(avanceTotal);
    $('avReg').textContent = fnum(regTotal);
    $('avMeta').textContent = fnum(metaTotal);
    $('avBP').textContent = fpct(recalcularGlobalBP(por_bp_filtrado));

    var pct = Math.max(0, Math.min(100, (avanceTotal || 0) * 100));
    setTimeout(function () { if ($('avTodoBar')) $('avTodoBar').style.width = pct + '%'; }, 60);

    // Ya no hay segmento aparte para colegio: esta incluido en el numero principal.
    var barColegio = $('avTodoBarColegio');
    if (barColegio) {
      barColegio.title = '';
      setTimeout(function () { barColegio.style.width = '0%'; }, 60);
    }
  }

  function open() {
    if ($('detModal')) $('detModal').classList.remove('hidden');
    if ($('avBpList')) $('avBpList').innerHTML = '<div style="padding:14px;color:#888">Cargando…</div>';
    fetch('/api/vacaciones/avance').then(function (r) { return r.json(); }).then(function (j) {
      if (!j.ok) { $('avBpList').innerHTML = '<div style="color:#c0392b;padding:14px">' + esc(j.error || 'Error') + '</div>'; return; }
      _ultimoPayload = { global: j.global || {}, por_bp: j.por_bp || [] };
      renderResumen();
    }).catch(function (e) { if ($('avBpList')) $('avBpList').innerHTML = '<div style="color:#c0392b;padding:14px">' + esc('' + e) + '</div>'; });
  }

  // Vista resumen: una tarjeta por BP, clic para ver el detalle por persona.
  // Respeta el toggle "Incluir César Reyes y Gabriel Chang" (filtra lista + recalcula % global).
  function renderResumen() {
    if (!_ultimoPayload) return;
    var ttl = $('avBpTitle'); if (ttl) ttl.textContent = 'Avance por Business Partner';
    var tg = $('avToggleNuevos'); if (tg) tg.style.display = '';
    var incluir = incluirNuevos();
    var por_bp = incluir ? _ultimoPayload.por_bp
      : _ultimoPayload.por_bp.filter(function (b) { return BPS_NUEVOS.indexOf(b.hrbp) === -1; });
    pintarGlobal(_ultimoPayload.global, por_bp);
    var html = por_bp.map(function (b) {
      var c = color(b.avance), w = Math.min(100, (b.avance || 0) * 100);
      return '<div class="av-bp av-bp--click" data-bp="' + esc(b.hrbp) + '" title="Ver detalle de ' + esc(b.hrbp) + '">' +
        '<div class="av-bp-head"><span class="av-bp-name">' + esc(b.hrbp) + ' <span class="av-bp-go">›</span></span>' +
        '<span class="av-bp-pct" style="color:' + c + '">' + fpct(b.avance) + '</span></div>' +
        '<div class="av-bar"><div class="av-bar-fill" style="width:' + w + '%;background:' + c + '"></div></div>' +
        '<div class="av-bp-foot">' + fnum(b.registro) + ' de ' + fnum(b.meta) + ' días · ' + fnum(b.n) + ' personas</div>' +
        '</div>';
    }).join('');
    $('avBpList').innerHTML = html || '<div style="padding:14px;color:#888">Sin datos por BP.</div>';
    Array.prototype.forEach.call($('avBpList').querySelectorAll('.av-bp--click'), function (el) {
      el.addEventListener('click', function () { abrirDetalle(el.getAttribute('data-bp')); });
    });
  }

  // Vista detalle: colaboradores del BP con sus registros reales de vacaciones.
  function abrirDetalle(bp) {
    var ttl = $('avBpTitle'); if (ttl) ttl.textContent = 'Detalle — ' + bp;
    var tg = $('avToggleNuevos'); if (tg) tg.style.display = 'none';
    $('avBpList').innerHTML = '<div style="padding:14px;color:#888">Cargando detalle…</div>';
    fetch('/api/vacaciones/bp_detalle?bp=' + encodeURIComponent(bp)).then(function (r) { return r.json(); }).then(function (j) {
      if (!j.ok) { $('avBpList').innerHTML = barraVolver() + '<div style="color:#c0392b;padding:14px">' + esc(j.error || 'Error') + '</div>'; wireVolver(); return; }
      var d = j.detalle || {}, ppl = d.personas || [];
      var head = '<div class="av-det-sum">' +
        '<b>' + esc(d.hrbp) + '</b> · ' + fnum(d.n) + ' personas · ' +
        fnum(d.registro) + ' de ' + fnum(d.meta) + ' días (' + fpct(d.avance) + ')</div>';
      var rows = ppl.map(function (p) {
        var c = color(p.avance), regs = p.registros || [];
        var det = regs.length
          ? '<table class="av-reg"><thead><tr><th>Inicio</th><th>Término</th><th>Días</th><th>Motivo</th><th>Periodo</th></tr></thead><tbody>' +
            regs.map(function (g) {
              return '<tr><td>' + esc(g.inicio) + '</td><td>' + esc(g.termino) + '</td><td style="text-align:right">' +
                fnum(g.dias) + '</td><td>' + esc(g.motivo) + '</td><td>' + esc(g.periodo) + '</td></tr>';
            }).join('') + '</tbody></table>'
          : '<div class="av-reg-none">Sin registros de vacaciones tomadas.</div>';
        return '<div class="av-pers">' +
          '<div class="av-pers-head">' +
            '<div class="av-pers-id"><span class="av-pers-name">' + esc(p.nombre) + '</span>' +
              '<span class="av-pers-sub">' + esc(p.area || p.departamento || '') + (p.puesto ? ' · ' + esc(p.puesto) : '') + '</span></div>' +
            '<div class="av-pers-num"><span style="color:' + c + ';font-weight:800">' + fpct(p.avance) + '</span>' +
              '<span class="av-pers-sub">' + fnum(p.gozado) + ' / ' + fnum(p.meta) + ' días</span></div>' +
          '</div>' + det + '</div>';
      }).join('');
      $('avBpList').innerHTML = barraVolver() + head + (rows || '<div style="padding:14px;color:#888">Sin personas.</div>');
      wireVolver();
    }).catch(function (e) { $('avBpList').innerHTML = barraVolver() + '<div style="color:#c0392b;padding:14px">' + esc('' + e) + '</div>'; wireVolver(); });
  }
  function barraVolver() { return '<button type="button" class="av-volver" id="avVolver">← Volver a Business Partners</button>'; }
  function wireVolver() { var v = $('avVolver'); if (v) v.addEventListener('click', open); }

  function close() { if ($('detModal')) $('detModal').classList.add('hidden'); }
  function wire() {
    var b = $('btnDetalleVac'); if (b) b.addEventListener('click', open);
    var b2 = $('kpiAlerta'); if (b2) b2.addEventListener('click', open);
    var c = $('detClose'); if (c) c.addEventListener('click', close);
    var m = $('detModal'); if (m) m.addEventListener('click', function (ev) { if (ev.target === m) close(); });
    var chk = $('avToggleNuevosChk');
    if (chk) {
      chk.checked = incluirNuevos();
      chk.addEventListener('change', function () {
        setIncluirNuevos(chk.checked);
        renderResumen(); // re-pinta con los datos ya cargados, sin refetch
      });
    }
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', wire);
  else wire();
})();


/* ============================================================================
 * Panel "Ranking por Unidades de Negocio" — Embudo jerárquico
 * Nivel 1: Unidad de Negocio (Sucursal)
 * Nivel 2: Gerencia (Departamento)
 * Nivel 3: Áreas dentro de la Gerencia
 * Nivel 4: Personas dentro del área
 * ==========================================================================*/
(function () {
  'use strict';
  function $(id) { return document.getElementById(id); }
  function fpct(v) { return (v == null || isNaN(v)) ? '—' : (v * 100).toFixed(1) + '%'; }
  function fnum(v) { return (v == null || isNaN(v)) ? '—' : Math.round(v).toLocaleString('es-PE'); }
  function esc(s) { return (s == null ? '' : String(s)).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function color(v) { if (v == null) return '#9aa7b2'; if (v >= 0.9) return '#1f9d55'; if (v >= 0.7) return '#e08a1e'; return '#c0392b'; }
  function normStr(s) { return (s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,''); }

  // ── Estado de navegación ──────────────────────────────────────────────────
  var _rawPersonas = null;
  var _udnMap      = null;  // { udnKey -> { nombre, meta, gozado, n, avance, gerencias: { gerKey -> { nombre, meta, gozado, n, avance, areas: { areaKey -> { ... personas } } } } } }
  
  var _searchTimer = null;

  // ── Construcción de la estructura jerárquica ──────────────────────────────
  function buildUdnMap(personas) {
    var map = {};
    personas.forEach(function(p) {
      var udn  = (p.unidad_negocio || 'SIN UNIDAD DE NEGOCIO').toUpperCase().trim();
      var ger  = (p.gerencia || p.departamento || 'SIN GERENCIA').toUpperCase().trim();
      var area = (p.area || '').trim();
      var areaKey = area || '(Sin área asignada)';

      if (!map[udn]) map[udn] = { nombre: udn, meta: 0, gozado: 0, n: 0, avance: 0, gerencias: {} };
      var u = map[udn];
      u.meta += Number(p.meta || 0); u.gozado += Number(p.gozado || 0); u.n++;

      if (!u.gerencias[ger]) u.gerencias[ger] = { nombre: ger, meta: 0, gozado: 0, n: 0, avance: 0, areas: {}, udn: udn };
      var g = u.gerencias[ger];
      g.meta += Number(p.meta || 0); g.gozado += Number(p.gozado || 0); g.n++;

      if (!g.areas[areaKey]) g.areas[areaKey] = { nombre: areaKey, sinArea: !area, meta: 0, gozado: 0, n: 0, avance: 0, personas: [], ger: ger, udn: udn };
      var a = g.areas[areaKey];
      a.meta += Number(p.meta || 0); a.gozado += Number(p.gozado || 0); a.n++;

      p._udn = udn; p._ger = ger; p._area = areaKey;
      a.personas.push(p);
    });

    Object.values(map).forEach(function(u) {
      u.avance = u.meta > 0 ? u.gozado / u.meta : 0;
      Object.values(u.gerencias).forEach(function(g) {
        g.avance = g.meta > 0 ? g.gozado / g.meta : 0;
        Object.values(g.areas).forEach(function(a) { 
          a.avance = a.meta > 0 ? a.gozado / a.meta : 0; 
        });
      });
    });
    return map;
  }

  // ── Carga de datos ────────────────────────────────────────────────────────
  function openAreaRank() {
    if ($('areaRankModal')) $('areaRankModal').classList.remove('hidden');
    if (_rawPersonas) { 
        populateDropdowns();
        renderLevel(); 
        return; 
    }
    if ($('areaRankList')) $('areaRankList').innerHTML = '<div style="padding:18px;color:#888">Cargando datos…</div>';
    
    fetch('/api/vacaciones/avance').then(function(r){ return r.json(); }).then(function(j){
      if (!j.ok) { $('areaRankList').innerHTML = '<div style="color:#c0392b;padding:14px">' + esc(j.error||'Error') + '</div>'; return; }
      var allPersonas = [];
      var fetches = (j.por_bp||[]).map(function(bp){
        return fetch('/api/vacaciones/bp_detalle?bp=' + encodeURIComponent(bp.hrbp))
          .then(function(r){ return r.json(); })
          .then(function(d){
            if (!d.ok) return;
            (d.detalle||{}).personas.forEach(function(p){ p.hrbp = bp.hrbp; allPersonas.push(p); });
          }).catch(function(){});
      });
      Promise.all(fetches).then(function(){
        _rawPersonas = allPersonas;
        _udnMap = buildUdnMap(allPersonas);
        populateDropdowns();
        renderLevel();
      });
    }).catch(function(e){ if ($('areaRankList')) $('areaRankList').innerHTML = '<div style="color:#c0392b;padding:14px">' + esc(''+e) + '</div>'; });
  }

  function getSel(id) { var e = $(id); return e ? e.value : ''; }
  function setSel(id, val) { var e = $(id); if (e) e.value = val; }

  function populateDropdowns() {
    if (!_udnMap) return;
    var fUdn = getSel('arFilterUdn');
    var fGer = getSel('arFilterGerencia');
    var fArea = getSel('arFilterArea');

    // UDNs
    var udns = Object.keys(_udnMap).sort();
    var selUdn = $('arFilterUdn');
    if (selUdn) {
      selUdn.innerHTML = '<option value="">Todas las Unidades de Negocio</option>' + 
        udns.map(function(u){ return '<option value="'+esc(u)+'">'+esc(u)+'</option>'; }).join('');
      selUdn.value = fUdn;
    }

    // Gerencias
    var gerencias = new Set();
    udns.forEach(function(u) {
      if (!fUdn || fUdn === u) {
        Object.keys(_udnMap[u].gerencias).forEach(function(g) { gerencias.add(g); });
      }
    });
    var selGer = $('arFilterGerencia');
    if (selGer) {
      var gArr = Array.from(gerencias).sort();
      selGer.innerHTML = '<option value="">Todas las Gerencias</option>' + 
        gArr.map(function(g){ return '<option value="'+esc(g)+'">'+esc(g)+'</option>'; }).join('');
      if (gArr.includes(fGer)) selGer.value = fGer;
      else fGer = '';
    }

    // Areas
    var areas = new Set();
    udns.forEach(function(u) {
      if (!fUdn || fUdn === u) {
        Object.keys(_udnMap[u].gerencias).forEach(function(g) {
          if (!fGer || fGer === g) {
            Object.keys(_udnMap[u].gerencias[g].areas).forEach(function(a) { areas.add(a); });
          }
        });
      }
    });
    var selArea = $('arFilterArea');
    if (selArea) {
      var aArr = Array.from(areas).sort();
      selArea.innerHTML = '<option value="">Todas las Áreas</option>' + 
        aArr.map(function(a){ return '<option value="'+esc(a)+'">'+esc(a)+'</option>'; }).join('');
      if (aArr.includes(fArea)) selArea.value = fArea;
      else fArea = '';
    }
  }

  function onFilterChange(level) {
    if (level === 'udn') {
      setSel('arFilterGerencia', '');
      setSel('arFilterArea', '');
    } else if (level === 'gerencia') {
      setSel('arFilterArea', '');
    }
    populateDropdowns();
    renderLevel();
  }

  // ── Renderizado por nivel ─────────────────────────────────────────────────
  function getSearch() { return normStr(($('areaRankSearch')||{}).value||''); }
  function getSortAsc() { return !!($('areaRankSortAsc')&&$('areaRankSortAsc').checked); }

  function renderLevel() {
    var fUdn = getSel('arFilterUdn');
    var fGer = getSel('arFilterGerencia');
    var fArea = getSel('arFilterArea');
    var q = getSearch();

    if (fArea || (fGer && fUdn && fArea)) {
        renderPersonas(fUdn, fGer, fArea);
    } else if (fGer || (fUdn && fGer)) {
        renderAreas(fUdn, fGer);
    } else if (fUdn) {
        renderGerencias(fUdn);
    } else if (q) {
        // If there is a search query but no filters, default to showing Personas across all
        renderPersonas('', '', '');
    } else {
        renderUdns();
    }
  }

  // ── Nivel 1: Unidades de Negocio ──────────────────────────────────────────
  function renderUdns() {
    var list = $('areaRankList'); if (!list||!_udnMap) return;
    var q = getSearch(); var asc = getSortAsc();
    var udns = Object.values(_udnMap).filter(function(u){
      if (!q) return true;
      return normStr(u.nombre).includes(q); // Simplified search for top level
    });
    udns.sort(function(a,b){ return asc ? a.avance-b.avance : b.avance-a.avance; });
    var rc = $('areaRankResultCount');
    if (rc) rc.textContent = udns.length + ' unidades de negocio' + (q ? ' (filtradas)' : '');
    if (!udns.length){ list.innerHTML='<div style="padding:14px;color:#888">Sin resultados.</div>'; return; }
    
    list.innerHTML = udns.map(function(u,i){
      var c=color(u.avance), w=Math.min(100,(u.avance||0)*100);
      var nGer = Object.keys(u.gerencias).length;
      return '<div class="area-rank-row av-bp--click" data-lvl="udn" data-val="'+esc(u.nombre)+'">' +
        '<div class="area-rank-head">' +
          '<span class="area-rank-name"><b style="color:#0f6ea5;margin-right:4px">#'+(i+1)+'</b>'+esc(u.nombre)+' <span class="av-bp-go">›</span></span>' +
          '<div style="display:flex;align-items:center;gap:8px">' +
            '<button class="ar-dl-btn" data-dl-udn="'+esc(u.nombre)+'" title="Descargar Excel">⬇</button>' +
            '<span class="area-rank-pct" style="color:'+c+'">'+fpct(u.avance)+'</span>' +
          '</div>' +
        '</div>' +
        '<div class="area-rank-bar"><div class="area-rank-bar-fill" style="width:'+w.toFixed(1)+'%;background:'+c+'"></div></div>' +
        '<div class="area-rank-foot">'+fnum(u.gozado)+' / '+fnum(u.meta)+' días · '+fnum(u.n)+' personas · '+nGer+' gerencias</div>' +
      '</div>';
    }).join('');
    bindListEvents(list);
  }

  // ── Nivel 2: Gerencias ────────────────────────────────────────────────────
  function renderGerencias(fUdn) {
    var list = $('areaRankList'); if (!list||!_udnMap) return;
    var q = getSearch(); var asc = getSortAsc();
    
    var gers = [];
    Object.values(_udnMap).forEach(function(u) {
        if (!fUdn || u.nombre === fUdn) {
            Object.values(u.gerencias).forEach(function(g) {
                if (!q || normStr(g.nombre).includes(q) || normStr(u.nombre).includes(q)) gers.push(g);
            });
        }
    });

    gers.sort(function(a,b){ return asc ? a.avance-b.avance : b.avance-a.avance; });
    var rc = $('areaRankResultCount');
    if (rc) rc.textContent = gers.length + ' gerencias' + (fUdn ? ' en ' + fUdn : '') + (q ? ' (filtradas)' : '');
    
    var btnVolver = '<button class="av-volver" style="margin-bottom:12px; font-weight:600; display:block" onclick="document.getElementById(\'arFilterUdn\').value=\'\'; document.getElementById(\'arFilterUdn\').dispatchEvent(new Event(\'change\'))">⬅ Volver a Todas las Unidades</button>';
    if (!gers.length){ list.innerHTML = btnVolver + '<div style="padding:14px;color:#888">Sin gerencias.</div>'; return; }
    
    list.innerHTML = btnVolver + gers.map(function(g,i){
      var c=color(g.avance), w=Math.min(100,(g.avance||0)*100);
      var nAreas = Object.keys(g.areas).length;
      return '<div class="area-rank-row area-rank-row--ger av-bp--click" data-lvl="gerencia" data-udn="'+esc(g.udn)+'" data-val="'+esc(g.nombre)+'">' +
        '<div class="area-rank-head">' +
          '<span class="area-rank-name"><b style="color:#0f6ea5;margin-right:4px">#'+(i+1)+'</b>'+esc(g.nombre)+' <span class="av-bp-go">›</span><br><span style="font-size:10px;color:#7a8a99;font-weight:normal">'+esc(g.udn)+'</span></span>' +
          '<div style="display:flex;align-items:center;gap:8px">' +
            '<button class="ar-dl-btn" data-dl-udn="'+esc(g.udn)+'" data-dl-ger="'+esc(g.nombre)+'" title="Descargar Excel">⬇</button>' +
            '<span class="area-rank-pct" style="color:'+c+'">'+fpct(g.avance)+'</span>' +
          '</div>' +
        '</div>' +
        '<div class="area-rank-bar"><div class="area-rank-bar-fill" style="width:'+w.toFixed(1)+'%;background:'+c+'"></div></div>' +
        '<div class="area-rank-foot">'+fnum(g.gozado)+' / '+fnum(g.meta)+' días · '+fnum(g.n)+' personas · '+nAreas+' áreas</div>' +
      '</div>';
    }).join('');
    bindListEvents(list);
  }

  // ── Nivel 3: Áreas ────────────────────────────────────────────────────────
  function renderAreas(fUdn, fGer) {
    var list = $('areaRankList'); if (!list||!_udnMap) return;
    var q = getSearch(); var asc = getSortAsc();
    
    var areas = [];
    Object.values(_udnMap).forEach(function(u) {
        if (!fUdn || u.nombre === fUdn) {
            Object.values(u.gerencias).forEach(function(g) {
                if (!fGer || g.nombre === fGer) {
                    Object.values(g.areas).forEach(function(a) {
                        if (!q || normStr(a.nombre).includes(q) || normStr(g.nombre).includes(q)) areas.push(a);
                    });
                }
            });
        }
    });

    areas.sort(function(a,b){ return asc ? a.avance-b.avance : b.avance-a.avance; });
    var rc = $('areaRankResultCount');
    if (rc) rc.textContent = areas.length + ' áreas' + (fGer ? ' en ' + fGer : '') + (q ? ' (filtradas)' : '');
    
    var btnVolver = '<button class="av-volver" style="margin-bottom:12px; font-weight:600; display:block" onclick="document.getElementById(\'arFilterGerencia\').value=\'\'; document.getElementById(\'arFilterGerencia\').dispatchEvent(new Event(\'change\'))">⬅ Volver a ' + esc(fUdn) + '</button>';
    if (!areas.length){ list.innerHTML = btnVolver + '<div style="padding:14px;color:#888">Sin áreas.</div>'; return; }
    
    list.innerHTML = btnVolver + areas.map(function(a,i){
      var c=color(a.avance), w=Math.min(100,(a.avance||0)*100);
      var lbl = a.sinArea ? '<em style="color:#9aa7b2">'+esc(a.nombre)+'</em>' : esc(a.nombre);
      return '<div class="area-rank-row area-rank-row--area av-bp--click" data-lvl="area" data-udn="'+esc(a.udn)+'" data-ger="'+esc(a.ger)+'" data-val="'+esc(a.nombre)+'">' +
        '<div class="area-rank-head">' +
          '<span class="area-rank-name"><b style="color:#0f6ea5;margin-right:4px">#'+(i+1)+'</b>'+lbl+' <span class="av-bp-go">›</span><br><span style="font-size:10px;color:#7a8a99;font-weight:normal">'+esc(a.ger)+'</span></span>' +
          '<div style="display:flex;align-items:center;gap:8px">' +
            '<button class="ar-dl-btn" data-dl-udn="'+esc(a.udn)+'" data-dl-ger="'+esc(a.ger)+'" data-dl-area="'+esc(a.nombre)+'" title="Descargar Excel">⬇</button>' +
            '<span class="area-rank-pct" style="color:'+c+'">'+fpct(a.avance)+'</span>' +
          '</div>' +
        '</div>' +
        '<div class="area-rank-bar"><div class="area-rank-bar-fill" style="width:'+w.toFixed(1)+'%;background:'+c+'"></div></div>' +
        '<div class="area-rank-foot">'+fnum(a.gozado)+' / '+fnum(a.meta)+' días · '+fnum(a.n)+' persona(s)</div>' +
      '</div>';
    }).join('');
    bindListEvents(list);
  }

  // ── Nivel 4: Personas ─────────────────────────────────────────────────────
  function renderPersonas(fUdn, fGer, fArea) {
    var list = $('areaRankList'); if (!list||!_udnMap||!_rawPersonas) return;
    var q = getSearch();
    
    var filtered = _rawPersonas.filter(function(p) {
        if (fUdn && p._udn !== fUdn) return false;
        if (fGer && p._ger !== fGer) return false;
        if (fArea && p._area !== fArea) return false;
        if (q && !(normStr(p.nombre).includes(q) || normStr(p._ger).includes(q) || normStr(p._area).includes(q))) return false;
        return true;
    });

    // Calcular KPI summary del grupo filtrado
    var sumMeta=0, sumGoz=0, n=filtered.length, cumplieron=0, sinIni=0;
    filtered.forEach(function(p) {
        sumMeta += Number(p.meta||0);
        sumGoz += Number(p.gozado||0);
        if ((p.avance||0)>=1) cumplieron++;
        else if (!p.gozado||p.gozado===0) sinIni++;
    });
    var enProceso = n - cumplieron - sinIni;
    var avgAvance = sumMeta > 0 ? (sumGoz / sumMeta) : 0;
    var c0 = color(avgAvance), pctNum = Math.min(100, avgAvance*100);

    var title = fArea || fGer || fUdn || 'Todas las personas';
    var subTitle = fArea ? fGer : (fGer ? fUdn : '');
    var textoVolver = fArea ? fGer : (fGer ? fUdn : 'Todas las Unidades');
    var btnVolver = '<button class="av-volver" style="margin-bottom:12px; font-weight:600; display:block" onclick="document.getElementById(\'arFilterArea\').value=\'\'; document.getElementById(\'arFilterArea\').dispatchEvent(new Event(\'change\'))">⬅ Volver a ' + esc(textoVolver) + '</button>';

    var html = '<div class="av-det-header">' +
      btnVolver +
      '<div class="av-det-title"><span class="av-det-area-name">'+esc(title)+'</span><span class="av-det-ger">'+esc(subTitle)+'</span></div>' +
      '<div class="av-det-kpis">' +
        '<span class="av-det-kpi"><b style="color:'+c0+'">'+fpct(avgAvance)+'</b><small>avance</small></span>' +
        '<span class="av-det-kpi"><b>'+fnum(n)+'</b><small>personas</small></span>' +
        '<span class="av-det-kpi"><b style="color:#1f9d55">'+fnum(cumplieron)+'</b><small>cumplieron</small></span>' +
        '<span class="av-det-kpi"><b style="color:#e08a1e">'+fnum(enProceso)+'</b><small>en proceso</small></span>' +
        '<span class="av-det-kpi"><b style="color:#c0392b">'+fnum(sinIni)+'</b><small>sin iniciar</small></span>' +
      '</div>' +
      '<div class="av-det-bar"><div class="av-det-bar-fill" style="width:'+pctNum.toFixed(1)+'%;background:'+c0+'"></div></div>' +
      '<div class="av-det-days">'+fnum(sumGoz)+' días gozados de '+fnum(sumMeta)+' días meta</div>' +
    '</div>';

    var sorted = filtered.slice().sort(function(x,y){
        var sx=(x.avance||0)>=1?2:(!x.gozado||x.gozado===0?0:1);
        var sy=(y.avance||0)>=1?2:(!y.gozado||y.gozado===0?0:1);
        if (sx!==sy) return sx-sy;
        return (Number(y.meta||0)-Number(y.gozado||0))-(Number(x.meta||0)-Number(x.gozado||0));
    });

    var rows = sorted.map(function(p){
      var c=color(p.avance), regs=p.registros||[];
      var saldo=Math.max(0,Number(p.meta||0)-Number(p.gozado||0));
      var badge=(p.avance||0)>=1?'<span class="av-badge av-badge-ok">✓ Cumplió</span>':
        (!p.gozado||p.gozado===0?'<span class="av-badge av-badge-red">Sin iniciar</span>':'<span class="av-badge av-badge-yellow">En proceso</span>');
      var det=regs.length
        ?'<table class="av-reg"><thead><tr><th>Inicio</th><th>Término</th><th>Días</th><th>Motivo</th><th>Periodo</th></tr></thead><tbody>'+
          regs.map(function(g){ return '<tr><td>'+esc(g.inicio)+'</td><td>'+esc(g.termino)+'</td><td style="text-align:right;font-weight:600">'+fnum(g.dias)+'</td><td>'+esc(g.motivo)+'</td><td>'+esc(g.periodo)+'</td></tr>'; }).join('')+'</tbody></table>'
        :'<div class="av-reg-none">Sin registros de vacaciones tomadas este trimestre.</div>';
      return '<div class="av-pers"><div class="av-pers-head"><div class="av-pers-id">'+
        '<span class="av-pers-name">'+esc(p.nombre)+'</span>'+badge+
        '<span class="av-pers-sub">'+esc(p.puesto||'Sin puesto')+' · HRBP: '+esc(p.hrbp||'N/A')+'<br><span style="opacity:0.8">'+esc(p._udn)+' / '+esc(p._ger)+'</span></span></div>'+
        '<div class="av-pers-num"><span style="color:'+c+';font-weight:800;font-size:1.1em">'+fpct(p.avance)+'</span>'+
        '<span class="av-pers-sub">'+fnum(p.gozado)+' goz. / '+fnum(p.meta)+' meta</span>'+
        (saldo>0?'<span class="av-pers-saldo">Saldo: '+fnum(saldo)+' d.</span>':'')+
        '<button style="margin-top:8px; padding:4px 8px; font-size:12px; cursor:pointer; background:#f0f9ff; color:#0369a1; border:1px solid #bae6fd; border-radius:4px; font-weight:bold" onclick="window.seleccionarColaboradorAlerta(\''+esc(p.nombre)+'\')">📩 Enviar Alerta Individual</button>' +
        '</div></div>'+det+'</div>';
    }).join('');

    var rc = $('areaRankResultCount');
    if (rc) rc.textContent = sorted.length + ' persona(s)' + (q?' (filtradas)':'');
    list.innerHTML = html + (rows||'<div style="padding:14px;color:#888">Sin personas.</div>');
  }

  // ── Descarga Excel ────────────────────────────────────────────────────────
  function descargarExcel(udn, ger, area) {
    var params = new URLSearchParams();
    if (udn)  params.set('unidad_negocio', udn);
    if (ger)  params.set('gerencia', ger);
    if (area) params.set('area', area);
    var q = ($('areaRankSearch')||{}).value||'';
    if (q) params.set('buscar', q);
    window.location.href = '/api/vacaciones/exportar_areas?' + params.toString();
  }

  // ── Event binding ─────────────────────────────────────────────────────────
  function bindListEvents(list) {
    list.querySelectorAll('.av-bp--click').forEach(function(el){
      el.addEventListener('click', function(ev){
        if (ev.target.closest('.ar-dl-btn')) return;
        var lvl = el.dataset.lvl, val = el.dataset.val;
        if (lvl === 'udn') {
            setSel('arFilterUdn', val);
            onFilterChange('udn');
        } else if (lvl === 'gerencia') {
            setSel('arFilterUdn', el.dataset.udn);
            setSel('arFilterGerencia', val);
            onFilterChange('gerencia');
        } else if (lvl === 'area') {
            setSel('arFilterUdn', el.dataset.udn);
            setSel('arFilterGerencia', el.dataset.ger);
            setSel('arFilterArea', val);
            onFilterChange('area');
        }
      });
    });
    list.querySelectorAll('[data-dl-udn]').forEach(function(btn){
      btn.addEventListener('click', function(ev){ 
        ev.stopPropagation(); 
        descargarExcel(btn.dataset.dlUdn, btn.dataset.dlGer, btn.dataset.dlArea); 
      });
    });
  }

  function closeAreaRank() { if ($('areaRankModal')) $('areaRankModal').classList.add('hidden'); }

  function wire() {
    var b = $('btnAreaRanking');      if (b)  b.addEventListener('click', openAreaRank);
    var c = $('areaRankClose');       if (c)  c.addEventListener('click', closeAreaRank);
    var m = $('areaRankModal');       if (m)  m.addEventListener('click', function(ev){ if (ev.target===m) closeAreaRank(); });
    var sa = $('areaRankSortAsc');    if (sa) sa.addEventListener('change', renderLevel);
    var dl = $('areaRankDownloadAll');if (dl) dl.addEventListener('click', function(){ 
        descargarExcel(getSel('arFilterUdn'), getSel('arFilterGerencia'), getSel('arFilterArea')); 
    });
    
    var fUdn = $('arFilterUdn'); if (fUdn) fUdn.addEventListener('change', function() { onFilterChange('udn'); });
    var fGer = $('arFilterGerencia'); if (fGer) fGer.addEventListener('change', function() { onFilterChange('gerencia'); });
    var fArea = $('arFilterArea'); if (fArea) fArea.addEventListener('change', function() { onFilterChange('area'); });

    // Búsqueda
    var sr = $('areaRankSearch');
    if (sr) {
      sr.addEventListener('keyup', function(){ clearTimeout(_searchTimer); _searchTimer = setTimeout(renderLevel, 250); });
      sr.addEventListener('change', function(){ clearTimeout(_searchTimer); _searchTimer = setTimeout(renderLevel, 250); });
    }
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', wire);
  else wire();
})();

window.seleccionarColaboradorAlerta = function(nombre) {
    var modal = document.getElementById('areaRankModal');
    if (modal) modal.classList.add('hidden'); 
    
    if (typeof wizardGo === 'function') {
        wizardGo(2); // Asumiendo que el paso de seleccionar destino es el 1 o 2.
    }
    
    // Tratamos de buscar la pestaña o modal de colab_buscar (Panel Individual)
    var btnIndividual = document.getElementById('btnTipoInd');
    if (btnIndividual) {
        btnIndividual.click();
    }
    
    var inpColab = document.getElementById('colab_buscar');
    if (inpColab) {
        inpColab.value = nombre;
        var btnColab = document.getElementById('btnColab_Buscar');
        if (btnColab) btnColab.click();
    }
    
    // Fallback: si hay un general de inpNombre
    var inpInd = document.getElementById('inpNombre');
    if (inpInd) {
        inpInd.value = nombre;
        inpInd.dispatchEvent(new Event('change'));
        inpInd.dispatchEvent(new Event('input'));
    }
    
    console.log("Seleccionado colaborador: " + nombre);
};

