const { app, BrowserWindow } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const http  = require('http');
const https = require('https');
const fs    = require('fs');

let mainWindow;
let pythonProcess;

// ─── Repositorio GitHub ───────────────────────────────────────────────────────
const GITHUB_REPO  = 'AtraccionDeTalento/Sistema-Vacaciones';
const GITHUB_BRANCH = 'main';
const RAW_BASE = `https://raw.githubusercontent.com/${GITHUB_REPO}/${GITHUB_BRANCH}`;
const API_COMMIT = `https://api.github.com/repos/${GITHUB_REPO}/commits/${GITHUB_BRANCH}`;

// Archivos que se actualizan desde GitHub (no incluir pa_config ni datos del usuario)
const ARCHIVOS_ACTUALIZABLES = [
  'servidor.py',
  'index_vacaciones.html',
  'enviar_cola_outlook.py',
  '_bp_map.py',
  'requirements.txt',
  'assets/js/app_completo.js',
  'assets/js/pipeline_vac.js',
  'assets/css/styles.css',
  'PIPELINE/motor/pipeline.py',
  'PIPELINE/motor/vac_lib.py',
  'PIPELINE/motor/config.json',
  'PIPELINE/bot_adryan/bot_adryan.py',
  'PIPELINE/bot_adryan/bot_maestro.py',
  'PIPELINE/bot_adryan/guardar_password.py',
];

// ─── Pantalla de carga ────────────────────────────────────────────────────────
function loadingPage(mensaje, subMensaje = '', esError = false) {
  const color  = esError ? '#fef2f2' : '#f0f9ff';
  const titCol = esError ? '#b91c1c' : '#0f6ea5';
  const subCol = esError ? '#dc2626' : '#64748b';
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <style>
    *{margin:0;padding:0;box-sizing:border-box}
    body{display:flex;flex-direction:column;justify-content:center;align-items:center;
         height:100vh;font-family:'Segoe UI',sans-serif;background:${color};gap:18px}
    .icon{font-size:56px}
    h2{color:${titCol};font-size:22px;font-weight:700;text-align:center;max-width:480px}
    p{color:${subCol};font-size:14px;text-align:center;max-width:480px;line-height:1.6}
    .bar{width:320px;height:6px;background:#e2e8f0;border-radius:99px;overflow:hidden}
    .fill{height:100%;background:${esError ? '#ef4444' : '#0ea5e9'};border-radius:99px;
          ${esError ? '' : 'animation:slide 1.4s ease-in-out infinite'};width:${esError ? '100%' : '40%'}}
    @keyframes slide{0%{transform:translateX(-100%)}100%{transform:translateX(900%)}}
  </style></head><body>
  <div class="icon">${esError ? '⚠️' : '🌴'}</div>
  <h2>${mensaje}</h2>
  ${subMensaje ? `<p>${subMensaje}</p>` : ''}
  ${!esError ? '<div class="bar"><div class="fill"></div></div>' : ''}
  </body></html>`;
  return 'data:text/html;base64,' + Buffer.from(html).toString('base64');
}

// ─── Ventana principal ────────────────────────────────────────────────────────
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280, height: 800,
    icon: path.join(__dirname, 'icon.png'),
    autoHideMenuBar: true,
    webPreferences: { nodeIntegration: false, contextIsolation: true }
  });
  mainWindow.maximize();
  mainWindow.loadURL(loadingPage('Sistema de Vacaciones USIL', 'Iniciando...'));
  mainWindow.on('closed', () => { mainWindow = null; });
}

// ─── Helpers de red ───────────────────────────────────────────────────────────
function httpsGet(url, timeoutMs = 30000) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, {
      headers: { 'User-Agent': 'SistemaVacaciones-USIL/1.0' },
      timeout: timeoutMs
    }, res => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        return httpsGet(res.headers.location, timeoutMs).then(resolve).catch(reject);
      }
      if (res.statusCode === 403) {
        return reject(new Error('GitHub rate limit (403)'));
      }
      if (res.statusCode !== 200) {
        return reject(new Error(`HTTP ${res.statusCode}`));
      }
      const chunks = [];
      res.on('data', d => chunks.push(d));
      res.on('end', () => resolve(Buffer.concat(chunks)));
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('Timeout')); });
  });
}

function httpsGetText(url) {
  return httpsGet(url).then(b => b.toString('utf8'));
}

// ─── AUTO-UPDATE desde GitHub ─────────────────────────────────────────────────
async function descargarConReintentos(url, intentos = 3, timeoutMs = 20000) {
  for (let i = 0; i < intentos; i++) {
    try {
      return await httpsGet(url, timeoutMs);
    } catch (e) {
      if (i === intentos - 1) throw e;
      console.log(`[UPDATE] Reintento ${i + 1} para ${url}: ${e.message}`);
      await new Promise(r => setTimeout(r, 1500 * (i + 1)));
    }
  }
}

async function verificarActualizacion(baseDir) {
  const versionFile = path.join(baseDir, '.version_commit');

  // Leer commit local guardado
  let commitLocal = '';
  try { commitLocal = fs.readFileSync(versionFile, 'utf8').trim(); } catch (_) {}

  // Consultar último commit en GitHub (con reintento)
  let commitRemoto = '';
  for (let intento = 0; intento < 3; intento++) {
    try {
      const json = await httpsGetText(API_COMMIT);
      const data = JSON.parse(json);
      commitRemoto = data.sha || '';
      break;
    } catch (e) {
      console.log(`[UPDATE] Error consultando GitHub (intento ${intento + 1}):`, e.message);
      if (intento < 2) await new Promise(r => setTimeout(r, 2000));
    }
  }

  if (!commitRemoto) {
    console.log('[UPDATE] Sin conexión o sin respuesta de GitHub — continúa con versión local.');
    return false;
  }

  if (commitRemoto === commitLocal) {
    console.log('[UPDATE] App al día:', commitRemoto.slice(0, 7));
    return false;
  }

  console.log(`[UPDATE] Nueva versión: ${commitRemoto.slice(0, 7)} (local: ${commitLocal.slice(0, 7) || 'ninguna'})`);
  if (mainWindow) mainWindow.loadURL(loadingPage(
    'Actualizando aplicación...',
    `Nueva versión disponible. Descargando archivos...<br><small style="opacity:.6">${commitRemoto.slice(0,7)}</small>`
  ));

  let actualizados = 0;
  let errores = 0;

  for (const archivo of ARCHIVOS_ACTUALIZABLES) {
    try {
      const url = `${RAW_BASE}/${encodeURIComponent(archivo).replace(/%2F/g, '/')}`;
      const contenido = await descargarConReintentos(url, 3, 25000);
      const destino = path.join(baseDir, archivo.replace(/\//g, path.sep));
      const dir = path.dirname(destino);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
      // Escribir a archivo temporal primero para evitar corrupción
      const tmp = destino + '.tmp';
      fs.writeFileSync(tmp, contenido);
      fs.renameSync(tmp, destino);
      actualizados++;
      console.log(`[UPDATE] ✓ ${archivo}`);
    } catch (e) {
      console.log(`[UPDATE] ✗ ${archivo}: ${e.message}`);
      errores++;
    }
  }

  // Solo guardar el commit si descargamos al menos la mayoría
  if (actualizados >= Math.ceil(ARCHIVOS_ACTUALIZABLES.length / 2)) {
    fs.writeFileSync(versionFile, commitRemoto, 'utf8');
    console.log(`[UPDATE] ✅ ${actualizados} archivos actualizados, ${errores} omitidos.`);
  } else {
    console.log(`[UPDATE] ⚠️ Solo ${actualizados}/${ARCHIVOS_ACTUALIZABLES.length} archivos descargados — no se guarda versión para reintentar en el próximo arranque.`);
  }

  return actualizados > 0;
}

// ─── Acceso directo en el Escritorio ─────────────────────────────────────────
function crearAccesoDirecto(baseDir) {
  try {
    const { execFile } = require('child_process');
    const exeRuta = app.getPath('exe');
    const escritorio = app.getPath('desktop');
    const lnk = path.join(escritorio, 'Sistema de Vacaciones USIL.lnk');
    if (fs.existsSync(lnk)) return; // ya existe

    // Usar PowerShell para crear el acceso directo
    const script = `
      $s=(New-Object -COM WScript.Shell).CreateShortcut('${lnk.replace(/\\/g, '\\\\')}');
      $s.TargetPath='${exeRuta.replace(/\\/g, '\\\\')}';
      $s.Description='Sistema de Vacaciones USIL';
      $s.WorkingDirectory='${baseDir.replace(/\\/g, '\\\\')}';
      $s.Save()
    `;
    execFile('powershell', ['-NoProfile', '-Command', script], err => {
      if (err) console.log('[SHORTCUT] Error:', err.message);
      else console.log('[SHORTCUT] Acceso directo creado en Escritorio');
    });
  } catch (e) {
    console.log('[SHORTCUT] No se pudo crear acceso directo:', e.message);
  }
}

// ─── Ejecutar comando Python ──────────────────────────────────────────────────
function runPythonCmd(pythonExe, args, cwd) {
  return new Promise((resolve, reject) => {
    const currentDir = cwd || (app.isPackaged ? path.dirname(app.getPath('exe')) : __dirname);
    const proc = spawn(pythonExe, args, { cwd: currentDir, shell: false });
    let out = '', err = '';
    proc.stdout.on('data', d => { out += d.toString(); });
    proc.stderr.on('data', d => { err += d.toString(); });
    proc.on('close', code => { code === 0 ? resolve(out) : reject(new Error(err || `Exit ${code}`)); });
    proc.on('error', reject);
  });
}

async function detectarPython() {
  for (const py of ['python', 'python3', 'py']) {
    try { await runPythonCmd(py, ['--version']); return py; } catch (_) {}
  }
  return null;
}

// ─── Setup + arranque del servidor ───────────────────────────────────────────
async function setupAndStartServer() {
  const baseDir = app.isPackaged ? path.dirname(app.getPath('exe')) : __dirname;
  const venvDir = app.isPackaged
    ? path.join(app.getPath('userData'), '.venv')
    : path.join(baseDir, '.venv');
  const venvPython  = path.join(venvDir, 'Scripts', 'python.exe');
  const serverScript = path.join(baseDir, 'servidor.py');
  const requirements = path.join(baseDir, 'requirements.txt');

  // Crear acceso directo al escritorio (silencioso, solo primera vez)
  if (app.isPackaged) crearAccesoDirecto(baseDir);

  // ── Auto-update ──────────────────────────────────────────────────────────
  if (mainWindow) mainWindow.loadURL(loadingPage(
    'Verificando actualizaciones...',
    'Consultando GitHub para ver si hay una versión más reciente...'
  ));
  try {
    await verificarActualizacion(baseDir);
  } catch (e) {
    console.log('[UPDATE] Error no fatal:', e.message);
  }

  // ── Arrancar Python ───────────────────────────────────────────────────────
  if (fs.existsSync(venvPython)) {
    console.log('[BOOT] .venv encontrado → arrancando servidor...');
    if (mainWindow) mainWindow.loadURL(loadingPage('Iniciando servidor...', 'Preparando el motor de datos...'));
    startPythonServer(venvPython, serverScript);
    return;
  }

  // Primera vez: buscar Python global
  if (mainWindow) mainWindow.loadURL(loadingPage(
    'Primera vez en este equipo',
    'Buscando Python instalado en el sistema...'
  ));

  const pythonGlobal = await detectarPython();
  if (!pythonGlobal) {
    if (mainWindow) mainWindow.loadURL(loadingPage(
      'Python no encontrado',
      'Este sistema requiere Python 3.10+.<br>Descárgalo desde <b>python.org</b>, instálalo marcando "Add to PATH" y vuelve a abrir la app.',
      true
    ));
    return;
  }

  // ¿Python global ya tiene las dependencias?
  if (mainWindow) mainWindow.loadURL(loadingPage('Verificando dependencias...', 'Comprobando librerías instaladas...'));
  try {
    await runPythonCmd(pythonGlobal, ['-c', 'import flask, pandas, openpyxl, win32com']);
    if (mainWindow) mainWindow.loadURL(loadingPage('Listo. Iniciando...', 'Abriendo el Sistema de Vacaciones USIL...'));
    startPythonServer(pythonGlobal, serverScript);
    return;
  } catch (_) {}

  // Crear entorno virtual
  if (mainWindow) mainWindow.loadURL(loadingPage(
    'Configurando entorno (primera vez)',
    'Creando entorno virtual Python...<br>Esto ocurre solo una vez en este equipo.'
  ));
  try {
    await runPythonCmd(pythonGlobal, ['-m', 'venv', venvDir]);
  } catch (e) {
    if (mainWindow) mainWindow.loadURL(loadingPage('Error al crear entorno', e.message, true));
    return;
  }

  // Instalar dependencias
  if (mainWindow) mainWindow.loadURL(loadingPage(
    'Instalando dependencias',
    'Descargando Flask, Pandas, OpenPyXL, pywin32...<br>Solo ocurre la primera vez.'
  ));
  try {
    await runPythonCmd(venvPython, ['-m', 'pip', 'install', '--upgrade', 'pip', '--quiet']);
    await runPythonCmd(venvPython, ['-m', 'pip', 'install', '-r', requirements, '--quiet']);
  } catch (e) {
    if (mainWindow) mainWindow.loadURL(loadingPage('Error al instalar dependencias', e.message, true));
    return;
  }

  if (mainWindow) mainWindow.loadURL(loadingPage('Listo. Iniciando...', 'Abriendo el Sistema de Vacaciones USIL...'));
  startPythonServer(venvPython, serverScript);
}

// ─── Arrancar Flask ───────────────────────────────────────────────────────────
function startPythonServer(pythonExe, serverScript) {
  const baseDir = app.isPackaged ? path.dirname(app.getPath('exe')) : __dirname;
  pythonProcess = spawn(pythonExe, [serverScript], { cwd: baseDir });
  pythonProcess.stdout.on('data', d => console.log(`[Python] ${d}`));
  pythonProcess.stderr.on('data', d => console.error(`[Python Err] ${d}`));
  pythonProcess.on('error', err => {
    if (mainWindow) mainWindow.loadURL(loadingPage('Error iniciando servidor', err.message, true));
  });
  setTimeout(checkServerReady, 2000);
}

// ─── Ping hasta que Flask responda ───────────────────────────────────────────
function checkServerReady() {
  const req = http.request(
    { hostname: '127.0.0.1', port: 5002, path: '/', method: 'GET' },
    res => {
      if ([200, 302, 304, 404].includes(res.statusCode)) {
        if (mainWindow) mainWindow.loadURL('http://127.0.0.1:5002');
      } else {
        setTimeout(checkServerReady, 1000);
      }
    }
  );
  req.on('error', () => setTimeout(checkServerReady, 1000));
  req.end();
}

// ─── Ciclo de vida ────────────────────────────────────────────────────────────
app.on('ready', () => {
  createWindow();
  setupAndStartServer().catch(err => {
    console.error('[BOOT] Error inesperado:', err);
    if (mainWindow) mainWindow.loadURL(loadingPage('Error inesperado al iniciar', err.message, true));
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('quit', () => {
  if (pythonProcess && !pythonProcess.killed) {
    try {
      spawn('taskkill', ['/pid', String(pythonProcess.pid), '/f', '/t'],
        { detached: true, stdio: 'ignore' }).unref();
    } catch (_) {
      try { pythonProcess.kill('SIGKILL'); } catch (_) {}
    }
  }
});
