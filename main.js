const { app, BrowserWindow } = require('electron');
const path = require('path');
const { spawn, execFile } = require('child_process');
const http = require('http');
const fs = require('fs');

let mainWindow;
let pythonProcess;

// ─── Pantalla de carga con estado ────────────────────────────────────────────
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

// ─── Crear ventana principal ──────────────────────────────────────────────────
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    icon: path.join(__dirname, 'icon.png'),
    autoHideMenuBar: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true
    }
  });

  mainWindow.maximize();
  mainWindow.loadURL(loadingPage('Iniciando Sistema de Vacaciones USIL...', 'Preparando el motor de datos...'));

  mainWindow.on('closed', function () {
    mainWindow = null;
  });
}

// ─── Ejecutar un comando Python y esperar a que termine ──────────────────────
function runPythonCmd(pythonExe, args, cwd) {
  return new Promise((resolve, reject) => {
    const currentDir = cwd || (app.isPackaged ? path.dirname(app.getPath('exe')) : __dirname);
    const proc = spawn(pythonExe, args, { cwd: currentDir, shell: false });
    let out = '';
    let err = '';
    proc.stdout.on('data', d => { out += d.toString(); });
    proc.stderr.on('data', d => { err += d.toString(); });
    proc.on('close', code => {
      if (code === 0) resolve(out);
      else reject(new Error(err || `Exit code ${code}`));
    });
    proc.on('error', reject);
  });
}

// ─── Buscar Python instalado en el sistema ───────────────────────────────────
async function detectarPython() {
  const candidatos = ['python', 'python3', 'py'];
  for (const py of candidatos) {
    try {
      await runPythonCmd(py, ['--version']);
      return py;
    } catch (e) {
      // siguiente candidato
    }
  }
  return null;
}

// ─── Setup + arranque del servidor ───────────────────────────────────────────
async function setupAndStartServer() {
  const baseDir = app.isPackaged ? path.dirname(app.getPath('exe')) : __dirname;
  const venvDir = app.isPackaged ? path.join(app.getPath('userData'), '.venv') : path.join(baseDir, '.venv');
  const venvPython  = path.join(venvDir, 'Scripts', 'python.exe');
  const serverScript = path.join(baseDir, 'servidor.py');
  const requirements = path.join(baseDir, 'requirements.txt');

  // 1. ¿Ya tiene .venv? Arrancar directo.
  if (fs.existsSync(venvPython)) {
    console.log('[BOOT] .venv encontrado → arrancando servidor...');
    startPythonServer(venvPython, serverScript);
    return;
  }

  // 2. Buscar Python global.
  if (mainWindow) mainWindow.loadURL(loadingPage(
    'Primera vez en este equipo',
    'Buscando Python instalado en el sistema...'
  ));

  const pythonGlobal = await detectarPython();

  if (!pythonGlobal) {
    if (mainWindow) mainWindow.loadURL(loadingPage(
      'Python no encontrado',
      'Este sistema requiere Python 3.10+ instalado.<br>' +
      'Descárgalo desde <b>python.org</b>, instálalo marcando "Add to PATH" y vuelve a abrir la app.',
      true
    ));
    return;
  }

  // 3. Verificar si el Python global ya tiene las dependencias requeridas instaladas
  if (mainWindow) mainWindow.loadURL(loadingPage(
    'Verificando tecnologías',
    'Verificando si ya tienes las dependencias de Python necesarias instaladas...'
  ));

  try {
    console.log('[BOOT] Probando si Python global ya tiene las librerías necesarias...');
    await runPythonCmd(pythonGlobal, ['-c', 'import flask, pandas, openpyxl']);
    console.log('[BOOT] ¡Librerías encontradas en Python global! Arrancando directamente...');
    
    if (mainWindow) mainWindow.loadURL(loadingPage(
      'Listo. Iniciando el sistema...',
      'Abriendo el Sistema de Vacaciones USIL...'
    ));

    startPythonServer(pythonGlobal, serverScript);
    return;
  } catch (e) {
    console.log('[BOOT] Falta alguna dependencia en Python global. Se procederá a crear un entorno virtual local...');
  }

  // 4. Crear entorno virtual.
  if (mainWindow) mainWindow.loadURL(loadingPage(
    'Configurando entorno (primera vez)',
    'Creando entorno virtual Python...<br>Esto puede tardar 1-2 minutos la primera vez.'
  ));

  try {
    await runPythonCmd(pythonGlobal, ['-m', 'venv', venvDir]);
    console.log('[BOOT] .venv creado OK en: ' + venvDir);
  } catch (e) {
    if (mainWindow) mainWindow.loadURL(loadingPage(
      'Error al crear entorno',
      `No se pudo crear el entorno virtual: ${e.message}`,
      true
    ));
    return;
  }

  // 5. Instalar dependencias.
  if (mainWindow) mainWindow.loadURL(loadingPage(
    'Instalando dependencias',
    'Descargando e instalando Flask, Pandas, OpenPyXL...<br>Solo ocurre la primera vez en este equipo.'
  ));

  try {
    await runPythonCmd(venvPython, ['-m', 'pip', 'install', '--upgrade', 'pip', '--quiet']);
    await runPythonCmd(venvPython, ['-m', 'pip', 'install', '-r', requirements, '--quiet']);
    console.log('[BOOT] Dependencias instaladas OK');
  } catch (e) {
    if (mainWindow) mainWindow.loadURL(loadingPage(
      'Error al instalar dependencias',
      `Problema instalando librerías: ${e.message}`,
      true
    ));
    return;
  }

  // 6. Todo listo → arrancar servidor.
  if (mainWindow) mainWindow.loadURL(loadingPage(
    'Listo. Iniciando el sistema...',
    'Abriendo el Sistema de Vacaciones USIL...'
  ));

  startPythonServer(venvPython, serverScript);
}

// ─── Arrancar proceso Python (Flask) ─────────────────────────────────────────
function startPythonServer(pythonExe, serverScript) {
  const baseDir = app.isPackaged ? path.dirname(app.getPath('exe')) : __dirname;
  pythonProcess = spawn(pythonExe, [serverScript], {
    cwd: baseDir
  });

  pythonProcess.stdout.on('data', (data) => {
    console.log(`[Python] ${data}`);
  });

  pythonProcess.stderr.on('data', (data) => {
    console.error(`[Python Err] ${data}`);
  });

  pythonProcess.on('error', (err) => {
    console.error('[Python] No se pudo iniciar:', err);
    if (mainWindow) mainWindow.loadURL(loadingPage(
      'Error iniciando el servidor',
      `Python no pudo arrancar: ${err.message}`,
      true
    ));
  });

  // Hacer "ping" al servidor hasta que responda
  setTimeout(checkServerReady, 2000);
}

// ─── Esperar a que Flask responda ────────────────────────────────────────────
function checkServerReady() {
  const req = http.request(
    { hostname: '127.0.0.1', port: 5002, path: '/', method: 'GET' },
    (res) => {
      if ([200, 302, 404].includes(res.statusCode)) {
        if (mainWindow) mainWindow.loadURL('http://127.0.0.1:5002');
      } else {
        setTimeout(checkServerReady, 1000);
      }
    }
  );
  req.on('error', () => setTimeout(checkServerReady, 1000));
  req.end();
}

// ─── Ciclo de vida de la app ──────────────────────────────────────────────────
app.on('ready', () => {
  createWindow();
  setupAndStartServer().catch(err => {
    console.error('[BOOT] Error inesperado:', err);
    if (mainWindow) mainWindow.loadURL(loadingPage(
      'Error inesperado al iniciar',
      err.message,
      true
    ));
  });
});

app.on('window-all-closed', function () {
  // Al cerrar todas las ventanas → cerrar la app + matar Python
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('quit', () => {
  // Garantizar que el servidor Flask muere con la app
  if (pythonProcess && !pythonProcess.killed) {
    try {
      // En Windows: matar el proceso y todos sus hijos
      spawn('taskkill', ['/pid', String(pythonProcess.pid), '/f', '/t'], {
        detached: true, stdio: 'ignore'
      }).unref();
    } catch (e) {
      try { pythonProcess.kill('SIGKILL'); } catch (_) {}
    }
  }
});
