@echo off
title Diagnostico de Actualizacion - Sistema de Vacaciones USIL
color 0B
set "SVU_BAT_DIR=%~dp0"
set "SVU_TMP=%TEMP%\svu_diagnostico_%RANDOM%.ps1"

powershell -NoProfile -Command "$m='#---PS1-BEGIN---'; $s=[IO.File]::ReadAllText('%~f0'); $i=$s.LastIndexOf($m); $code=$s.Substring($i+$m.Length); [IO.File]::WriteAllText('%SVU_TMP%',$code)"

if not exist "%SVU_TMP%" (
    echo [ERROR] No se pudo preparar el script de diagnostico.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%SVU_TMP%"

del "%SVU_TMP%" >nul 2>&1
echo.
pause
exit /b 0

#---PS1-BEGIN---
$ErrorActionPreference = 'SilentlyContinue'
$Repo   = 'AtraccionDeTalento/Sistema-Vacaciones'
$Branch = 'main'

Write-Host "============================================================"
Write-Host "  DIAGNOSTICO DE ACTUALIZACION AUTOMATICA"
Write-Host "  Sistema de Vacaciones USIL"
Write-Host "============================================================"
Write-Host "  Ejecuta este archivo en la PC que NO se esta actualizando."
Write-Host "  Al final, copia toda esta pantalla y enviala a soporte."
Write-Host "============================================================"
Write-Host ""

# ------------------------------------------------------------------
# [1/7] Ubicar la instalacion (rutas conocidas, luego preguntar)
# ------------------------------------------------------------------
Write-Host "[1/7] Buscando la instalacion..."

$exeName = 'Sistema Vacaciones USIL.exe'
$candidatos = @(
    (Join-Path $env:LOCALAPPDATA 'Programs\sistema-vacaciones-usil'),
    (Join-Path $env:SVU_BAT_DIR 'dist_electron\win-unpacked')
)

$exeDir = $null
foreach ($c in $candidatos) {
    if ($c -and (Test-Path (Join-Path $c $exeName))) { $exeDir = $c; break }
}

if ($exeDir) {
    Write-Host "    Encontrado en ruta conocida: $exeDir"
} else {
    Write-Host "    No se encontro en las rutas conocidas."
    Write-Host ""
    Write-Host "    Truco: abre el icono de 'Sistema de Vacaciones USIL' en el Escritorio o Menu"
    Write-Host "    Inicio, clic derecho -> Abrir ubicacion del archivo. Esa carpeta es la que"
    Write-Host "    necesitas pegar abajo (arrastrala a esta ventana para pegar la ruta)."
    Write-Host ""
    $manual = Read-Host "    Pega aqui la ruta de la CARPETA de instalacion (donde esta el .exe) y presiona Enter"
    if ($manual) {
        $manual = $manual.Trim().Trim('"').TrimEnd('\')
        if (Test-Path $manual) { $exeDir = $manual }
    }
}

if (-not $exeDir) {
    Write-Host ""
    Write-Host "[ERROR] No se pudo determinar la carpeta de instalacion. Fin del diagnostico."
    Write-Host ""
    Read-Host "Presiona Enter para cerrar"
    exit 1
}

Write-Host "    Carpeta de instalacion: $exeDir"
if ($exeDir -match 'OneDrive') {
    Write-Host "    [AVISO] Esta instalado dentro de una carpeta OneDrive."
    Write-Host "            Ademas de este diagnostico, revisa el icono de OneDrive en la"
    Write-Host "            barra de tareas de esta PC (que diga 'Al dia', con sesion"
    Write-Host "            iniciada y sin sincronizacion pausada)."
}
Write-Host ""

# ------------------------------------------------------------------
# [2/7] Version instalada localmente
# ------------------------------------------------------------------
Write-Host "[2/7] Version instalada en esta PC..."
$verFile = Join-Path $exeDir '.version_commit'
$localVer = $null
if (Test-Path $verFile) {
    $localVer = (Get-Content $verFile -Raw).Trim()
    Write-Host "    Commit local: $localVer"
} else {
    Write-Host "    [AVISO] No existe .version_commit -- esta instalacion nunca completo una actualizacion."
}
Write-Host ""

# ------------------------------------------------------------------
# [3/7] servidor.py presente (senal de instalacion completa)
# ------------------------------------------------------------------
Write-Host "[3/7] Archivo servidor.py en esta instalacion..."
$srv = Join-Path $exeDir 'servidor.py'
if (Test-Path $srv) {
    Write-Host "    servidor.py encontrado. Modificado: $((Get-Item $srv).LastWriteTime)"
} else {
    Write-Host "    [ERROR] No se encontro servidor.py en $exeDir -- instalacion incompleta o corrupta."
}
Write-Host ""

# ------------------------------------------------------------------
# [4/7] Conectividad a GitHub + deteccion de proxy
# ------------------------------------------------------------------
Write-Host "[4/7] Probando conexion a GitHub..."
foreach ($t in 'https://github.com', 'https://api.github.com', 'https://raw.githubusercontent.com') {
    try {
        $r = Invoke-WebRequest -Uri $t -Method Head -TimeoutSec 10 -UseBasicParsing -ErrorAction Stop
        Write-Host ("    {0,-40} HTTP {1}" -f $t, [int]$r.StatusCode)
    } catch {
        $resp = $_.Exception.Response
        if ($resp) {
            Write-Host ("    {0,-40} HTTP {1}" -f $t, [int]$resp.StatusCode)
        } else {
            Write-Host ("    {0,-40} SIN RESPUESTA: {1}" -f $t, $_.Exception.Message)
        }
    }
}
Write-Host ""
Write-Host "    Proxy configurado en Windows (WinHTTP):"
netsh winhttp show proxy | ForEach-Object { Write-Host "    $_" }
if ($env:HTTP_PROXY -or $env:HTTPS_PROXY) {
    Write-Host "    Variables de entorno detectadas: HTTP_PROXY=$($env:HTTP_PROXY) HTTPS_PROXY=$($env:HTTPS_PROXY)"
}
Write-Host "    [NOTA] La app se conecta con Node.js directo y NO usa el proxy de Windows"
Write-Host "           ni variables HTTP_PROXY/HTTPS_PROXY. Si esta PC necesita proxy para"
Write-Host "           salir a internet, la auto-actualizacion falla aunque el navegador funcione."
Write-Host ""

# ------------------------------------------------------------------
# [5/7] Ultimo commit en GitHub y comparacion contra la version local
# ------------------------------------------------------------------
Write-Host "[5/7] Consultando ultimo commit en GitHub (rama $Branch)..."
$remoteVer = $null
try {
    $resp2 = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/commits/$Branch" -Headers @{ 'User-Agent' = 'diag-sistema-vacaciones' } -TimeoutSec 15 -UseBasicParsing -ErrorAction Stop
    $remoteVer = $resp2.sha
    Write-Host "    Commit remoto (GitHub): $remoteVer"
} catch {
    Write-Host "    [ERROR] No se pudo consultar GitHub: $($_.Exception.Message)"
}
Write-Host ""
if (-not $remoteVer) {
    Write-Host "    No se pudo comparar: fallo la conexion a GitHub (ver el paso 4)."
} elseif (-not $localVer) {
    Write-Host "    [AVISO] No se puede comparar porque esta PC nunca completo una actualizacion."
} elseif ($localVer -eq $remoteVer) {
    Write-Host "    [OK] Esta PC esta al dia con GitHub."
} else {
    Write-Host "    [DESACTUALIZADO]"
    Write-Host "    Local:  $localVer"
    Write-Host "    Remoto: $remoteVer"
}
Write-Host ""

# ------------------------------------------------------------------
# [6/7] Comparacion ARCHIVO POR ARCHIVO contra GitHub (no solo el resumen
# global de .version_commit, que puede decir "al dia" aunque algunos archivos
# individuales hayan fallado al descargarse en una actualizacion parcial).
# ------------------------------------------------------------------
Write-Host "[6/7] Comparando cada archivo local contra la version en GitHub..."
Write-Host "      (esto es lo mas preciso: dice archivo por archivo si esta"
Write-Host "       al dia, desactualizado, o si falta por completo)"
Write-Host ""

# Misma lista que ARCHIVOS_ACTUALIZABLES en main.js -- si se agrega un archivo
# nuevo a esa lista, agregarlo tambien aqui.
$archivos = @(
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
    'PIPELINE/bot_adryan/guardar_password.py'
)

function Get-Sha256OfBytes($bytes) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { -join ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) }
    finally { $sha.Dispose() }
}

$okCount = 0; $desactCount = 0; $faltaCount = 0; $errCount = 0
foreach ($rel in $archivos) {
    $relDisplay = $rel
    $local = Join-Path $exeDir ($rel -replace '/', '\')
    if (-not (Test-Path $local)) {
        Write-Host ("    [FALTA]          {0}" -f $relDisplay)
        $faltaCount++
        continue
    }
    try {
        $localBytes = [IO.File]::ReadAllBytes($local)
        $localHash = Get-Sha256OfBytes $localBytes
        $url = "https://raw.githubusercontent.com/$Repo/$Branch/$rel"
        $remoteBytes = $null
        for ($intento = 1; $intento -le 2; $intento++) {
            try {
                $remoteBytes = (Invoke-WebRequest -Uri $url -TimeoutSec 30 -UseBasicParsing -ErrorAction Stop).Content
                break
            } catch {
                if ($intento -eq 2) { throw }
            }
        }
        if ($remoteBytes -is [string]) { $remoteBytes = [Text.Encoding]::UTF8.GetBytes($remoteBytes) }
        $remoteHash = Get-Sha256OfBytes $remoteBytes
        if ($localHash -eq $remoteHash) {
            Write-Host ("    [OK]             {0}" -f $relDisplay)
            $okCount++
        } else {
            $localMod = (Get-Item $local).LastWriteTime
            Write-Host ("    [DESACTUALIZADO] {0}  (modificado local: {1})" -f $relDisplay, $localMod)
            $desactCount++
        }
    } catch {
        Write-Host ("    [ERROR AL COMPARAR] {0} -- {1}" -f $relDisplay, $_.Exception.Message)
        $errCount++
    }
}
Write-Host ""
Write-Host ("    Resumen: {0} al dia, {1} desactualizados, {2} faltantes, {3} no se pudieron comparar" -f $okCount, $desactCount, $faltaCount, $errCount)
if ($desactCount -gt 0 -or $faltaCount -gt 0) {
    Write-Host "    [DIAGNOSTICO] Esta PC tiene una actualizacion INCOMPLETA: unos archivos se"
    Write-Host "    bajaron y otros no. Suele pasar cuando la conexion se corta a mitad de la"
    Write-Host "    descarga. Cierra la app por completo (Administrador de tareas si hace falta)"
    Write-Host "    y abrela de nuevo con buena conexion; el auto-actualizador ahora reintenta"
    Write-Host "    el set completo si detecta que algo quedo a medias."
}
Write-Host ""

# ------------------------------------------------------------------
# [7/7] Log propio de la app (update.log) - dice exactamente que paso
# ------------------------------------------------------------------
Write-Host "[7/7] Ultimas lineas de update.log (registro propio de la app)..."
$logFile = Join-Path $exeDir 'update.log'
if (Test-Path $logFile) {
    Get-Content $logFile -Tail 30 | ForEach-Object { Write-Host "    $_" }
} else {
    Write-Host "    No existe update.log -- la app nunca ejecuto el chequeo de actualizacion en esta carpeta."
    Write-Host "    (verifica que el acceso directo abra este .exe y no una copia vieja en otra ruta)."
}
Write-Host ""

# ------------------------------------------------------------------
# INFO EXTRA: procesos de la app corriendo ahora mismo
# ------------------------------------------------------------------
Write-Host "INFO EXTRA: procesos 'Sistema Vacaciones USIL' activos..."
$procs = Get-Process -Name 'Sistema Vacaciones USIL' -ErrorAction SilentlyContinue
if ($procs) {
    $procs | ForEach-Object { Write-Host "    PID $($_.Id) -> $($_.Path)" }
    Write-Host "    Si hay una ruta distinta a la de arriba, hay MAS de una instalacion en esta PC."
} else {
    Write-Host "    No hay ninguna instancia abierta ahora mismo."
}

Write-Host ""
Write-Host "============================================================"
Write-Host "  DIAGNOSTICO TERMINADO"
Write-Host "  - Si el paso 4 muestra 'SIN RESPUESTA' en github.com, hay un"
Write-Host "    firewall, antivirus o proxy bloqueando la conexion."
Write-Host "  - Si el paso 5 dice DESACTUALIZADO, cierra la app por completo"
Write-Host "    (Administrador de tareas si hace falta) y vuelve a abrirla:"
Write-Host "    se actualiza sola al iniciar, si hay internet."
Write-Host "  - Copia toda esta pantalla y enviala a soporte si necesitas ayuda."
Write-Host "============================================================"
Write-Host ""
