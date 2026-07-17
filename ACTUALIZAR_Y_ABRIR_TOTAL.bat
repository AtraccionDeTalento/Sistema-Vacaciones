@echo off
title Sistema de Vacaciones USIL - Actualizar y Abrir
color 0B
set "SVU_BAT_DIR=%~dp0"
set "SVU_TMP=%TEMP%\svu_todo_%RANDOM%.ps1"

powershell -NoProfile -Command "$m='#---PS1-BEGIN---'; $s=[IO.File]::ReadAllText('%~f0'); $i=$s.LastIndexOf($m); $code=$s.Substring($i+$m.Length); [IO.File]::WriteAllText('%SVU_TMP%',$code)"

if not exist "%SVU_TMP%" (
    echo [ERROR] No se pudo preparar el script.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%SVU_TMP%"

del "%SVU_TMP%" >nul 2>&1
exit /b 0

#---PS1-BEGIN---
$ErrorActionPreference = 'SilentlyContinue'
$Repo   = 'AtraccionDeTalento/Sistema-Vacaciones'
$Branch = 'main'
$exeName = 'Sistema Vacaciones USIL.exe'
$BatDir = $env:SVU_BAT_DIR
if (-not $BatDir) { $BatDir = (Get-Location).Path }
$BatDir = $BatDir.TrimEnd('\')

Write-Host "============================================================"
Write-Host "  SISTEMA DE VACACIONES USIL -- Actualizar y Abrir"
Write-Host "============================================================"
Write-Host "  Este .bat es el UNICO punto de entrada recomendado: reemplaza"
Write-Host "  abrir 'Sistema Vacaciones USIL.exe' directo. Cada vez que se"
Write-Host "  usa, primero fuerza la actualizacion de codigo (sin depender"
Write-Host "  de git -- descarga directo de GitHub via HTTPS, funciona en"
Write-Host "  cualquier PC), la verifica con SHA256, y SOLO despues abre"
Write-Host "  la app. Asi nunca se vuelve a quedar corriendo codigo viejo"
Write-Host "  sin que nadie se de cuenta."
Write-Host "============================================================"
Write-Host ""

# ------------------------------------------------------------------
# [1/5] Ubicar la instalacion (misma logica auto-detect que el resto de la
# suite: LOCALAPPDATA, dist_electron/win-unpacked junto al .bat, la propia
# carpeta del .bat, acceso directo del Escritorio).
# ------------------------------------------------------------------
Write-Host "[1/5] Ubicando la instalacion..."
$candidatosRaw = New-Object System.Collections.Generic.List[object]
function Add-Candidato($ruta, $etiqueta) {
    if ($ruta) { $candidatosRaw.Add([pscustomobject]@{ ruta = $ruta.TrimEnd('\'); etiqueta = $etiqueta }) }
}
Add-Candidato (Join-Path $env:LOCALAPPDATA 'Programs\sistema-vacaciones-usil') 'Instalacion estandar (LOCALAPPDATA)'
Add-Candidato (Join-Path $BatDir 'dist_electron\win-unpacked') 'Build empaquetado (dist_electron/win-unpacked)'
Add-Candidato $BatDir 'Carpeta donde esta este .bat'
try {
    $shell = New-Object -ComObject WScript.Shell
    $lnkPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Sistema de Vacaciones USIL.lnk'
    if (Test-Path $lnkPath) {
        $target = $shell.CreateShortcut($lnkPath).TargetPath
        if ($target -and (Test-Path $target)) { Add-Candidato (Split-Path $target -Parent) 'Acceso directo del Escritorio' }
    }
} catch {}
$vistos = @{}
$candidatos = @()
foreach ($c in $candidatosRaw) {
    try { $abs = (Resolve-Path -LiteralPath $c.ruta -ErrorAction Stop).Path } catch { $abs = $c.ruta }
    $key = $abs.ToLowerInvariant()
    if (-not $vistos.ContainsKey($key) -and (Test-Path $abs)) {
        $vistos[$key] = $true
        $candidatos += [pscustomobject]@{ ruta = $abs; etiqueta = $c.etiqueta }
    }
}
$candidatos = @($candidatos | Where-Object { Test-Path (Join-Path $_.ruta $exeName) })
if (-not $candidatos -or $candidatos.Count -eq 0) {
    $manual = Read-Host "    No se encontro instalacion automaticamente. Pega la ruta de la carpeta (donde esta el .exe)"
    if ($manual) {
        $manual = $manual.Trim().Trim('"').TrimEnd('\')
        if (Test-Path $manual) { $candidatos = @([pscustomobject]@{ ruta = $manual; etiqueta = 'Ruta indicada manualmente' }) }
    }
}
if (-not $candidatos -or $candidatos.Count -eq 0) {
    Write-Host "[ERROR] No se encontro ninguna instalacion. Fin."
    Read-Host "Presiona Enter para cerrar"
    exit 1
}
$Aqui = $candidatos[0].ruta
Write-Host "    Instalacion: $Aqui"
Write-Host ""

# ------------------------------------------------------------------
# [2/5] Matar procesos viejos (codigo en memoria no se actualiza solo)
# ------------------------------------------------------------------
Write-Host "[2/5] Cerrando cualquier instancia vieja corriendo..."
Get-Process -Name 'Sistema Vacaciones USIL' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Get-CimInstance Win32_Process -Filter "Name='python.exe' OR Name='pythonw.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -match 'servidor\.py' -or $_.CommandLine -match [regex]::Escape($Aqui) } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
Start-Sleep -Seconds 1
Write-Host "    OK."
Write-Host ""

# ------------------------------------------------------------------
# [3/5] Forzar actualizacion de codigo via GitHub HTTPS (NO requiere git
# instalado ni un repo .git local -- a diferencia de ACTUALIZAR_Y_ABRIR.bat,
# que llamaba "git pull" y por eso no funcionaba en las copias distribuidas
# a otras PCs, donde no hay repo git). Verifica cada archivo con SHA256.
# ------------------------------------------------------------------
$archivos = @(
    'servidor.py', 'index_vacaciones.html', 'enviar_cola_outlook.py', '_bp_map.py',
    'requirements.txt', 'assets/js/app_completo.js', 'assets/js/pipeline_vac.js',
    'assets/css/styles.css', 'PIPELINE/motor/pipeline.py', 'PIPELINE/motor/vac_lib.py',
    'PIPELINE/motor/config.json', 'PIPELINE/bot_adryan/bot_adryan.py',
    'PIPELINE/bot_adryan/bot_maestro.py', 'PIPELINE/bot_adryan/guardar_password.py'
)
function Get-Sha256OfBytes($bytes) {
    $texto = [Text.Encoding]::UTF8.GetString($bytes) -replace "`r`n", "`n"
    $bytesNorm = [Text.Encoding]::UTF8.GetBytes($texto)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { -join ($sha.ComputeHash($bytesNorm) | ForEach-Object { $_.ToString('x2') }) }
    finally { $sha.Dispose() }
}

Write-Host "[3/5] Forzando actualizacion de codigo desde GitHub..."
$remoteVer = $null
try {
    $resp = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/commits/$Branch" -Headers @{ 'User-Agent' = 'actualizar-y-abrir-total' } -TimeoutSec 15 -UseBasicParsing -ErrorAction Stop
    $remoteVer = $resp.sha
    Write-Host "    Commit remoto: $remoteVer"
} catch {
    Write-Host "    [AVISO] Sin conexion a GitHub -- se abre la app con el codigo que ya tiene en disco."
}

if ($remoteVer) {
    $ok = 0; $fallidos = @()
    foreach ($rel in $archivos) {
        $destino = Join-Path $Aqui ($rel -replace '/', '\')
        $dir = Split-Path $destino -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $url = "https://raw.githubusercontent.com/$Repo/$Branch/$rel"
        $bajado = $false; $contenidoBytes = $null
        for ($intento = 1; $intento -le 3; $intento++) {
            try {
                $contenido = (Invoke-WebRequest -Uri $url -TimeoutSec 30 -UseBasicParsing -ErrorAction Stop).Content
                if ($contenido -is [string]) { $contenido = [Text.Encoding]::UTF8.GetBytes($contenido) }
                $contenidoBytes = $contenido
                $tmp = "$destino.tmp"
                [IO.File]::WriteAllBytes($tmp, $contenido)
                if (Test-Path $destino) { Remove-Item $destino -Force -ErrorAction Stop }
                Rename-Item $tmp $destino -Force
                $bajado = $true
                break
            } catch { Start-Sleep -Milliseconds 500 }
        }
        if ($bajado) {
            try {
                $verif = [IO.File]::ReadAllBytes($destino)
                if ((Get-Sha256OfBytes $verif) -eq (Get-Sha256OfBytes $contenidoBytes)) { $ok++ } else { $fallidos += $rel }
            } catch { $fallidos += $rel }
        } else { $fallidos += $rel }
    }
    Write-Host "    $ok/$($archivos.Count) archivos al dia y verificados."
    if ($fallidos.Count -eq 0) {
        Set-Content -Path (Join-Path $Aqui '.version_commit') -Value $remoteVer -Encoding UTF8 -NoNewline
        Write-Host "    [OK] Instalacion marcada al dia ($($remoteVer.Substring(0,7)))."
    } else {
        Write-Host "    [AVISO] Fallaron: $($fallidos -join ', ') -- se reintentara la proxima vez."
    }
}
Write-Host ""

# ------------------------------------------------------------------
# [4/5] Abrir la app
# ------------------------------------------------------------------
Write-Host "[4/5] Abriendo la app..."
$exePath = Join-Path $Aqui $exeName
if (Test-Path $exePath) {
    Start-Process $exePath
    Write-Host "    Abierta: $exePath"
} else {
    Write-Host "    [ERROR] No se encontro $exeName en $Aqui"
}
Write-Host ""

# ------------------------------------------------------------------
# [5/5] Confirmar que el proceso que quedo corriendo usa el codigo nuevo
# (mismo chequeo de REPARAR_TOTAL.bat: hash EN MEMORIA vs GitHub).
# ------------------------------------------------------------------
Write-Host "[5/5] Verificando que el proceso levantado usa el codigo nuevo..."
$health = $null
for ($i = 1; $i -le 20; $i++) {
    try { $health = Invoke-RestMethod -Uri "http://127.0.0.1:5002/api/health" -TimeoutSec 5 -ErrorAction Stop; if ($health) { break } } catch {}
    Start-Sleep -Seconds 3
}
if ($health -and $remoteVer) {
    try {
        $diag = Invoke-RestMethod -Uri "http://127.0.0.1:5002/api/diagnostico/kpis" -TimeoutSec 90 -ErrorAction Stop
        $hashMem = $diag.diagnostico.version_codigo.servidor_py_sha256
        $remoteServidorBytes = (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/$Repo/$Branch/servidor.py" -TimeoutSec 30 -UseBasicParsing -ErrorAction Stop).Content
        if ($remoteServidorBytes -is [string]) { $remoteServidorBytes = [Text.Encoding]::UTF8.GetBytes($remoteServidorBytes) }
        $hashRemoto = Get-Sha256OfBytes $remoteServidorBytes
        if ($hashMem -and $hashMem -eq $hashRemoto) {
            Write-Host "    [OK] CONFIRMADO: la app abierta esta corriendo el codigo mas reciente."
        } elseif ($hashMem) {
            Write-Host "    [AVISO] El proceso abierto NO coincide con GitHub -- puede haber otra instalacion corriendo."
        } else {
            Write-Host "    (Version anterior sin endpoint de diagnostico -- no se puede verificar por hash, pero los archivos ya se actualizaron.)"
        }
    } catch {
        Write-Host "    (No se pudo verificar por hash -- la app puede seguir cargando su primer arranque, que tarda varios minutos.)"
    }
} else {
    Write-Host "    (El servidor aun no respondio -- normal en el primer arranque, que puede tardar varios minutos en precargar los Excel.)"
}
Write-Host ""
Write-Host "============================================================"
Write-Host "  LISTO. Deja esta ventana abrirse y cerrar sola, o ciérrala"
Write-Host "  cuando veas la app abierta."
Write-Host "============================================================"
Start-Sleep -Seconds 3
