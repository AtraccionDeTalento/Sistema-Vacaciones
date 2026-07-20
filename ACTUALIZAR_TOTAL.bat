@echo off
title Sistema de Vacaciones USIL - Actualizar Total (codigo + limpieza)
color 0A
set "SVU_BAT_DIR=%~dp0"
set "SVU_TMP=%TEMP%\svu_actualizartotal_%RANDOM%.ps1"

powershell -NoProfile -Command "$m='#---PS1-BEGIN---'; $s=[IO.File]::ReadAllText('%~f0'); $i=$s.LastIndexOf($m); $code=$s.Substring($i+$m.Length); [IO.File]::WriteAllText('%SVU_TMP%',$code)"

if not exist "%SVU_TMP%" (
    echo [ERROR] No se pudo preparar el script.
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
$exeName = 'Sistema Vacaciones USIL.exe'
$BatDir = $env:SVU_BAT_DIR
if (-not $BatDir) { $BatDir = (Get-Location).Path }
$BatDir = $BatDir.TrimEnd('\')

Write-Host "============================================================"
Write-Host "  ACTUALIZAR TOTAL -- Sistema de Vacaciones USIL"
Write-Host "============================================================"
Write-Host "  Un solo clic para dejar esta instalacion IGUAL a la version"
Write-Host "  mas reciente, SIN reinstalar el programa:"
Write-Host "    1. Mata procesos viejos (el codigo en memoria no se"
Write-Host "       actualiza solo)."
Write-Host "    2. Descarga y verifica (SHA256) todo el codigo desde"
Write-Host "       GitHub -- igual que REPARAR_TOTAL.bat."
Write-Host "    3. LIMPIA duplicados y basura real: copias viejas de"
Write-Host "       Maestro de Personal, backups sueltos del Excel de"
Write-Host "       vacaciones (_previo_, _BACKUP_, _viejo_), cache"
Write-Host "       (.pkl) desactualizado. NO borra el Excel que la app"
Write-Host "       esta usando ahora mismo -- la app sigue funcionando"
Write-Host "       de inmediato, sin tener que volver a loguearse en"
Write-Host "       Adryan ni esperar una descarga nueva."
Write-Host "    4. Reabre la app y confirma con hash que quedo"
Write-Host "       corriendo el codigo nuevo de verdad."
Write-Host "============================================================"
Write-Host ""

$resultado = [ordered]@{
    timestamp_utc = (Get-Date).ToUniversalTime().ToString('o')
    maquina       = $env:COMPUTERNAME
    usuario       = $env:USERNAME
    pasos         = [ordered]@{}
}

# ------------------------------------------------------------------
# [1/6] Ubicar la instalacion
# ------------------------------------------------------------------
Write-Host "[1/6] Buscando la instalacion en esta PC..."
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
# [2/6] Matar procesos viejos
# ------------------------------------------------------------------
Write-Host "[2/6] Cerrando procesos viejos..."
$matados = 0
Get-Process -Name 'Sistema Vacaciones USIL' -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue; $matados++ }
Get-CimInstance Win32_Process -Filter "Name='python.exe' OR Name='pythonw.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -match 'servidor\.py' -or $_.CommandLine -match [regex]::Escape($Aqui) } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue; $matados++ }
if ($matados -gt 0) { Start-Sleep -Seconds 2 }
Write-Host "    $matados proceso(s) cerrado(s)."
Write-Host ""

# ------------------------------------------------------------------
# [3/6] Descargar y verificar codigo desde GitHub
# ------------------------------------------------------------------
$archivos = @(
    'servidor.py', 'index_vacaciones.html', 'enviar_cola_outlook.py', '_bp_map.py',
    'requirements.txt', 'assets/js/app_completo.js', 'assets/js/pipeline_vac.js',
    'assets/css/styles.css', 'PIPELINE/motor/pipeline.py', 'PIPELINE/motor/vac_lib.py',
    'PIPELINE/motor/config.json', 'PIPELINE/motor/requirements.txt',
    'PIPELINE/bot_adryan/bot_adryan.py', 'PIPELINE/bot_adryan/bot_maestro.py',
    'PIPELINE/bot_adryan/guardar_password.py'
)
function Get-Sha256OfBytes($bytes) {
    $texto = [Text.Encoding]::UTF8.GetString($bytes) -replace "`r`n", "`n"
    $bytesNorm = [Text.Encoding]::UTF8.GetBytes($texto)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { -join ($sha.ComputeHash($bytesNorm) | ForEach-Object { $_.ToString('x2') }) }
    finally { $sha.Dispose() }
}

Write-Host "[3/6] Descargando y verificando codigo desde GitHub..."
$remoteVer = $null
try {
    $resp = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/commits/$Branch" -Headers @{ 'User-Agent' = 'actualizar-total' } -TimeoutSec 15 -UseBasicParsing -ErrorAction Stop
    $remoteVer = $resp.sha
    Write-Host "    Commit remoto: $remoteVer"
} catch {
    Write-Host "    [ERROR] Sin conexion a GitHub -- no se puede continuar. Revisa internet."
    Read-Host "Presiona Enter para cerrar"
    exit 1
}
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
        $verif = [IO.File]::ReadAllBytes($destino)
        if ((Get-Sha256OfBytes $verif) -eq (Get-Sha256OfBytes $contenidoBytes)) { $ok++ } else { $fallidos += $rel }
    } else { $fallidos += $rel }
}
Write-Host "    $ok/$($archivos.Count) archivos al dia y verificados."
if ($fallidos.Count -eq 0) {
    Set-Content -Path (Join-Path $Aqui '.version_commit') -Value $remoteVer -Encoding UTF8 -NoNewline
} else {
    Write-Host "    [AVISO] Fallaron: $($fallidos -join ', ')"
}
$resultado.pasos['3_codigo'] = @{ ok = $ok; total = $archivos.Count; fallidos = $fallidos; commit = $remoteVer }
Write-Host ""

# ------------------------------------------------------------------
# [4/6] Limpiar duplicados y basura real -- SIN tocar el Excel que la app
# esta usando ahora mismo. Solo mueve a respaldo lo que es inequivocamente
# una copia vieja: backups del pipeline (_previo_, _BACKUP_, _viejo_) y
# copias antiguas del Maestro de Personal (se queda solo con la mas
# reciente por fecha en el nombre). Asi la app sigue funcionando de
# inmediato despues de este script, sin depender de volver a loguearse en
# Adryan ni de esperar una descarga nueva.
# ------------------------------------------------------------------
Write-Host "[4/6] Limpiando duplicados y respaldos sueltos..."
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$respaldo = Join-Path $Aqui "_LIMPIEZA_AUTOMATICA_$stamp"
$movidos = 0

$dsDir = Join-Path $Aqui 'DATA SENSIBLE'
if (Test-Path $dsDir) {
    $junkPatterns = @('*_previo_*.xlsx', '*__previo_*.xlsx', '*_BACKUP_*.xlsx', '*_viejo_*.xlsx', '*__viejo_*.xlsx')
    foreach ($pat in $junkPatterns) {
        Get-ChildItem -Path $dsDir -Filter $pat -File -ErrorAction SilentlyContinue | ForEach-Object {
            $dest = Join-Path $respaldo 'DATA SENSIBLE'
            New-Item -ItemType Directory -Path $dest -Force | Out-Null
            try {
                Move-Item -LiteralPath $_.FullName -Destination (Join-Path $dest $_.Name) -Force -ErrorAction Stop
                Write-Host "    [OK] Movido a respaldo: DATA SENSIBLE\$($_.Name)"
                $movidos++
            } catch { Write-Host "    [AVISO] No se pudo mover $($_.Name) (puede estar en uso)." }
        }
    }
}

$datasDir = Join-Path $Aqui 'DATAS'
if (Test-Path $datasDir) {
    $maestros = Get-ChildItem -Path $datasDir -Filter 'PersonalMaestroReporte*.xlsx' -File -ErrorAction SilentlyContinue |
        Sort-Object Name -Descending
    if ($maestros -and $maestros.Count -gt 1) {
        Write-Host "    Maestro de Personal: $($maestros.Count) copias encontradas, se conserva la mas reciente ($($maestros[0].Name))."
        for ($i = 1; $i -lt $maestros.Count; $i++) {
            $dest = Join-Path $respaldo 'DATAS'
            New-Item -ItemType Directory -Path $dest -Force | Out-Null
            try {
                Move-Item -LiteralPath $maestros[$i].FullName -Destination (Join-Path $dest $maestros[$i].Name) -Force -ErrorAction Stop
                Write-Host "    [OK] Movido a respaldo: DATAS\$($maestros[$i].Name)"
                $movidos++
            } catch { Write-Host "    [AVISO] No se pudo mover $($maestros[$i].Name) (puede estar en uso)." }
        }
    }
    # Cache: se borra siempre (barato de reconstruir, y evita leer un calculo
    # viejo si el Excel cambio pero el mtime no se detecto bien).
    $cacheDir = Join-Path $datasDir '__cache__'
    if (Test-Path $cacheDir) {
        Get-ChildItem -Path $cacheDir -Filter '*.pkl' -ErrorAction SilentlyContinue | ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop; Write-Host "    [OK] Cache limpiado: $($_.Name)" }
            catch {}
        }
    }
}

if ($movidos -eq 0) {
    Write-Host "    Nada que limpiar -- ya estaba ordenado."
} else {
    Write-Host "    $movidos archivo(s) movidos a respaldo (reversible): $respaldo"
}
$resultado.pasos['4_limpieza'] = @{ movidos = $movidos; respaldo = if ($movidos -gt 0) { $respaldo } else { $null } }
Write-Host ""

# ------------------------------------------------------------------
# [5/6] Reabrir la app
# ------------------------------------------------------------------
Write-Host "[5/6] Reabriendo la app..."
$exePath = Join-Path $Aqui $exeName
if (Test-Path $exePath) {
    Start-Process $exePath
    Write-Host "    Abierta: $exePath"
} else {
    Write-Host "    [ERROR] No se encontro $exeName en $Aqui"
}
Write-Host ""

# ------------------------------------------------------------------
# [6/6] Verificar que el proceso levantado usa el codigo nuevo (hash en
# memoria vs GitHub -- no basta con que el archivo en disco haya cambiado).
# ------------------------------------------------------------------
Write-Host "[6/6] Verificando que el proceso levantado usa el codigo nuevo..."
$health = $null
for ($i = 1; $i -le 20; $i++) {
    try { $health = Invoke-RestMethod -Uri "http://127.0.0.1:5002/api/health" -TimeoutSec 5 -ErrorAction Stop; if ($health) { break } } catch {}
    Start-Sleep -Seconds 3
}
$exito = $false
if ($health -and $remoteVer) {
    try {
        $diag = Invoke-RestMethod -Uri "http://127.0.0.1:5002/api/diagnostico/kpis" -TimeoutSec 90 -ErrorAction Stop
        $hashMem = $diag.diagnostico.version_codigo.servidor_py_sha256
        $remoteServidorBytes = (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/$Repo/$Branch/servidor.py" -TimeoutSec 30 -UseBasicParsing -ErrorAction Stop).Content
        if ($remoteServidorBytes -is [string]) { $remoteServidorBytes = [Text.Encoding]::UTF8.GetBytes($remoteServidorBytes) }
        $hashRemoto = Get-Sha256OfBytes $remoteServidorBytes
        if ($hashMem -and $hashMem -eq $hashRemoto) {
            Write-Host "    [OK] CONFIRMADO: el proceso corriendo usa el codigo mas reciente."
            $exito = $true
        } else {
            Write-Host "    [AVISO] No se pudo confirmar por hash -- revisa la app manualmente."
        }
    } catch {
        Write-Host "    (No se pudo verificar por hash -- normal si es el primer arranque y aun esta precargando, puede tardar varios minutos.)"
    }
} else {
    Write-Host "    (El servidor aun no respondio -- normal en el primer arranque tras limpiar cache, puede tardar mas de lo usual.)"
}
Write-Host ""
Write-Host "============================================================"
if ($exito) {
    Write-Host "  LISTO: instalacion actualizada y verificada."
} else {
    Write-Host "  Instalacion actualizada. No se pudo confirmar el reinicio por hash --"
    Write-Host "  si el dashboard se ve raro, cierra la app por completo y abrela de nuevo."
}
if ($resultado.pasos['4_limpieza'].respaldo) {
    Write-Host "  Respaldo de la limpieza (borralo cuando confirmes que todo esta bien):"
    Write-Host "    $($resultado.pasos['4_limpieza'].respaldo)"
}
Write-Host "============================================================"

$carpetaSalida = if ($env:SVU_BAT_DIR) { $env:SVU_BAT_DIR } else { $env:TEMP }
try {
    ($resultado | ConvertTo-Json -Depth 10) | Set-Content -Path (Join-Path $carpetaSalida 'ACTUALIZAR_TOTAL_RESULTADO.json') -Encoding UTF8
} catch {}
