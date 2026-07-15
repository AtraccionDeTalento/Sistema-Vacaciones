@echo off
title Forzar Actualizacion - Sistema de Vacaciones USIL
color 0A
set "SVU_BAT_DIR=%~dp0"
set "SVU_TMP=%TEMP%\svu_forzar_%RANDOM%.ps1"

powershell -NoProfile -Command "$m='#---PS1-BEGIN---'; $s=[IO.File]::ReadAllText('%~f0'); $i=$s.LastIndexOf($m); $code=$s.Substring($i+$m.Length); [IO.File]::WriteAllText('%SVU_TMP%',$code)"

if not exist "%SVU_TMP%" (
    echo [ERROR] No se pudo preparar el script de actualizacion.
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

# Mismos archivos que ARCHIVOS_ACTUALIZABLES en main.js. Si se agrega uno
# nuevo alla, agregarlo tambien aqui (y en DIAGNOSTICO_ACTUALIZACION.bat).
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

$logPath = Join-Path $(if ($env:SVU_BAT_DIR) { $env:SVU_BAT_DIR } else { $env:TEMP }) 'FORZAR_ACTUALIZACION_LOG.txt'
try { Start-Transcript -Path $logPath -Force | Out-Null } catch {}

Write-Host "============================================================"
Write-Host "  FORZAR ACTUALIZACION"
Write-Host "  Sistema de Vacaciones USIL"
Write-Host "============================================================"
Write-Host "  Baja TODOS los archivos actualizables directo de GitHub y"
Write-Host "  los reemplaza en la instalacion, sin importar si el chequeo"
Write-Host "  normal de version dice que ya esta al dia. Usalo cuando"
Write-Host "  DIAGNOSTICO_ACTUALIZACION.bat marco alguna copia como"
Write-Host "  DESACTUALIZADA y quieres arreglarla ya mismo, sin esperar"
Write-Host "  a que la app lo intente sola en su proximo arranque."
Write-Host "============================================================"
Write-Host ""

# ------------------------------------------------------------------
# [1/5] Ubicar instalacion(es) empaquetadas -- misma logica de deteccion
# que DIAGNOSTICO_ACTUALIZACION.bat, pero aqui NO se incluye la carpeta
# fuente (con .git): forzar una descarga encima del repo de desarrollo no
# tiene sentido y podria pisar cambios sin commitear.
# ------------------------------------------------------------------
Write-Host "[1/5] Buscando instalaciones empaquetadas en esta PC..."

$candidatosRaw = New-Object System.Collections.Generic.List[object]
function Add-Candidato($ruta, $etiqueta) {
    if ($ruta) { $candidatosRaw.Add([pscustomobject]@{ ruta = $ruta.TrimEnd('\'); etiqueta = $etiqueta }) }
}
Add-Candidato (Join-Path $env:LOCALAPPDATA 'Programs\sistema-vacaciones-usil') 'Instalacion estandar (LOCALAPPDATA)'
Add-Candidato (Join-Path $env:SVU_BAT_DIR 'dist_electron\win-unpacked') 'Build empaquetado local (dist_electron/win-unpacked)'
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
    if (-not $vistos.ContainsKey($key) -and (Test-Path $abs) -and -not (Test-Path (Join-Path $abs '.git'))) {
        $vistos[$key] = $true
        $candidatos += [pscustomobject]@{ ruta = $abs; etiqueta = $c.etiqueta }
    }
}
$candidatos = @($candidatos | Where-Object { Test-Path (Join-Path $_.ruta $exeName) })

if (-not $candidatos -or $candidatos.Count -eq 0) {
    Write-Host "    No se encontro ninguna instalacion empaquetada en las rutas conocidas."
    Write-Host ""
    Write-Host "    Truco: abre el icono de 'Sistema de Vacaciones USIL' en el Escritorio o Menu"
    Write-Host "    Inicio, clic derecho -> Abrir ubicacion del archivo. Esa carpeta es la que"
    Write-Host "    necesitas pegar abajo (arrastrala a esta ventana para pegar la ruta)."
    Write-Host ""
    $manual = Read-Host "    Pega aqui la ruta de la CARPETA de instalacion (donde esta el .exe) y presiona Enter"
    if ($manual) {
        $manual = $manual.Trim().Trim('"').TrimEnd('\')
        if (Test-Path $manual) { $candidatos = @([pscustomobject]@{ ruta = $manual; etiqueta = 'Ruta indicada manualmente' }) }
    }
}

if (-not $candidatos -or $candidatos.Count -eq 0) {
    Write-Host ""
    Write-Host "[ERROR] No se encontro ninguna instalacion para actualizar. Fin."
    Write-Host ""
    Read-Host "Presiona Enter para cerrar"
    exit 1
}

$objetivo = $null
if ($candidatos.Count -eq 1) {
    $objetivo = $candidatos[0]
    Write-Host "    Se va a actualizar: $($objetivo.ruta)  ($($objetivo.etiqueta))"
} else {
    Write-Host "    Se encontraron $($candidatos.Count) instalaciones:"
    for ($i = 0; $i -lt $candidatos.Count; $i++) {
        Write-Host "      [$($i+1)] $($candidatos[$i].ruta)  ($($candidatos[$i].etiqueta))"
    }
    Write-Host ""
    $sel = Read-Host "    Escribe el numero de la que quieres forzar (o 'todas')"
    if ($sel -and $sel.Trim().ToLowerInvariant() -eq 'todas') {
        $objetivo = 'TODAS'
    } else {
        $idx = 0
        if ([int]::TryParse($sel, [ref]$idx) -and $idx -ge 1 -and $idx -le $candidatos.Count) {
            $objetivo = $candidatos[$idx - 1]
        }
    }
}
if (-not $objetivo) {
    Write-Host "[ERROR] Seleccion invalida. Fin."
    Read-Host "Presiona Enter para cerrar"
    exit 1
}
$objetivos = if ($objetivo -eq 'TODAS') { $candidatos } else { @($objetivo) }
Write-Host ""

# ------------------------------------------------------------------
# [2/5] Ultimo commit en GitHub (lo que se va a instalar)
# ------------------------------------------------------------------
Write-Host "[2/5] Consultando ultimo commit en GitHub (rama $Branch)..."
$remoteVer = $null
try {
    $resp = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/commits/$Branch" -Headers @{ 'User-Agent' = 'forzar-sistema-vacaciones' } -TimeoutSec 15 -UseBasicParsing -ErrorAction Stop
    $remoteVer = $resp.sha
    Write-Host "    Commit remoto: $remoteVer"
} catch {
    Write-Host "    [ERROR] No se pudo consultar GitHub: $($_.Exception.Message)"
}
Write-Host ""
if (-not $remoteVer) {
    Write-Host "[ERROR] Sin conexion a GitHub, no se puede continuar. Revisa tu internet e intenta de nuevo."
    Read-Host "Presiona Enter para cerrar"
    exit 1
}

# ------------------------------------------------------------------
# [3/5] Verificar (y ofrecer cerrar) procesos activos de esta instalacion,
# para no pisar archivos mientras estan en uso.
# ------------------------------------------------------------------
Write-Host "[3/5] Revisando si la app esta abierta ahora mismo..."
$procs = Get-Process -Name 'Sistema Vacaciones USIL' -ErrorAction SilentlyContinue
$procsPython = Get-Process -Name 'python', 'pythonw' -ErrorAction SilentlyContinue
if ($procs -or $procsPython) {
    Write-Host "    La app (o su servidor Python) esta corriendo ahora mismo."
    $cerrar = Read-Host "    ¿Cerrarla para poder actualizar sin problemas? (S/N)"
    if ($cerrar -and $cerrar.Trim().ToUpperInvariant() -eq 'S') {
        $procs | Stop-Process -Force -ErrorAction SilentlyContinue
        $procsPython | Where-Object { $_.MainWindowTitle -eq '' -or $_.Path -like '*Sistema*' } | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Write-Host "    Cerrada."
    } else {
        Write-Host "    Continuando sin cerrarla -- si algun archivo da error de acceso, ciérrala e intenta de nuevo."
    }
}
Write-Host ""

# ------------------------------------------------------------------
# [4/5] Descargar TODOS los archivos, siempre, sin comparar version
# ------------------------------------------------------------------
Write-Host "[4/5] Descargando archivos desde GitHub y reemplazando..."
Write-Host ""

foreach ($obj in $objetivos) {
    Write-Host "------------------------------------------------------------"
    Write-Host "  INSTALACION: $($obj.ruta)"
    Write-Host "------------------------------------------------------------"
    $ok = 0; $fallidos = @()
    foreach ($rel in $archivos) {
        $destino = Join-Path $obj.ruta ($rel -replace '/', '\')
        $dir = Split-Path $destino -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $url = "https://raw.githubusercontent.com/$Repo/$Branch/$rel"
        $bajado = $false
        for ($intento = 1; $intento -le 3; $intento++) {
            try {
                $contenido = (Invoke-WebRequest -Uri $url -TimeoutSec 30 -UseBasicParsing -ErrorAction Stop).Content
                # Invoke-WebRequest devuelve .Content como string para respuestas de
                # texto (la mayoria de estos archivos) y como byte[] para binarias;
                # WriteAllBytes exige byte[] siempre, así que hay que convertir.
                if ($contenido -is [string]) { $contenido = [Text.Encoding]::UTF8.GetBytes($contenido) }
                $tmp = "$destino.tmp"
                [IO.File]::WriteAllBytes($tmp, $contenido)
                if (Test-Path $destino) { Remove-Item $destino -Force -ErrorAction Stop }
                Rename-Item $tmp $destino -Force
                $bajado = $true
                break
            } catch {
                if ($intento -eq 3) {
                    $msg = "$($_.Exception.Message)"
                    if ($msg.Length -gt 200) { $msg = $msg.Substring(0, 200) + '...' }
                    Write-Host ("    [FALLO] {0} -- {1}" -f $rel, $msg)
                } else {
                    Start-Sleep -Milliseconds 500
                }
            }
        }
        if ($bajado) {
            Write-Host ("    [OK]    {0}" -f $rel)
            $ok++
        } else {
            $fallidos += $rel
        }
    }
    Write-Host ""
    Write-Host ("    Resultado: {0}/{1} archivos actualizados." -f $ok, $archivos.Count)
    if ($fallidos.Count -eq 0) {
        try {
            Set-Content -Path (Join-Path $obj.ruta '.version_commit') -Value $remoteVer -Encoding UTF8 -NoNewline
            Write-Host "    [OK] .version_commit actualizado a $remoteVer -- esta instalacion queda marcada al dia."
        } catch {
            Write-Host "    [AVISO] No se pudo escribir .version_commit: $($_.Exception.Message)"
        }
    } else {
        Write-Host "    [AVISO] Quedaron $($fallidos.Count) archivo(s) sin bajar: $($fallidos -join ', ')"
        Write-Host "    No se marca como completa -- vuelve a correr este mismo .bat para reintentar solo lo que falta."
    }
    Write-Host ""
}

# ------------------------------------------------------------------
# [5/5] Cierre
# ------------------------------------------------------------------
Write-Host "============================================================"
Write-Host "  LISTO"
Write-Host "  - Corre DIAGNOSTICO_ACTUALIZACION.bat para confirmar que"
Write-Host "    todo quedo en [OK]."
Write-Host "  - Log completo guardado en: $logPath"
Write-Host "============================================================"
try { Stop-Transcript | Out-Null } catch {}
