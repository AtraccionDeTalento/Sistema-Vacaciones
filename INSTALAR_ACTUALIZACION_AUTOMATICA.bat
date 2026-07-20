@echo off
title Instalar Actualizacion Automatica - Sistema de Vacaciones USIL
color 0A
set "SVU_BAT_DIR=%~dp0"
set "SVU_TMP=%TEMP%\svu_instalar_%RANDOM%.ps1"

powershell -NoProfile -Command "$m='#---PS1-BEGIN---'; $s=[IO.File]::ReadAllText('%~f0'); $i=$s.LastIndexOf($m); $code=$s.Substring($i+$m.Length); [IO.File]::WriteAllText('%SVU_TMP%',$code)"

if not exist "%SVU_TMP%" (
    echo [ERROR] No se pudo preparar el script de instalacion.
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
$nombreTarea = 'SistemaVacaciones_AutoActualizar'

Write-Host "============================================================"
Write-Host "  INSTALAR ACTUALIZACION AUTOMATICA"
Write-Host "  Sistema de Vacaciones USIL"
Write-Host "============================================================"
Write-Host "  Esto se corre UNA SOLA VEZ por PC. Deja una tarea programada"
Write-Host "  de Windows que revisa GitHub cada pocas horas y en cada"
Write-Host "  inicio de sesion, y actualiza los archivos solo si hay"
Write-Host "  cambios -- todo en segundo plano, sin ventanas, sin que"
Write-Host "  nadie tenga que abrir nada nunca mas."
Write-Host "============================================================"
Write-Host ""

# ------------------------------------------------------------------
# [1/4] Ubicar instalacion(es) empaquetadas (misma logica que
# DIAGNOSTICO_ACTUALIZACION.bat / FORZAR_ACTUALIZACION.bat). Sin carpeta
# fuente: no tiene sentido programar auto-actualizacion sobre un repo git.
# ------------------------------------------------------------------
Write-Host "[1/4] Buscando instalaciones empaquetadas en esta PC..."

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
    Write-Host "[ERROR] No se encontro ninguna instalacion. Fin."
    Write-Host ""
    Read-Host "Presiona Enter para cerrar"
    exit 1
}
Write-Host "    Se va(n) a programar $($candidatos.Count) instalacion(es):"
foreach ($c in $candidatos) { Write-Host "      - $($c.ruta)  ($($c.etiqueta))" }
Write-Host ""

# ------------------------------------------------------------------
# [2/4] Escribir, dentro de CADA instalacion encontrada, el script
# silencioso que hace el trabajo real (descarga y reemplaza SOLO si el
# commit remoto cambio). Queda parametrizado con su propia ruta -- cada
# instalacion tiene su propia copia, nada compartido ni fijo entre PCs.
# ------------------------------------------------------------------
Write-Host "[2/4] Escribiendo el script de actualizacion silenciosa..."

$archivosLista = @(
    'servidor.py', 'index_vacaciones.html', 'enviar_cola_outlook.py', '_bp_map.py',
    'requirements.txt', 'assets/js/app_completo.js', 'assets/js/pipeline_vac.js',
    'assets/css/styles.css', 'PIPELINE/motor/pipeline.py', 'PIPELINE/motor/vac_lib.py',
    'PIPELINE/motor/config.json', 'PIPELINE/motor/requirements.txt', 'PIPELINE/bot_adryan/bot_adryan.py',
    'PIPELINE/bot_adryan/bot_maestro.py', 'PIPELINE/bot_adryan/guardar_password.py'
) -join "','"

$plantillaScript = @'
# Auto-generado por INSTALAR_ACTUALIZACION_AUTOMATICA.bat -- no editar a mano.
# Corre en segundo plano (Programador de tareas), sin ventana, sin pedir nada.
# Solo descarga y reemplaza archivos si el commit remoto de GitHub cambio
# desde la ultima vez (no reintenta si ya esta al dia -- evita trafico y
# escrituras a disco innecesarias en cada corrida periodica).
$ErrorActionPreference = 'SilentlyContinue'
$Repo = 'REPO_PLACEHOLDER'
$Branch = 'BRANCH_PLACEHOLDER'
$destinoBase = 'RUTA_PLACEHOLDER'
$logFile = Join-Path $destinoBase 'auto_actualizar.log'
$archivos = @('ARCHIVOS_PLACEHOLDER')

function Log($msg) {
    $linea = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg"
    try {
        if ((Test-Path $logFile) -and (Get-Item $logFile).Length -gt 300KB) {
            Move-Item $logFile "$logFile.old" -Force
        }
        Add-Content -Path $logFile -Value $linea -Encoding UTF8
    } catch {}
}

$remoteVer = $null
try {
    $resp = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/commits/$Branch" -Headers @{ 'User-Agent' = 'svu-auto-update' } -TimeoutSec 15 -UseBasicParsing -ErrorAction Stop
    $remoteVer = $resp.sha
} catch {
    Log "Sin conexion a GitHub, se omite esta corrida: $($_.Exception.Message)"
    exit 0
}

$verFile = Join-Path $destinoBase '.version_commit'
$localVer = $null
if (Test-Path $verFile) { $localVer = (Get-Content $verFile -Raw -ErrorAction SilentlyContinue).Trim() }

if ($localVer -eq $remoteVer) {
    exit 0  # ya esta al dia, no hacer nada (no genera ruido en el log cada corrida)
}

Log "Version distinta detectada (local=$localVer remoto=$remoteVer). Actualizando..."
$ok = 0; $fallidos = @()
foreach ($rel in $archivos) {
    $destino = Join-Path $destinoBase ($rel -replace '/', '\')
    $dir = Split-Path $destino -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $url = "https://raw.githubusercontent.com/$Repo/$Branch/$rel"
    $bajado = $false
    for ($i = 1; $i -le 3; $i++) {
        try {
            $contenido = (Invoke-WebRequest -Uri $url -TimeoutSec 30 -UseBasicParsing -ErrorAction Stop).Content
            if ($contenido -is [string]) { $contenido = [Text.Encoding]::UTF8.GetBytes($contenido) }
            $tmp = "$destino.tmp"
            [IO.File]::WriteAllBytes($tmp, $contenido)
            if (Test-Path $destino) { Remove-Item $destino -Force -ErrorAction Stop }
            Rename-Item $tmp $destino -Force
            $bajado = $true
            break
        } catch { Start-Sleep -Milliseconds 500 }
    }
    if ($bajado) { $ok++ } else { $fallidos += $rel }
}

if ($fallidos.Count -eq 0) {
    Set-Content -Path $verFile -Value $remoteVer -Encoding UTF8 -NoNewline
    Log "Actualizacion completa: $ok/$($archivos.Count) OK. Version local ahora: $remoteVer"
} else {
    Log "Actualizacion INCOMPLETA: $ok/$($archivos.Count) OK, fallaron: $($fallidos -join ', '). No se marca version -- se reintenta en la proxima corrida."
}
'@

foreach ($c in $candidatos) {
    $scriptDestino = Join-Path $c.ruta 'auto_actualizar_silencioso.ps1'
    $contenido = $plantillaScript.Replace('REPO_PLACEHOLDER', $Repo).Replace('BRANCH_PLACEHOLDER', $Branch).Replace('RUTA_PLACEHOLDER', $c.ruta).Replace('ARCHIVOS_PLACEHOLDER', $archivosLista)
    Set-Content -Path $scriptDestino -Value $contenido -Encoding UTF8
    Write-Host "    Escrito: $scriptDestino"
}
Write-Host ""

# ------------------------------------------------------------------
# [3/4] Registrar la tarea programada: al iniciar sesion + repitiendo
# cada 4 horas de forma indefinida mientras la PC este encendida. No
# depende de una hora fija (si la PC estaba apagada a esa hora, nunca
# hubiera corrido) ni de que alguien abra la app.
#
# Si el Programador de Tareas rechaza la creacion por permisos (comun en PCs
# corporativas con politicas que restringen esto a usuarios normales -- se
# confirmo que pasa en esta misma maquina, con y sin schtasks.exe), se cae
# automaticamente a un plan B que NO necesita ningun permiso especial: dejar
# un .vbs en la carpeta de Inicio del usuario (%APPDATA%\...\Startup), que
# Windows ejecuta solo en cada inicio de sesion. Da menos frecuencia (una vez
# por sesion en vez de cada 4 horas) pero funciona en cualquier cuenta.
# ------------------------------------------------------------------
Write-Host "[3/4] Programando la actualizacion automatica..."

function Instalar-EnCarpetaInicio($ruta, $etiqueta) {
    $scriptRuta = Join-Path $ruta 'auto_actualizar_silencioso.ps1'
    $startupDir = [Environment]::GetFolderPath('Startup')
    $nombreVbs = if ($idx -le 1) { 'SistemaVacaciones_AutoActualizar.vbs' } else { "SistemaVacaciones_AutoActualizar_$idx.vbs" }
    $vbsRuta = Join-Path $startupDir $nombreVbs
    # En VBScript, "" dentro de un literal de texto es una comilla literal. Con
    # esto la linea generada corre: powershell.exe ... -File "<scriptRuta>"
    $vbsContenido = @'
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""{0}""", 0, False
'@ -f $scriptRuta
    try {
        Set-Content -Path $vbsRuta -Value $vbsContenido -Encoding ASCII
        Write-Host "    [OK] Se va a actualizar sola en cada inicio de sesion (carpeta de Inicio): $vbsRuta"
        return $true
    } catch {
        Write-Host "    [ERROR] Tampoco se pudo escribir en la carpeta de Inicio: $($_.Exception.Message)"
        return $false
    }
}

$idx = 0
foreach ($c in $candidatos) {
    $idx++
    $nombreTareaFinal = if ($candidatos.Count -eq 1) { $nombreTarea } else { "$nombreTarea`_$idx" }
    $scriptRuta = Join-Path $c.ruta 'auto_actualizar_silencioso.ps1'
    $logrado = $false
    try {
        $accion = New-ScheduledTaskAction -Execute 'powershell.exe' `
            -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptRuta`""
        $disparadorLogon = New-ScheduledTaskTrigger -AtLogOn
        # [TimeSpan]::MaxValue no es un valor de duracion valido para el XML del
        # Programador de Tareas (excede el limite del esquema). 10 anios es, en la
        # practica, "para siempre" para este proposito.
        $disparadorRepetido = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 4) -RepetitionDuration (New-TimeSpan -Days 3650)
        $configuracion = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -Hidden -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

        Unregister-ScheduledTask -TaskName $nombreTareaFinal -Confirm:$false -ErrorAction SilentlyContinue

        Register-ScheduledTask -TaskName $nombreTareaFinal `
            -Action $accion `
            -Trigger @($disparadorLogon, $disparadorRepetido) `
            -Settings $configuracion `
            -Description "Revisa GitHub ($Repo) cada pocas horas y actualiza $($c.ruta) en segundo plano si hay cambios. Creada por INSTALAR_ACTUALIZACION_AUTOMATICA.bat." `
            -ErrorAction Stop | Out-Null

        Write-Host "    [OK] Tarea '$nombreTareaFinal' creada para $($c.ruta) (revisa cada 4h + en cada inicio de sesion)"
        $logrado = $true
    } catch {
        Write-Host "    [AVISO] El Programador de Tareas rechazo la creacion (permisos): $($_.Exception.Message)"
        Write-Host "    Usando plan B: carpeta de Inicio de Windows..."
        $logrado = Instalar-EnCarpetaInicio $c.ruta $c.etiqueta
    }
    if (-not $logrado) {
        Write-Host "    [ERROR] No se pudo programar la actualizacion automatica para $($c.ruta) por ningun metodo."
    }
}
Write-Host ""

# ------------------------------------------------------------------
# [4/4] Correr una vez ahora mismo, para no esperar 4 horas a la primera
# actualizacion y confirmar que funciona de punta a punta.
# ------------------------------------------------------------------
Write-Host "[4/4] Corriendo una actualizacion de prueba ahora mismo..."
foreach ($c in $candidatos) {
    $scriptRuta = Join-Path $c.ruta 'auto_actualizar_silencioso.ps1'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptRuta
    $logRuta = Join-Path $c.ruta 'auto_actualizar.log'
    if (Test-Path $logRuta) {
        Write-Host "    $($c.ruta):"
        Get-Content $logRuta -Tail 5 | ForEach-Object { Write-Host "      $_" }
    } else {
        $verFile = Join-Path $c.ruta '.version_commit'
        $v = if (Test-Path $verFile) { (Get-Content $verFile -Raw).Trim() } else { '(ninguna)' }
        Write-Host "    $($c.ruta): sin cambios que registrar (ya estaba al dia, version: $v)"
    }
}
Write-Host ""

Write-Host "============================================================"
Write-Host "  LISTO -- instalado para siempre"
Write-Host "  De ahora en adelante, esta PC va a revisar GitHub sola cada"
Write-Host "  4 horas y en cada inicio de sesion, y va a actualizar los"
Write-Host "  archivos en segundo plano si hay cambios -- nadie tiene que"
Write-Host "  abrir ningun .bat de nuevo."
Write-Host ""
Write-Host "  Para revisar el historial de una instalacion:"
Write-Host "    <carpeta de instalacion>\auto_actualizar.log"
Write-Host "  Para ver/quitar la tarea programada: abre 'Programador de"
Write-Host "  tareas' de Windows y busca '$nombreTarea'."
Write-Host "============================================================"
