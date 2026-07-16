@echo off
title Diagnostico de Datos (KPIs) - Sistema de Vacaciones USIL
color 0D
set "SVU_BAT_DIR=%~dp0"
set "SVU_TMP=%TEMP%\svu_diagdatos_%RANDOM%.ps1"

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
$carpetaSalida = if ($env:SVU_BAT_DIR) { $env:SVU_BAT_DIR } else { $env:TEMP }
$txtPath  = Join-Path $carpetaSalida 'DIAGNOSTICO_DATOS_RESULTADO.txt'
$jsonPath = Join-Path $carpetaSalida 'DIAGNOSTICO_DATOS_RESULTADO.json'
try { Start-Transcript -Path $txtPath -Force | Out-Null } catch {}

Write-Host "============================================================"
Write-Host "  DIAGNOSTICO DE DATOS (KPIs) - Sistema de Vacaciones USIL"
Write-Host "============================================================"
Write-Host "  Corre esto EN EL MOMENTO en que veas un numero raro en el"
Write-Host "  dashboard (ej. Avance de Meta con un salto que no cuadra)."
Write-Host "  Necesita que la app este ABIERTA (usa su servidor local en"
Write-Host "  el puerto 5002) para poder leer el estado real en ese"
Write-Host "  instante -- si esperas a que se cierre, el momento se pierde."
Write-Host "  Al terminar, envia estos DOS archivos (quedan junto a este .bat):"
Write-Host "    - DIAGNOSTICO_DATOS_RESULTADO.txt"
Write-Host "    - DIAGNOSTICO_DATOS_RESULTADO.json"
Write-Host "============================================================"
Write-Host ""

# ------------------------------------------------------------------
# [1/4] Ver TODOS los procesos python.exe / Electron corriendo -- si hay mas
# de una copia del servidor viva a la vez (algo que paso varias veces
# durante el desarrollo de este sistema), cada una mantiene su propio cache
# en memoria y pueden responder cosas distintas segun cual atienda cada
# request. Esto por si solo puede explicar un numero que salta y luego
# vuelve a bajar sin que ningun archivo haya cambiado de verdad.
# ------------------------------------------------------------------
Write-Host "[1/4] Procesos de Python / la app escuchando en el puerto 5002..."
$conexiones = Get-NetTCPConnection -LocalPort 5002 -State Listen -ErrorAction SilentlyContinue
$procesosServidor = @()
if ($conexiones) {
    $pids = $conexiones | Select-Object -ExpandProperty OwningProcess -Unique
    Write-Host "    $($pids.Count) proceso(s) escuchando en el puerto 5002:"
    foreach ($p in $pids) {
        try {
            $proc = Get-Process -Id $p -ErrorAction Stop
            $cmd = (Get-CimInstance Win32_Process -Filter "ProcessId=$p" -ErrorAction SilentlyContinue).CommandLine
            Write-Host "      PID $p -> $($proc.Path)"
            Write-Host "        Inicio: $($proc.StartTime)  |  CommandLine: $cmd"
            $procesosServidor += @{ pid = $p; ruta = "$($proc.Path)"; inicio = "$($proc.StartTime)"; commandLine = "$cmd" }
        } catch {
            Write-Host "      PID $p (no se pudo leer detalle)"
            $procesosServidor += @{ pid = $p }
        }
    }
    if ($pids.Count -gt 1) {
        Write-Host ""
        Write-Host "    [SOSPECHOSO] Hay MAS DE UN proceso en el mismo puerto. Solo uno de"
        Write-Host "    ellos puede estar realmente atendiendo Windows en un momento dado,"
        Write-Host "    pero si se turnan o si uno cae y otro lo reemplaza a mitad de sesion,"
        Write-Host "    cada uno responde con SU PROPIO cache en memoria -- eso explica un"
        Write-Host "    numero que cambia solo sin que ningun archivo se haya movido."
    }
} else {
    Write-Host "    Nada escuchando en el puerto 5002 -- la app no esta abierta ahora mismo."
    Write-Host "    Abrela primero y vuelve a correr este diagnostico."
}
Write-Host ""

# ------------------------------------------------------------------
# [2/4] Consultar /api/health -- confirma que version de servidor.py tiene
# cargada en memoria el proceso que esta respondiendo ahora.
# ------------------------------------------------------------------
Write-Host "[2/4] Consultando /api/health..."
$health = $null
try {
    $health = Invoke-RestMethod -Uri "http://127.0.0.1:5002/api/health" -TimeoutSec 10 -ErrorAction Stop
    Write-Host "    boot_time: $($health.boot_time)"
    Write-Host "    codigo_desactualizado: $($health.codigo_desactualizado)"
} catch {
    Write-Host "    [ERROR] No se pudo consultar /api/health: $($_.Exception.Message)"
    Write-Host "    (Necesitas la app abierta y este .bat con una version que ya tenga"
    Write-Host "    /api/diagnostico/kpis -- si es una version vieja, actualizala primero"
    Write-Host "    con FORZAR_ACTUALIZACION.bat y vuelve a intentar.)"
}
Write-Host ""

# ------------------------------------------------------------------
# [3/4] El endpoint de diagnostico completo: de donde sale cada KPI, si hay
# una actualizacion corriendo ahora mismo, y los mtimes de cada Excel fuente.
# ------------------------------------------------------------------
Write-Host "[3/4] Consultando /api/diagnostico/kpis (puede tardar si el Excel"
Write-Host "      de vacaciones aun no estaba en cache -- hasta ~25s la primera vez)..."
$diag = $null
try {
    $diag = Invoke-RestMethod -Uri "http://127.0.0.1:5002/api/diagnostico/kpis" -TimeoutSec 60 -ErrorAction Stop
    $d = $diag.diagnostico

    Write-Host ""
    Write-Host "    -- Actualizacion en curso ahora mismo --"
    Write-Host "    actualizacion_en_curso: $($d.actualizacion_en_curso)"
    if ($d.aviso) { Write-Host "    [AVISO] $($d.aviso)" }

    Write-Host ""
    Write-Host "    -- Archivo de vacaciones (fuente de Avance de Meta) --"
    Write-Host "    ruta: $($d.vacaciones_data_file.ruta)"
    Write-Host "    mtime: $($d.vacaciones_data_file.mtime_legible)"

    Write-Host ""
    Write-Host "    -- estado_pipeline.json (cache que puede ganarle al Excel en vivo) --"
    if ($d.estado_pipeline_json.existe) {
        Write-Host "    timestamp declarado: $($d.estado_pipeline_json.timestamp_declarado)  ($($d.estado_pipeline_json.antiguedad_horas)h de antiguedad)"
        Write-Host "    seria_usado_ahora_para_avance: $($d.estado_pipeline_json.seria_usado_ahora_para_avance)"
        Write-Host "    avance segun este JSON: $([math]::Round($d.estado_pipeline_json.avance * 100, 1))%  ($($d.estado_pipeline_json.registrado_total) / $($d.estado_pipeline_json.meta_total) dias)"
    } else {
        Write-Host "    No existe."
    }

    Write-Host ""
    Write-Host "    -- KPIs actuales (lo que el dashboard esta mostrando ahora) --"
    Write-Host "    avance: $([math]::Round($d.kpis_actuales.avance * 100, 1))%  ($($d.kpis_actuales.registrado_total) / $($d.kpis_actuales.meta_total) dias)"
    Write-Host "    con_meta: $($d.meta_vac_counts.con_meta)  cumplieron: $($d.meta_vac_counts.cumplieron)  sin_iniciar: $($d.meta_vac_counts.sin_iniciar)"

    Write-Host ""
    Write-Host "    -- Maestro de personal (filtro anti-fantasma) --"
    Write-Host "    archivo usado: $($d.maestro.archivo_usado)"
    Write-Host "    total personas: $($d.maestro.total_personas)"

    Write-Host ""
    Write-Host "    -- Archivo de objetivos (envio masivo) --"
    Write-Host "    ruta: $($d.objetivos_data_file.ruta)"
} catch {
    Write-Host "    [ERROR] No se pudo consultar /api/diagnostico/kpis: $($_.Exception.Message)"
}
Write-Host ""

# ------------------------------------------------------------------
# [4/4] Guardar resultado estructurado (health + procesos + diagnostico) en
# un solo JSON, para poder leerlo de forma inequivoca despues.
# ------------------------------------------------------------------
$resultado = [ordered]@{
    timestamp_utc    = (Get-Date).ToUniversalTime().ToString('o')
    maquina          = $env:COMPUTERNAME
    usuario          = $env:USERNAME
    procesosServidor = $procesosServidor
    health           = $health
    diagnosticoKpis  = if ($diag) { $diag.diagnostico } else { $null }
}
try {
    ($resultado | ConvertTo-Json -Depth 10) | Set-Content -Path $jsonPath -Encoding UTF8
    Write-Host "Resultado estructurado guardado en: $jsonPath"
} catch {
    Write-Host "[AVISO] No se pudo guardar el JSON: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "============================================================"
Write-Host "  LISTO"
Write-Host "  Envia los dos archivos generados junto a este .bat:"
Write-Host "      $txtPath"
Write-Host "      $jsonPath"
Write-Host "============================================================"
try { Stop-Transcript | Out-Null } catch {}
