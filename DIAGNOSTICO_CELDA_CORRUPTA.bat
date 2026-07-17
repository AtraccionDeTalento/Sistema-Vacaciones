@echo off
title Diagnostico de Celda Corrupta - Sistema de Vacaciones USIL
color 0E
set "SVU_BAT_DIR=%~dp0"
set "SVU_TMP=%TEMP%\svu_diagcelda_%RANDOM%.ps1"

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
$txtPath  = Join-Path $carpetaSalida 'DIAGNOSTICO_CELDA_CORRUPTA_RESULTADO.txt'
$jsonPath = Join-Path $carpetaSalida 'DIAGNOSTICO_CELDA_CORRUPTA_RESULTADO.json'
try { Start-Transcript -Path $txtPath -Force | Out-Null } catch {}

Write-Host "============================================================"
Write-Host "  DIAGNOSTICO DE CELDA CORRUPTA - Sistema de Vacaciones USIL"
Write-Host "============================================================"
Write-Host "  El sistema ya avisa cuando el avance del Excel en vivo no"
Write-Host "  cuadra con el resto de los datos (mensaje [PIPE-KPI] AVISO"
Write-Host "  en la consola de la app), pero no dice CUAL fila es. Este"
Write-Host "  script procesa el mismo Excel con la misma logica del"
Write-Host "  servidor y muestra las filas mas sospechosas de estar"
Write-Host "  inflando el total -- para ir directo a la matricula exacta"
Write-Host "  en vez de revisar 1000+ filas a mano."
Write-Host "  Necesita la app ABIERTA (usa su servidor local, puerto 5002)."
Write-Host "============================================================"
Write-Host ""

Write-Host "Consultando /api/diagnostico/celda-sospechosa (puede tardar si"
Write-Host "el Excel aun no estaba en cache -- hasta ~25s la primera vez)..."
Write-Host ""

$r = $null
try {
    $r = Invoke-RestMethod -Uri "http://127.0.0.1:5002/api/diagnostico/celda-sospechosa" -TimeoutSec 90 -ErrorAction Stop
} catch {
    Write-Host "[ERROR] No se pudo consultar el endpoint: $($_.Exception.Message)"
    Write-Host "(Necesitas la app abierta, y una version que ya tenga este endpoint --"
    Write-Host "si es una version vieja, corre primero REPARAR_TOTAL.bat)"
    try { Stop-Transcript | Out-Null } catch {}
    exit 1
}

if (-not $r.ok) {
    Write-Host "[ERROR] El servidor respondio con error: $($r.error)"
    try { Stop-Transcript | Out-Null } catch {}
    exit 1
}

Write-Host "Filas analizadas en total: $($r.total_filas_analizadas)"
Write-Host ""

if ($r.cantidad_sospechosas -eq 0) {
    Write-Host "============================================================"
    Write-Host "  NO SE ENCONTRARON FILAS NUMERICAMENTE IMPOSIBLES"
    Write-Host "============================================================"
    Write-Host "  Ninguna fila supera 60 dias gozados, ni supera 3x su propio"
    Write-Host "  objetivo, ni tiene un objetivo mayor a 60 dias. El Excel que"
    Write-Host "  esta leyendo ESTA copia del sistema se ve limpio."
    Write-Host ""
    Write-Host "  Si en esta PC el dashboard SIGUE mostrando un numero raro,"
    Write-Host "  la causa mas probable es que el Excel que se esta leyendo en"
    Write-Host "  esta PC es DISTINTO al que genero el aviso -- compara el"
    Write-Host "  'sha256' de este resultado contra DIAGNOSTICO_DATOS.bat"
    Write-Host "  corrido en la PC que SI mostro el problema."
} else {
    Write-Host "============================================================"
    Write-Host "  [SOSPECHOSO] $($r.cantidad_sospechosas) fila(s) numericamente imposible(s)"
    Write-Host "============================================================"
    Write-Host "  Estas filas aportan $($r.dias_gozados_que_aportan_las_sospechosas) dias gozados de un total de"
    Write-Host "  $($r.dias_gozados_con_meta_total) ($($r.porcentaje_del_total_que_explican)% del avance mostrado)."
    Write-Host ""
    foreach ($f in $r.filas_sospechosas) {
        Write-Host "  ------------------------------------------------------------"
        Write-Host "  Matricula: $($f.matricula)   Nombre: $($f.nombre)"
        Write-Host "  Departamento: $($f.departamento)   Area: $($f.area)   HRBP: $($f.hrbp)"
        Write-Host "  Objetivo: $($f.objetivo)   Dias gozados: $($f.dias_gozados)   Vencidas: $($f.vencidas)   Pendientes: $($f.pendientes)"
    }
    Write-Host ""
    Write-Host "  SIGUIENTE PASO: abre el Excel 'Reporte Vacaciones Objetivo...'"
    Write-Host "  hoja BASE GENERAL (y la hoja 'base'), busca la(s) matricula(s)"
    Write-Host "  de arriba, y revisa a mano la celda de dias/objetivo -- ahi"
    Write-Host "  esta el valor corrupto (formato de fecha mal leido como numero,"
    Write-Host "  copy-paste que arrastro una formula, etc)."
}

Write-Host ""
Write-Host "  -- Top 15 general (sin filtrar, para contexto) --"
foreach ($f in $r.top_15_general_sin_filtrar) {
    Write-Host ("    {0,-12} {1,-40} obj={2,-6} gozados={3}" -f $f.matricula, $f.nombre, $f.objetivo, $f.dias_gozados)
}

Write-Host ""
Write-Host "============================================================"
Write-Host "  LISTO. Envia estos dos archivos:"
Write-Host "      $txtPath"
Write-Host "      $jsonPath"
Write-Host "============================================================"

try {
    ($r | ConvertTo-Json -Depth 10) | Set-Content -Path $jsonPath -Encoding UTF8
} catch {}
try { Stop-Transcript | Out-Null } catch {}
