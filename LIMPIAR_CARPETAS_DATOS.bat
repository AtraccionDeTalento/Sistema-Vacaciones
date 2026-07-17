@echo off
title Limpiar Carpetas de Datos - Sistema de Vacaciones USIL
color 0E
set "SVU_BAT_DIR=%~dp0"
set "SVU_TMP=%TEMP%\svu_limpiar_%RANDOM%.ps1"

powershell -NoProfile -Command "$m='#---PS1-BEGIN---'; $s=[IO.File]::ReadAllText('%~f0'); $i=$s.LastIndexOf($m); $code=$s.Substring($i+$m.Length); [IO.File]::WriteAllText('%SVU_TMP%',$code)"

if not exist "%SVU_TMP%" (
    echo [ERROR] No se pudo preparar el script de limpieza.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%SVU_TMP%"

del "%SVU_TMP%" >nul 2>&1
echo.
pause
exit /b 0

#---PS1-BEGIN---
$ErrorActionPreference = 'Stop'
$exeName = 'Sistema Vacaciones USIL.exe'
$BatDir = $env:SVU_BAT_DIR
if (-not $BatDir) { $BatDir = (Get-Location).Path }
$BatDir = $BatDir.TrimEnd('\')

Write-Host "============================================================"
Write-Host "  LIMPIAR CARPETAS DE DATOS - Sistema de Vacaciones USIL"
Write-Host "============================================================"
Write-Host "  Mueve (NO borra) todos los .xlsx / .xls / .xlsm de DATAS/ y"
Write-Host "  DATA SENSIBLE/ a una carpeta de respaldo con fecha, dejando"
Write-Host "  esas dos carpetas vacias -- listas para que la proxima"
Write-Host "  descarga (bot de Adryan / pipeline) sea la UNICA version"
Write-Host "  presente, sin copias viejas compitiendo o confundiendo al"
Write-Host "  sistema sobre cual archivo usar."
Write-Host "  No importa desde donde se copie/corra este .bat -- busca solo"
Write-Host "  la instalacion real del sistema en esta PC (misma logica que"
Write-Host "  FORZAR_ACTUALIZACION.bat / DIAGNOSTICO_ACTUALIZACION.bat)."
Write-Host ""

# ------------------------------------------------------------------
# [0/4] Encontrar la instalacion real -- NO asumir que este .bat esta junto
# a DATAS/DATA SENSIBLE. Se probo hoy: copiado a Descargas por error, no
# encontro nada y no aviso con claridad de que estaba buscando en el lugar
# equivocado. Ahora busca en las rutas conocidas + carpeta fuente (.git) +
# acceso directo del Escritorio, igual que los otros .bat de la suite.
# ------------------------------------------------------------------
Write-Host "[0/4] Buscando la instalacion real del sistema en esta PC..."
$candidatosRaw = New-Object System.Collections.Generic.List[object]
function Add-Candidato($ruta, $etiqueta) {
    if ($ruta) { $candidatosRaw.Add([pscustomobject]@{ ruta = $ruta.TrimEnd('\'); etiqueta = $etiqueta }) }
}
Add-Candidato (Join-Path $env:LOCALAPPDATA 'Programs\sistema-vacaciones-usil') 'Instalacion estandar (LOCALAPPDATA)'
Add-Candidato (Join-Path $BatDir 'dist_electron\win-unpacked') 'Build empaquetado local (dist_electron/win-unpacked, junto al .bat)'
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
# Valida: tiene el .exe (instalacion empaquetada), O tiene servidor.py directo
# (carpeta fuente/dev, con o sin .git).
$candidatos = @($candidatos | Where-Object {
    (Test-Path (Join-Path $_.ruta $exeName)) -or (Test-Path (Join-Path $_.ruta 'servidor.py'))
})

if (-not $candidatos -or $candidatos.Count -eq 0) {
    Write-Host "    No se encontro ninguna instalacion en las rutas conocidas."
    Write-Host ""
    Write-Host "    Truco: abre el icono de 'Sistema de Vacaciones USIL' en el Escritorio o Menu"
    Write-Host "    Inicio, clic derecho -> Abrir ubicacion del archivo. Esa carpeta es la que"
    Write-Host "    necesitas pegar abajo (arrastrala a esta ventana para pegar la ruta)."
    Write-Host ""
    $manual = Read-Host "    Pega aqui la ruta de la CARPETA de instalacion (donde esta el .exe o servidor.py) y presiona Enter"
    if ($manual) {
        $manual = $manual.Trim().Trim('"').TrimEnd('\')
        if (Test-Path $manual) { $candidatos = @([pscustomobject]@{ ruta = $manual; etiqueta = 'Ruta indicada manualmente' }) }
    }
}
if (-not $candidatos -or $candidatos.Count -eq 0) {
    Write-Host ""
    Write-Host "[ERROR] No se encontro ninguna instalacion. Fin."
    Read-Host "Presiona Enter para cerrar"
    exit 1
}

$Aqui = $null
if ($candidatos.Count -eq 1) {
    $Aqui = $candidatos[0].ruta
    Write-Host "    Instalacion encontrada: $Aqui  ($($candidatos[0].etiqueta))"
} else {
    Write-Host "    Se encontraron $($candidatos.Count) instalaciones:"
    for ($i = 0; $i -lt $candidatos.Count; $i++) {
        Write-Host "      [$($i+1)] $($candidatos[$i].ruta)  ($($candidatos[$i].etiqueta))"
    }
    Write-Host ""
    $sel = Read-Host "    Escribe el numero de la que quieres limpiar"
    $idx = 0
    if ([int]::TryParse($sel, [ref]$idx) -and $idx -ge 1 -and $idx -le $candidatos.Count) {
        $Aqui = $candidatos[$idx - 1].ruta
    }
}
if (-not $Aqui) {
    Write-Host "[ERROR] Seleccion invalida. Fin."
    Read-Host "Presiona Enter para cerrar"
    exit 1
}
Write-Host ""

$carpetasDatos = @('DATAS', 'DATA SENSIBLE')
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$respaldo = Join-Path $Aqui "_RESPALDO_LIMPIEZA_$stamp"
Write-Host "  Carpeta de esta instalacion: $Aqui"
Write-Host "  Los archivos NO se pierden -- quedan en:"
Write-Host "    $respaldo"
Write-Host "============================================================"
Write-Host ""

# ------------------------------------------------------------------
# [1/4] Avisar si la app / el bot estan corriendo -- mover archivos que
# Excel/el servidor tienen abiertos puede fallar o dejar el Excel bloqueado
# a medio mover. Mejor advertir antes que a mitad de la operacion.
# ------------------------------------------------------------------
Write-Host "[1/4] Revisando si algo tiene los archivos abiertos ahora mismo..."
$procsApp = Get-Process -Name 'Sistema Vacaciones USIL' -ErrorAction SilentlyContinue
$procsPy  = Get-CimInstance Win32_Process -Filter "Name='python.exe' OR Name='pythonw.exe'" -ErrorAction SilentlyContinue |
            Where-Object { $_.CommandLine -match 'servidor\.py' -or $_.CommandLine -match [regex]::Escape($Aqui) }
$procsExcel = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
if ($procsApp -or $procsPy -or $procsExcel) {
    Write-Host "    Se detecto la app y/o Excel abiertos:"
    $procsApp   | ForEach-Object { Write-Host "      - Sistema Vacaciones USIL.exe (PID $($_.Id))" }
    $procsPy    | ForEach-Object { Write-Host "      - python.exe (PID $($_.ProcessId))" }
    $procsExcel | ForEach-Object { Write-Host "      - EXCEL.EXE (PID $($_.Id))" }
    Write-Host ""
    $cerrar = Read-Host "    ¿Cerrarlos ahora para mover los archivos sin bloqueos? (S/N)"
    if ($cerrar -and $cerrar.Trim().ToUpperInvariant() -eq 'S') {
        $procsApp | Stop-Process -Force -ErrorAction SilentlyContinue
        $procsPy | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
        $procsExcel | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Write-Host "    Cerrados."
    } else {
        Write-Host "    Continuando sin cerrarlos -- si algun archivo no se puede mover, ciérralos e intenta de nuevo."
    }
}
Write-Host ""

# ------------------------------------------------------------------
# [2/4] Mover los .xlsx/.xls/.xlsm de cada carpeta al respaldo
# ------------------------------------------------------------------
Write-Host "[2/4] Moviendo archivos a respaldo..."
$totalMovidos = 0
$fallidos = @()
foreach ($carpeta in $carpetasDatos) {
    $rutaCarpeta = Join-Path $Aqui $carpeta
    if (-not (Test-Path $rutaCarpeta)) {
        Write-Host "    [AVISO] No existe la carpeta '$carpeta' en esta instalacion -- se omite."
        continue
    }
    $destino = Join-Path $respaldo $carpeta
    New-Item -ItemType Directory -Path $destino -Force | Out-Null

    $archivos = Get-ChildItem -Path $rutaCarpeta -File -Include '*.xlsx', '*.xls', '*.xlsm' -ErrorAction SilentlyContinue
    if (-not $archivos -or $archivos.Count -eq 0) {
        Write-Host "    $carpeta -- ya estaba vacia (0 archivos Excel)."
        continue
    }
    Write-Host "    $carpeta -- moviendo $($archivos.Count) archivo(s)..."
    foreach ($f in $archivos) {
        try {
            Move-Item -LiteralPath $f.FullName -Destination (Join-Path $destino $f.Name) -Force -ErrorAction Stop
            Write-Host "      [OK] $($f.Name)"
            $totalMovidos++
        } catch {
            Write-Host "      [FALLO] $($f.Name) -- $($_.Exception.Message)"
            $fallidos += $f.FullName
        }
    }
}
Write-Host ""

# ------------------------------------------------------------------
# [3/4] Resumen + tambien limpiar el cache pickle, que quedaria apuntando a
# un mtime de un archivo que ya no esta -- forzar que la proxima lectura
# recompute desde cero contra el Excel nuevo que se vaya a descargar.
# ------------------------------------------------------------------
$cacheDir = Join-Path $Aqui 'DATAS\__cache__'
if (Test-Path $cacheDir) {
    $pkls = Get-ChildItem -Path $cacheDir -Filter '*.pkl' -ErrorAction SilentlyContinue
    if ($pkls) {
        Write-Host "[3/4] Limpiando cache (.pkl) para forzar recalculo con los datos frescos..."
        foreach ($p in $pkls) {
            try {
                Move-Item -LiteralPath $p.FullName -Destination (Join-Path $respaldo ("cache_" + $p.Name)) -Force -ErrorAction Stop
                Write-Host "    [OK] $($p.Name)"
            } catch {
                Write-Host "    [AVISO] No se pudo mover $($p.Name): $($_.Exception.Message)"
            }
        }
    }
}
Write-Host ""

Write-Host "============================================================"
if ($fallidos.Count -eq 0) {
    Write-Host "  LISTO. $totalMovidos archivo(s) movidos a respaldo."
    Write-Host "  DATAS/ y DATA SENSIBLE/ quedaron limpias."
    Write-Host ""
    Write-Host "  SIGUIENTE PASO: correr PIPELINE\bot_adryan\ACTUALIZAR_AUTO.bat"
    Write-Host "  para bajar maestro + vacaciones frescos directo de Adryan."
} else {
    Write-Host "  [AVISO] $totalMovidos movidos OK, pero $($fallidos.Count) fallaron:"
    $fallidos | ForEach-Object { Write-Host "    - $_" }
    Write-Host "  Cierra Excel/la app por completo y vuelve a correr este .bat."
}
Write-Host "  Respaldo (para recuperar si algo sale mal):"
Write-Host "    $respaldo"
Write-Host "============================================================"
