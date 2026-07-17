@echo off
title Reparar Total - Sistema de Vacaciones USIL
color 0C
set "SVU_BAT_DIR=%~dp0"
set "SVU_TMP=%TEMP%\svu_repartotal_%RANDOM%.ps1"

powershell -NoProfile -Command "$m='#---PS1-BEGIN---'; $s=[IO.File]::ReadAllText('%~f0'); $i=$s.LastIndexOf($m); $code=$s.Substring($i+$m.Length); [IO.File]::WriteAllText('%SVU_TMP%',$code)"

if not exist "%SVU_TMP%" (
    echo [ERROR] No se pudo preparar el script de reparacion.
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

$carpetaSalida = if ($env:SVU_BAT_DIR) { $env:SVU_BAT_DIR } else { $env:TEMP }
$txtPath  = Join-Path $carpetaSalida 'REPARAR_TOTAL_RESULTADO.txt'
$jsonPath = Join-Path $carpetaSalida 'REPARAR_TOTAL_RESULTADO.json'
try { Start-Transcript -Path $txtPath -Force | Out-Null } catch {}

Write-Host "============================================================"
Write-Host "  REPARAR TOTAL -- Sistema de Vacaciones USIL"
Write-Host "============================================================"
Write-Host "  A diferencia de FORZAR_ACTUALIZACION.bat (que solo baja"
Write-Host "  archivos), este script hace las 4 cosas necesarias en orden"
Write-Host "  y VERIFICA cada una antes de seguir, para no repetir lo de"
Write-Host "  ayer (archivos bajados sin error, pero el problema seguia):"
Write-Host "    1. MATA a la fuerza cualquier proceso viejo (el archivo"
Write-Host "       puede cambiar en disco, pero si el proceso Python ya"
Write-Host "       esta corriendo, sigue usando el codigo QUE TENIA EN"
Write-Host "       MEMORIA desde que arranco -- Waitress no recarga solo)."
Write-Host "    2. Descarga y reemplaza TODOS los archivos, en TODAS las"
Write-Host "       instalaciones que encuentre en esta PC (no solo una)."
Write-Host "    3. Verifica con SHA256 que lo que quedo en disco es IGUAL"
Write-Host "       byte a byte a lo que hay en GitHub ahora mismo."
Write-Host "    4. Vuelve a abrir la app y consulta su propio servidor"
Write-Host "       para confirmar con hash que el proceso EN MEMORIA (no"
Write-Host "       solo el archivo en disco) es el codigo nuevo."
Write-Host "  Si el paso 4 no queda en [OK], el .bat te dice exactamente"
Write-Host "  cual de los 4 pasos fallo -- no hay que adivinar."
Write-Host "============================================================"
Write-Host ""

$resultado = [ordered]@{
    timestamp_utc = (Get-Date).ToUniversalTime().ToString('o')
    maquina       = $env:COMPUTERNAME
    usuario       = $env:USERNAME
    pasos         = [ordered]@{}
}

# ------------------------------------------------------------------
# [1/6] Ubicar TODAS las copias del sistema en esta PC
# ------------------------------------------------------------------
Write-Host "[1/6] Buscando TODAS las copias del sistema en esta PC..."
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
    Read-Host "Presiona Enter para cerrar"
    exit 1
}
Write-Host "    Se van a reparar TODAS estas copias (no solo una, para no dejar ninguna"
Write-Host "    desactualizada compitiendo con la que la persona realmente abre):"
foreach ($c in $candidatos) { Write-Host "      - [$($c.etiqueta)] $($c.ruta)" }
Write-Host ""
$resultado.pasos['1_instalaciones_encontradas'] = @($candidatos | ForEach-Object { @{ ruta = $_.ruta; etiqueta = $_.etiqueta } })

# ------------------------------------------------------------------
# [2/6] MATAR a la fuerza cualquier proceso viejo -- por nombre de app Y por
# CommandLine (cubre el caso de que alguien corra "python servidor.py" a mano,
# como paso ayer: PID 22936 -> python.exe con servidor.py en la linea de
# comandos, que Get-Process -Name 'python' sin filtro de CommandLine no
# distingue de otro python.exe cualquiera en la PC).
# ------------------------------------------------------------------
Write-Host "[2/6] Matando a la fuerza cualquier proceso viejo de la app..."
$matados = @()
$procsApp = Get-Process -Name 'Sistema Vacaciones USIL' -ErrorAction SilentlyContinue
foreach ($p in $procsApp) {
    Write-Host "    Matando PID $($p.Id) ($($p.Path))"
    Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    $matados += @{ pid = $p.Id; ruta = "$($p.Path)"; motivo = 'proceso Sistema Vacaciones USIL.exe' }
}
$procsPy = Get-CimInstance Win32_Process -Filter "Name='python.exe' OR Name='pythonw.exe'" -ErrorAction SilentlyContinue
foreach ($p in $procsPy) {
    if ($p.CommandLine -match 'servidor\.py' -or $p.CommandLine -match 'SISTEMA DE VACACIONES' -or $p.CommandLine -match 'sistema-vacaciones-usil') {
        Write-Host "    Matando PID $($p.ProcessId) -- $($p.CommandLine)"
        Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
        $matados += @{ pid = $p.ProcessId; comandoLinea = "$($p.CommandLine)"; motivo = 'python.exe corriendo servidor.py de este sistema' }
    }
}
if ($matados.Count -eq 0) {
    Write-Host "    No habia ningun proceso viejo corriendo -- nada que matar."
} else {
    Write-Host "    $($matados.Count) proceso(s) viejo(s) terminado(s). Esperando 3s a que liberen los archivos..."
    Start-Sleep -Seconds 3
}
$resultado.pasos['2_procesos_matados'] = $matados
Write-Host ""

# ------------------------------------------------------------------
# [3/6] Descargar y reemplazar TODOS los archivos, en TODAS las instalaciones
# ------------------------------------------------------------------
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
    # Normaliza CRLF -> LF antes de hashear -- mismo criterio que
    # DIAGNOSTICO_ACTUALIZACION.bat, para que la comparacion sea sobre el
    # CONTENIDO real y no sobre si el archivo local tiene finales de linea
    # de Windows.
    $texto = [Text.Encoding]::UTF8.GetString($bytes) -replace "`r`n", "`n"
    $bytesNorm = [Text.Encoding]::UTF8.GetBytes($texto)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { -join ($sha.ComputeHash($bytesNorm) | ForEach-Object { $_.ToString('x2') }) }
    finally { $sha.Dispose() }
}

Write-Host "[3/6] Descargando archivos desde GitHub y reemplazando en cada instalacion..."
$remoteVer = $null
try {
    $resp = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/commits/$Branch" -Headers @{ 'User-Agent' = 'reparar-total-sistema-vacaciones' } -TimeoutSec 15 -UseBasicParsing -ErrorAction Stop
    $remoteVer = $resp.sha
    Write-Host "    Commit remoto (lo que se va a instalar): $remoteVer"
} catch {
    Write-Host "    [ERROR] No se pudo consultar GitHub: $($_.Exception.Message)"
    Write-Host "    Sin internet no se puede continuar. Revisa la conexion e intenta de nuevo."
    $resultado.pasos['3_error'] = 'sin conexion a GitHub'
    ($resultado | ConvertTo-Json -Depth 10) | Set-Content -Path $jsonPath -Encoding UTF8
    Read-Host "Presiona Enter para cerrar"
    exit 1
}
Write-Host ""

$pasoDescarga = @()
foreach ($obj in $candidatos) {
    Write-Host "------------------------------------------------------------"
    Write-Host "  INSTALACION: $($obj.ruta)"
    Write-Host "------------------------------------------------------------"
    $ok = 0; $verificadosIguales = 0; $fallidos = @()
    foreach ($rel in $archivos) {
        $destino = Join-Path $obj.ruta ($rel -replace '/', '\')
        $dir = Split-Path $destino -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $url = "https://raw.githubusercontent.com/$Repo/$Branch/$rel"
        $bajado = $false
        $contenidoBytes = $null
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
            } catch {
                if ($intento -eq 3) {
                    $msg = "$($_.Exception.Message)"
                    if ($msg.Length -gt 200) { $msg = $msg.Substring(0, 200) + '...' }
                    Write-Host ("    [FALLO DESCARGA] {0} -- {1}" -f $rel, $msg)
                } else {
                    Start-Sleep -Milliseconds 500
                }
            }
        }
        if (-not $bajado) { $fallidos += $rel; continue }

        # Verificacion inmediata: releer lo que quedo en disco y comparar hash
        # contra lo que se bajo -- no confiar en que "no hubo excepcion" =
        # "quedo bien" (un antivirus puede tocar el archivo despues de escrito).
        try {
            $localBytesVerif = [IO.File]::ReadAllBytes($destino)
            if ((Get-Sha256OfBytes $localBytesVerif) -eq (Get-Sha256OfBytes $contenidoBytes)) {
                Write-Host ("    [OK]    {0}" -f $rel)
                $ok++; $verificadosIguales++
            } else {
                Write-Host ("    [AVISO] {0} se escribio pero el hash en disco no coincide con lo bajado (algo lo toco despues)" -f $rel)
                $fallidos += $rel
            }
        } catch {
            Write-Host ("    [AVISO] {0} no se pudo releer para verificar: {1}" -f $rel, $_.Exception.Message)
            $fallidos += $rel
        }
    }
    Write-Host ""
    Write-Host ("    Resultado: {0}/{1} archivos actualizados Y verificados." -f $verificadosIguales, $archivos.Count)
    if ($fallidos.Count -eq 0) {
        try {
            Set-Content -Path (Join-Path $obj.ruta '.version_commit') -Value $remoteVer -Encoding UTF8 -NoNewline
            Write-Host "    [OK] .version_commit actualizado a $remoteVer"
        } catch {}
    } else {
        Write-Host "    [AVISO] Quedaron $($fallidos.Count) archivo(s) sin verificar: $($fallidos -join ', ')"
    }
    Write-Host ""
    $pasoDescarga += @{ ruta = $obj.ruta; ok = $verificadosIguales; total = $archivos.Count; fallidos = $fallidos }
}
$resultado.pasos['3_descarga_y_verificacion'] = $pasoDescarga

# ------------------------------------------------------------------
# [4/6] Reabrir la app (usa el acceso directo del Escritorio si existe, o el
# .exe de la primera instalacion encontrada)
# ------------------------------------------------------------------
Write-Host "[4/6] Reabriendo la app con el codigo nuevo..."
$abierto = $false
try {
    $lnkPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Sistema de Vacaciones USIL.lnk'
    if (Test-Path $lnkPath) {
        Start-Process $lnkPath
        $abierto = $true
        Write-Host "    Abierta via acceso directo del Escritorio."
    }
} catch {}
if (-not $abierto) {
    $exePath = Join-Path $candidatos[0].ruta $exeName
    if (Test-Path $exePath) {
        Start-Process $exePath
        $abierto = $true
        Write-Host "    Abierta via: $exePath"
    }
}
if (-not $abierto) {
    Write-Host "    [ERROR] No se pudo reabrir la app automaticamente -- abrela manualmente ahora."
}
$resultado.pasos['4_reabierto'] = $abierto
Write-Host ""

# ------------------------------------------------------------------
# [5/6] Esperar a que el servidor local responda, y verificar con SHA256 que
# el PROCESO EN MEMORIA (no solo el archivo en disco) es el codigo nuevo.
# Esto es lo que faltaba ayer: confirmar el codigo que esta EJECUTANDO, no
# solo el que quedo guardado.
# ------------------------------------------------------------------
Write-Host "[5/6] Esperando a que el servidor local (puerto 5002) responda..."
$health = $null
for ($i = 1; $i -le 20; $i++) {
    try {
        $health = Invoke-RestMethod -Uri "http://127.0.0.1:5002/api/health" -TimeoutSec 5 -ErrorAction Stop
        if ($health) { break }
    } catch {}
    Start-Sleep -Seconds 2
}
if (-not $health) {
    Write-Host "    [ERROR] El servidor no respondio despues de ~40s. Puede que la app tarde mas en"
    Write-Host "    abrir (primera carga de Excel) -- espera un poco y corre DIAGNOSTICO_DATOS.bat a mano."
    $resultado.pasos['5_servidor_respondio'] = $false
} else {
    Write-Host "    Servidor respondiendo. boot_time: $($health.boot_time)"
    $resultado.pasos['5_servidor_respondio'] = $true
    $resultado.pasos['5_boot_time'] = $health.boot_time

    Write-Host ""
    Write-Host "    Consultando hash de servidor.py EN MEMORIA (via /api/diagnostico/kpis)..."
    try {
        $diag = Invoke-RestMethod -Uri "http://127.0.0.1:5002/api/diagnostico/kpis" -TimeoutSec 60 -ErrorAction Stop
        $hashLocalEnMemoria = $diag.diagnostico.version_codigo.servidor_py_sha256

        if (-not $hashLocalEnMemoria) {
            Write-Host "    [AVISO] Este servidor todavia no tiene el campo version_codigo -- significa que"
            Write-Host "    el servidor.py que quedo instalado sigue siendo mas viejo que el que agrego ese"
            Write-Host "    campo. Si esto pasa, la actualizacion realmente NO se aplico -- revisa el paso 3"
            Write-Host "    arriba por archivos fallidos, y que no haya OTRO python.exe corriendo desde otra"
            Write-Host "    carpeta que no este en la lista de instalaciones detectadas."
            $resultado.pasos['5_hash_verificado'] = $false
        } else {
            $remoteServidorBytes = (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/$Repo/$Branch/servidor.py" -TimeoutSec 30 -UseBasicParsing -ErrorAction Stop).Content
            if ($remoteServidorBytes -is [string]) { $remoteServidorBytes = [Text.Encoding]::UTF8.GetBytes($remoteServidorBytes) }
            $hashRemoto = Get-Sha256OfBytes $remoteServidorBytes

            Write-Host "    Hash EN MEMORIA (lo que el servidor esta ejecutando ahora): $hashLocalEnMemoria"
            Write-Host "    Hash en GitHub (lo que deberia estar corriendo):            $hashRemoto"
            if ($hashLocalEnMemoria -eq $hashRemoto) {
                Write-Host ""
                Write-Host "    [OK] CONFIRMADO: el proceso que esta corriendo AHORA usa el codigo mas"
                Write-Host "    reciente de GitHub. No es un archivo viejo en memoria."
                $resultado.pasos['5_hash_verificado'] = $true
            } else {
                Write-Host ""
                Write-Host "    [FALLO] El proceso en memoria NO coincide con GitHub. El servidor.py en"
                Write-Host "    disco puede estar bien, pero el proceso que respondio sigue siendo OTRO"
                Write-Host "    (ej. quedo un segundo proceso viejo en otro puerto redirigido, o esta PC"
                Write-Host "    tiene una instalacion adicional no detectada en el paso 1). Revisa cuantos"
                Write-Host "    procesos hay con: Get-Process 'Sistema Vacaciones USIL','python','pythonw'"
                $resultado.pasos['5_hash_verificado'] = $false
            }
        }

        Write-Host ""
        Write-Host "    Avance de Meta que muestra AHORA el servidor ya reiniciado:"
        Write-Host "    avance: $([math]::Round($diag.diagnostico.kpis_actuales.avance * 100, 1))%  ($($diag.diagnostico.kpis_actuales.registrado_total) / $($diag.diagnostico.kpis_actuales.meta_total) dias)"
        if ($diag.diagnostico.estado_pipeline_json.paso_validacion_cruzada -eq $false) {
            Write-Host "    (el JSON cacheado del pipeline fue RECHAZADO por inconsistente -- se esta usando el Excel en vivo, correcto)"
        }
        $resultado.pasos['5_avance_final'] = $diag.diagnostico.kpis_actuales
    } catch {
        Write-Host "    [ERROR] No se pudo consultar /api/diagnostico/kpis: $($_.Exception.Message)"
        $resultado.pasos['5_hash_verificado'] = $false
    }
}
Write-Host ""

# ------------------------------------------------------------------
# [6/6] Veredicto final
# ------------------------------------------------------------------
Write-Host "============================================================"
$exito = ($resultado.pasos['5_hash_verificado'] -eq $true)
if ($exito) {
    Write-Host "  RESULTADO: [OK] REPARADO Y VERIFICADO"
    Write-Host "  El proceso corriendo ahora mismo usa el codigo mas reciente."
} else {
    Write-Host "  RESULTADO: [NO CONFIRMADO] Revisa los avisos en rojo arriba."
    Write-Host "  Envia $txtPath y $jsonPath para diagnosticar el siguiente paso."
}
Write-Host "============================================================"
Write-Host "  Archivos generados junto a este .bat:"
Write-Host "      $txtPath"
Write-Host "      $jsonPath"
Write-Host "============================================================"

try {
    ($resultado | ConvertTo-Json -Depth 10) | Set-Content -Path $jsonPath -Encoding UTF8
} catch {}
try { Stop-Transcript | Out-Null } catch {}
