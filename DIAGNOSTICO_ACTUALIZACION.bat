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

# Carpeta donde vive el .bat (no una ruta fija de PC: cada maquina lo puede
# tener en otro lado). Ahi mismo se guardan los resultados, para que sea facil
# de encontrar y de reenviar. Dos salidas:
#  - .txt: transcripcion legible completa (para leer/copiar a mano)
#  - .json: los mismos datos en formato estructurado, para que se puedan leer
#    de forma automatica y generar un .bat de arreglo especifico para esta PC
#    sin tener que adivinar nada a partir de texto libre.
$carpetaSalida = if ($env:SVU_BAT_DIR) { $env:SVU_BAT_DIR } else { $env:TEMP }
$txtPath  = Join-Path $carpetaSalida 'DIAGNOSTICO_RESULTADO.txt'
$jsonPath = Join-Path $carpetaSalida 'DIAGNOSTICO_RESULTADO.json'
try { Start-Transcript -Path $txtPath -Force | Out-Null } catch {}

$resultado = [ordered]@{
    timestamp_utc     = (Get-Date).ToUniversalTime().ToString('o')
    maquina           = $env:COMPUTERNAME
    usuario           = $env:USERNAME
    versionRemota     = $null
    conectividad      = @{}
    proxyWinHttp      = $null
    instalaciones     = @()
    procesosActivos   = @()
}

Write-Host "============================================================"
Write-Host "  DIAGNOSTICO DE ACTUALIZACION AUTOMATICA"
Write-Host "  Sistema de Vacaciones USIL"
Write-Host "============================================================"
Write-Host "  Ejecuta este archivo en la PC que NO se esta actualizando."
Write-Host "  Revisa TODAS las copias del sistema que encuentre en esta PC"
Write-Host "  (puede haber mas de una: una carpeta fuente/dev y una o mas"
Write-Host "  instalaciones empaquetadas -- cada una se actualiza por separado)."
Write-Host "  Al terminar, envia estos DOS archivos (quedan junto a este .bat):"
Write-Host "    - DIAGNOSTICO_RESULTADO.txt"
Write-Host "    - DIAGNOSTICO_RESULTADO.json"
Write-Host "============================================================"
Write-Host ""

# ------------------------------------------------------------------
# [1/6] Ubicar TODAS las copias del sistema en esta PC, no solo la primera
# que aparezca. En una PC de desarrollo puede convivir la carpeta fuente
# (de donde se sube a GitHub, SIEMPRE al dia por definicion) junto con una o
# mas instalaciones empaquetadas viejas (dist_electron/win-unpacked, la
# instalada via el instalador, o la que apunte el acceso directo del
# Escritorio) -- cada una se actualiza de forma independiente y confundir
# una con otra es justamente lo que generaba falsas alarmas de "desactualizado".
# ------------------------------------------------------------------
Write-Host "[1/6] Buscando TODAS las copias del sistema en esta PC..."

$exeName = 'Sistema Vacaciones USIL.exe'
$candidatosRaw = New-Object System.Collections.Generic.List[object]

function Add-Candidato($ruta, $etiqueta) {
    if ($ruta) { $candidatosRaw.Add([pscustomobject]@{ ruta = $ruta.TrimEnd('\'); etiqueta = $etiqueta }) }
}

Add-Candidato (Join-Path $env:LOCALAPPDATA 'Programs\sistema-vacaciones-usil') 'Instalacion estandar (LOCALAPPDATA)'
Add-Candidato (Join-Path $env:SVU_BAT_DIR 'dist_electron\win-unpacked') 'Build empaquetado local (dist_electron/win-unpacked)'

# Resolver el acceso directo del Escritorio (main.js lo crea apuntando al
# .exe real de la instalacion en uso) -- suele ser la forma mas confiable de
# encontrar la instalacion que la persona realmente abre todos los dias.
try {
    $shell = New-Object -ComObject WScript.Shell
    $lnkPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Sistema de Vacaciones USIL.lnk'
    if (Test-Path $lnkPath) {
        $target = $shell.CreateShortcut($lnkPath).TargetPath
        if ($target -and (Test-Path $target)) {
            Add-Candidato (Split-Path $target -Parent) 'Acceso directo del Escritorio'
        }
    }
} catch {}

# La propia carpeta donde vive este .bat, SI tiene servidor.py directo (es
# decir, si este diagnostico se corrio desde el repo fuente). Se distingue
# de una instalacion empaquetada por la presencia de una carpeta .git.
if ($env:SVU_BAT_DIR -and (Test-Path (Join-Path $env:SVU_BAT_DIR 'servidor.py'))) {
    $etiquetaFuente = if (Test-Path (Join-Path $env:SVU_BAT_DIR '.git')) { 'CARPETA FUENTE (repo de desarrollo)' } else { 'Carpeta con servidor.py (raiz del .bat)' }
    Add-Candidato $env:SVU_BAT_DIR $etiquetaFuente
}

# Deduplicar por ruta absoluta (insensible a mayusculas), preferir la
# primera etiqueta encontrada para cada ruta.
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

# Filtrar a candidatos validos: o bien tienen el .exe (instalacion empaquetada)
# o bien tienen servidor.py + .git (carpeta fuente).
# @(...) fuerza que el resultado sea SIEMPRE un array: si Where-Object deja un
# solo elemento, PowerShell lo "desenvuelve" a un objeto suelto sin .Count,
# lo que rompia el flujo en cualquier PC con exactamente UNA instalacion (el
# caso mas comun en una maquina de usuario final, sin carpeta fuente).
$instalacionesValidas = @($candidatos | Where-Object {
    (Test-Path (Join-Path $_.ruta $exeName)) -or (Test-Path (Join-Path $_.ruta '.git'))
})

if (-not $instalacionesValidas -or $instalacionesValidas.Count -eq 0) {
    Write-Host "    No se encontro ninguna copia en las rutas conocidas."
    Write-Host ""
    Write-Host "    Truco: abre el icono de 'Sistema de Vacaciones USIL' en el Escritorio o Menu"
    Write-Host "    Inicio, clic derecho -> Abrir ubicacion del archivo. Esa carpeta es la que"
    Write-Host "    necesitas pegar abajo (arrastrala a esta ventana para pegar la ruta)."
    Write-Host ""
    $manual = Read-Host "    Pega aqui la ruta de la CARPETA de instalacion (donde esta el .exe) y presiona Enter"
    if ($manual) {
        $manual = $manual.Trim().Trim('"').TrimEnd('\')
        if (Test-Path $manual) {
            $instalacionesValidas = @([pscustomobject]@{ ruta = $manual; etiqueta = 'Ruta indicada manualmente' })
        }
    }
}

if (-not $instalacionesValidas -or $instalacionesValidas.Count -eq 0) {
    Write-Host ""
    Write-Host "[ERROR] No se pudo determinar ninguna carpeta de instalacion. Fin del diagnostico."
    Write-Host ""
    Read-Host "Presiona Enter para cerrar"
    exit 1
}

Write-Host "    Se encontraron $($instalacionesValidas.Count) copia(s) del sistema en esta PC:"
foreach ($inst in $instalacionesValidas) {
    Write-Host "      - [$($inst.etiqueta)] $($inst.ruta)"
}
Write-Host ""

# ------------------------------------------------------------------
# [2/6] Conectividad a GitHub + deteccion de proxy (una sola vez, es global,
# no depende de cual copia se este revisando)
# ------------------------------------------------------------------
Write-Host "[2/6] Probando conexion a GitHub..."
foreach ($t in 'https://github.com', 'https://api.github.com', 'https://raw.githubusercontent.com') {
    try {
        $r = Invoke-WebRequest -Uri $t -Method Head -TimeoutSec 10 -UseBasicParsing -ErrorAction Stop
        Write-Host ("    {0,-40} HTTP {1}" -f $t, [int]$r.StatusCode)
        $resultado.conectividad[$t] = [int]$r.StatusCode
    } catch {
        $resp = $_.Exception.Response
        if ($resp) {
            Write-Host ("    {0,-40} HTTP {1}" -f $t, [int]$resp.StatusCode)
            $resultado.conectividad[$t] = [int]$resp.StatusCode
        } else {
            Write-Host ("    {0,-40} SIN RESPUESTA: {1}" -f $t, $_.Exception.Message)
            $resultado.conectividad[$t] = "SIN_RESPUESTA: $($_.Exception.Message)"
        }
    }
}
Write-Host ""
Write-Host "    Proxy configurado en Windows (WinHTTP):"
$proxyTxt = netsh winhttp show proxy
$proxyTxt | ForEach-Object { Write-Host "    $_" }
$resultado.proxyWinHttp = ($proxyTxt -join ' | ')
if ($env:HTTP_PROXY -or $env:HTTPS_PROXY) {
    Write-Host "    Variables de entorno detectadas: HTTP_PROXY=$($env:HTTP_PROXY) HTTPS_PROXY=$($env:HTTPS_PROXY)"
}
Write-Host "    [NOTA] La app se conecta con Node.js directo y NO usa el proxy de Windows"
Write-Host "           ni variables HTTP_PROXY/HTTPS_PROXY. Si esta PC necesita proxy para"
Write-Host "           salir a internet, la auto-actualizacion falla aunque el navegador funcione."
Write-Host ""

# ------------------------------------------------------------------
# [3/6] Ultimo commit en GitHub (una sola vez, es la referencia contra la que
# se compara CADA copia encontrada)
# ------------------------------------------------------------------
Write-Host "[3/6] Consultando ultimo commit en GitHub (rama $Branch)..."
$remoteVer = $null
try {
    $resp2 = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/commits/$Branch" -Headers @{ 'User-Agent' = 'diag-sistema-vacaciones' } -TimeoutSec 15 -UseBasicParsing -ErrorAction Stop
    $remoteVer = $resp2.sha
    Write-Host "    Commit remoto (GitHub): $remoteVer"
} catch {
    Write-Host "    [ERROR] No se pudo consultar GitHub: $($_.Exception.Message)"
}
$resultado.versionRemota = $remoteVer
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
    # Normaliza CRLF -> LF antes de hashear. Sin esto, una carpeta fuente
    # (clonada con Git en Windows, donde core.autocrlf suele venir en 'true')
    # tiene los archivos de texto con CRLF en disco, mientras que
    # raw.githubusercontent.com sirve el blob tal cual se guardo (LF) -- eso
    # daba falsos "DESACTUALIZADO" en archivos que en realidad son identicos
    # (confirmado con 'git diff origin/main', que no muestra ninguna diferencia
    # real). Con esta normalizacion la comparacion es sobre el CONTENIDO, no
    # sobre que herramienta escribio el archivo en disco. Todos los archivos
    # de ARCHIVOS_ACTUALIZABLES son texto (.py/.html/.js/.css/.json/.txt), asi
    # que decodificar como UTF8 y reemplazar CRLF->LF es seguro y rapido
    # (evita iterar byte por byte, lento en PowerShell para archivos grandes).
    $texto = [Text.Encoding]::UTF8.GetString($bytes) -replace "`r`n", "`n"
    $bytes = [Text.Encoding]::UTF8.GetBytes($texto)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { -join ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) }
    finally { $sha.Dispose() }
}

# Revisa UNA copia del sistema de punta a punta: version local, servidor.py,
# archivo por archivo contra GitHub, y su update.log. Se llama una vez por
# cada copia encontrada en el paso [1/6], para que ninguna quede fuera y no
# se confunda una copia desactualizada con otra que si esta al dia.
function Test-Instalacion($ruta, $etiqueta) {
    Write-Host "------------------------------------------------------------"
    Write-Host "  COPIA: $ruta"
    Write-Host "  ($etiqueta)"
    Write-Host "------------------------------------------------------------"

    $esFuente = Test-Path (Join-Path $ruta '.git')
    if ($esFuente) {
        Write-Host "    Esta es la CARPETA FUENTE (repo de desarrollo): por definicion siempre"
        Write-Host "    esta al dia, es de aqui que se sube a GitHub. No se compara contra si misma"
        Write-Host "    de forma util -- se muestra solo para que quede claro que NO es una"
        Write-Host "    instalacion empaquetada y no depende del auto-actualizador."
        Write-Host ""
    }

    $entradaInst = [ordered]@{
        ruta             = $ruta
        etiqueta         = $etiqueta
        esCarpetaFuente  = $esFuente
        versionLocal     = $null
        alDiaPorVersion  = $null
        servidorPyExiste = $false
        servidorPyMod    = $null
        archivos         = @()
        resumenArchivos  = @{ ok = 0; desactualizados = 0; faltantes = 0; errores = 0 }
        updateLogTail    = @()
    }

    if (-not $esFuente) {
        $verFile = Join-Path $ruta '.version_commit'
        if (Test-Path $verFile) {
            $entradaInst.versionLocal = (Get-Content $verFile -Raw).Trim()
            Write-Host "    Commit local: $($entradaInst.versionLocal)"
        } else {
            Write-Host "    [AVISO] No existe .version_commit -- esta instalacion nunca completo una actualizacion."
        }
        if ($remoteVer -and $entradaInst.versionLocal) {
            $entradaInst.alDiaPorVersion = ($entradaInst.versionLocal -eq $remoteVer)
            Write-Host $(if ($entradaInst.alDiaPorVersion) { "    [OK] Esta copia esta al dia con GitHub." } else { "    [DESACTUALIZADO] Local: $($entradaInst.versionLocal) / Remoto: $remoteVer" })
        }
        Write-Host ""
    }

    $srv = Join-Path $ruta 'servidor.py'
    if (Test-Path $srv) {
        $entradaInst.servidorPyExiste = $true
        $entradaInst.servidorPyMod = (Get-Item $srv).LastWriteTime.ToString('o')
        Write-Host "    servidor.py encontrado. Modificado: $((Get-Item $srv).LastWriteTime)"
    } else {
        Write-Host "    [ERROR] No se encontro servidor.py en $ruta -- instalacion incompleta o corrupta."
    }
    Write-Host ""

    Write-Host "    Comparando archivo por archivo contra GitHub..."
    $okCount = 0; $desactCount = 0; $faltaCount = 0; $errCount = 0
    foreach ($rel in $archivos) {
        $local = Join-Path $ruta ($rel -replace '/', '\')
        $entrada = [ordered]@{ ruta = $rel; estado = $null; modificadoLocal = $null; detalle = $null }
        if (-not (Test-Path $local)) {
            Write-Host ("      [FALTA]          {0}" -f $rel)
            $faltaCount++
            $entrada.estado = 'FALTA'
            $entradaInst.archivos += $entrada
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
            $entrada.modificadoLocal = (Get-Item $local).LastWriteTime.ToString('o')
            if ($localHash -eq $remoteHash) {
                Write-Host ("      [OK]             {0}" -f $rel)
                $okCount++
                $entrada.estado = 'OK'
            } else {
                Write-Host ("      [DESACTUALIZADO] {0}  (modificado local: {1})" -f $rel, $entrada.modificadoLocal)
                $desactCount++
                $entrada.estado = 'DESACTUALIZADO'
            }
        } catch {
            Write-Host ("      [ERROR AL COMPARAR] {0} -- {1}" -f $rel, $_.Exception.Message)
            $errCount++
            $entrada.estado = 'ERROR'
            $entrada.detalle = $_.Exception.Message
        }
        $entradaInst.archivos += $entrada
    }
    $entradaInst.resumenArchivos = @{ ok = $okCount; desactualizados = $desactCount; faltantes = $faltaCount; errores = $errCount }
    Write-Host ""
    Write-Host ("    Resumen de esta copia: {0} al dia, {1} desactualizados, {2} faltantes, {3} no se pudieron comparar" -f $okCount, $desactCount, $faltaCount, $errCount)
    if (-not $esFuente -and ($desactCount -gt 0 -or $faltaCount -gt 0)) {
        Write-Host "    [DIAGNOSTICO] Esta copia tiene una actualizacion INCOMPLETA: unos archivos se"
        Write-Host "    bajaron y otros no. Suele pasar cuando la conexion se corta a mitad de la"
        Write-Host "    descarga. Cierra la app por completo (Administrador de tareas si hace falta)"
        Write-Host "    y abrela de nuevo con buena conexion; el auto-actualizador ahora reintenta"
        Write-Host "    el set completo si detecta que algo quedo a medias."
    }
    Write-Host ""

    $logFile = Join-Path $ruta 'update.log'
    if (Test-Path $logFile) {
        Write-Host "    Ultimas lineas de update.log:"
        $tail = Get-Content $logFile -Tail 15
        $tail | ForEach-Object { Write-Host "      $_" }
        $entradaInst.updateLogTail = @($tail)
    } else {
        Write-Host "    No existe update.log en esta copia -- nunca ejecuto el chequeo de actualizacion aqui."
    }
    Write-Host ""

    return $entradaInst
}

# ------------------------------------------------------------------
# [4/6] Revisar cada copia encontrada en el paso [1/6]
# ------------------------------------------------------------------
Write-Host "[4/6] Revisando cada copia encontrada ($($instalacionesValidas.Count))..."
Write-Host ""
foreach ($inst in $instalacionesValidas) {
    $resultado.instalaciones += (Test-Instalacion $inst.ruta $inst.etiqueta)
}

# ------------------------------------------------------------------
# [5/6] Procesos de la app corriendo ahora mismo (ayuda a saber cual copia
# es la que la persona esta usando de verdad en este momento)
# ------------------------------------------------------------------
Write-Host "[5/6] Procesos 'Sistema Vacaciones USIL' activos ahora mismo..."
$procs = Get-Process -Name 'Sistema Vacaciones USIL' -ErrorAction SilentlyContinue
if ($procs) {
    $procs | ForEach-Object {
        Write-Host "    PID $($_.Id) -> $($_.Path)"
        $resultado.procesosActivos += @{ pid = $_.Id; ruta = "$($_.Path)" }
    }
} else {
    Write-Host "    No hay ninguna instancia abierta ahora mismo."
}
Write-Host ""

# ------------------------------------------------------------------
# [6/6] Resumen final + guardar resultados
# ------------------------------------------------------------------
Write-Host "============================================================"
Write-Host "  RESUMEN POR COPIA"
Write-Host "============================================================"
foreach ($i in $resultado.instalaciones) {
    $veredicto = if ($i.esCarpetaFuente) { 'CARPETA FUENTE (no aplica auto-actualizador)' }
                 elseif ($i.resumenArchivos.desactualizados -gt 0 -or $i.resumenArchivos.faltantes -gt 0) { 'DESACTUALIZADA' }
                 elseif ($i.resumenArchivos.ok -gt 0) { 'AL DIA' }
                 else { 'NO SE PUDO VERIFICAR' }
    Write-Host "  [$veredicto] $($i.ruta)  ($($i.etiqueta))"
}
Write-Host "============================================================"
Write-Host ""
Write-Host "  - Si el paso 2 muestra 'SIN RESPUESTA' en github.com, hay un"
Write-Host "    firewall, antivirus o proxy bloqueando la conexion."
Write-Host "  - Si alguna copia sale DESACTUALIZADA (que no sea la carpeta fuente),"
Write-Host "    cierra la app por completo (Administrador de tareas si hace falta)"
Write-Host "    y vuelve a abrirla: se actualiza sola al iniciar, si hay internet."
Write-Host "  - Envia los dos archivos generados junto a este .bat:"
Write-Host "      $txtPath"
Write-Host "      $jsonPath"
Write-Host "============================================================"
Write-Host ""

try {
    ($resultado | ConvertTo-Json -Depth 8) | Set-Content -Path $jsonPath -Encoding UTF8
    Write-Host "Resultado estructurado guardado en: $jsonPath"
} catch {
    Write-Host "[AVISO] No se pudo guardar el JSON de resultado: $($_.Exception.Message)"
}
try { Stop-Transcript | Out-Null } catch {}
Write-Host "Transcripcion completa guardada en: $txtPath"
