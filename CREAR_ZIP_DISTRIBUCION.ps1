# CREAR_ZIP_DISTRIBUCION.ps1
# Genera un ZIP limpio del Sistema de Vacaciones listo para distribuir.
# Excluye: .venv, __pycache__, .git, node_modules, datos sensibles, logs, secrets.
#
# Uso: clic derecho -> Ejecutar con PowerShell

$ErrorActionPreference = 'Stop'
$ROOT  = Split-Path $MyInvocation.MyCommand.Path -Parent
$FECHA = Get-Date -Format 'yyyyMMdd'
$DEST  = Join-Path $ROOT "SISTEMA DE VACACIONES_v$FECHA.zip"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Sistema de Vacaciones USIL - Crear ZIP de distribucion" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Carpetas y archivos a EXCLUIR del ZIP
$EXCLUIR = @(
    '.venv',
    '__pycache__',
    '.git',
    '.claude',
    'node_modules',
    'dist_electron',
    'DATAS',
    'DATA SENSIBLE',
    'RESPALDOS',
    'alertas_cola',
    'PIPELINE\.venv',
    'PIPELINE\__pycache__',
    'PIPELINE\SALIDAS',
    'PIPELINE\SUBIDOS',
    'PIPELINE\DATOS ACTUALIZADOS',
    'PIPELINE\bot_adryan\logs'
)

# Archivos a EXCLUIR (sensibles o runtime)
$EXCLUIR_ARCHIVOS = @(
    'pa_config.json',           # contiene secretos (SMTP, webhook)
    'confirmaciones_vacaciones.json',
    'log_envios.json',
    'run_sistema.log',
    'PIPELINE\estado_pipeline.json',
    'PIPELINE\bot_adryan\sesion_adryan.json',
    'PIPELINE\bot_adryan\config_bot.json',
    'PIPELINE\bot_adryan\estado_bot.json'
)

# Extensiones a excluir
$EXCLUIR_EXT = @('.pkl', '.log', '.pyc', '.pyo', '.xlsx', '.xls', '.xlsm')

Write-Host "[..] Recopilando archivos..." -ForegroundColor Yellow

$archivos = Get-ChildItem -Path $ROOT -Recurse -File | Where-Object {
    $ruta = $_.FullName

    # Excluir por carpeta
    $enCarpetaExcluida = $false
    foreach ($exc in $EXCLUIR) {
        $patron = [System.IO.Path]::Combine($ROOT, $exc)
        if ($ruta.StartsWith($patron, [System.StringComparison]::OrdinalIgnoreCase)) {
            $enCarpetaExcluida = $true
            break
        }
    }
    if ($enCarpetaExcluida) { return $false }

    # Excluir por nombre de archivo relativo
    $rel = $ruta.Substring($ROOT.Length + 1)
    foreach ($exc in $EXCLUIR_ARCHIVOS) {
        if ($rel -eq $exc -or $rel -ieq $exc) { return $false }
    }

    # Excluir por extension
    if ($EXCLUIR_EXT -contains $_.Extension.ToLower()) { return $false }

    return $true
}

Write-Host "[OK] $($archivos.Count) archivos incluidos." -ForegroundColor Green

# Eliminar ZIP anterior si existe
if (Test-Path $DEST) { Remove-Item $DEST -Force }

Write-Host "[..] Creando ZIP: $DEST" -ForegroundColor Yellow

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open($DEST, 'Create')

foreach ($archivo in $archivos) {
    $rel = $archivo.FullName.Substring($ROOT.Length + 1)
    $entrada = "SISTEMA DE VACACIONES\$rel"
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $archivo.FullName, $entrada, 'Optimal') | Out-Null
}

$zip.Dispose()

$tam = (Get-Item $DEST).Length / 1MB
Write-Host ""
Write-Host "[OK] ZIP creado exitosamente!" -ForegroundColor Green
Write-Host "     Archivo: $(Split-Path $DEST -Leaf)" -ForegroundColor White
Write-Host "     Tamano:  $([math]::Round($tam, 1)) MB" -ForegroundColor White
Write-Host ""
Write-Host "INSTRUCCIONES PARA EL USUARIO DESTINO:" -ForegroundColor Cyan
Write-Host "  1. Extraer el ZIP en cualquier carpeta" -ForegroundColor White
Write-Host "  2. Ejecutar INSTALAR_SISTEMA_VACACIONES.bat (una sola vez)" -ForegroundColor White
Write-Host "  3. Completar pa_config.json con correo y contrasena de aplicacion" -ForegroundColor White
Write-Host "  4. Copiar los archivos Excel de DATAS\ y DATA SENSIBLE\" -ForegroundColor White
Write-Host "  5. Doble clic en el acceso directo del Escritorio" -ForegroundColor White
Write-Host ""
Write-Host "NOTA: Los archivos Excel (DATAS\, DATA SENSIBLE\) NO se incluyen" -ForegroundColor DarkYellow
Write-Host "      en el ZIP por ser datos sensibles. Copiarlos manualmente." -ForegroundColor DarkYellow
Write-Host ""

Read-Host "Presiona Enter para cerrar"
