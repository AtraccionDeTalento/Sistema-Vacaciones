@echo off
setlocal enabledelayedexpansion

set ROOT=%~dp0
set LOG=%ROOT%run_sistema.log
set PORT=5002
set PYW=%ROOT%.venv\Scripts\pythonw.exe
set SRV=%ROOT%servidor.py

:: Crear/actualizar acceso directo en el Escritorio para procesar la cola manualmente
powershell -NoProfile -Command ^
  "$s=(New-Object -COM WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\Enviar Cola Vacaciones.lnk');" ^
  "$s.TargetPath='%ROOT%ENVIAR_COLA_OUTLOOK.bat';" ^
  "$s.WorkingDirectory='%ROOT%';" ^
  "$s.Description='Enviar correos de alertas pendientes via Outlook';" ^
  "$s.IconLocation='%SystemRoot%\System32\SHELL32.dll,13';" ^
  "$s.Save()" 2>nul

:: Si el puerto ya responde, verificar que el proceso tenga el servidor.py
:: actual cargado en memoria (via /api/health). Si el archivo en disco es mas
:: nuevo que lo que el proceso cargo al arrancar, el codigo esta desactualizado
:: (Flask/Waitress no recargan solos): se mata el proceso viejo y se sigue
:: abajo para levantar uno nuevo, en vez de abrir el navegador sobre codigo viejo.
powershell -NoProfile -Command ^
  "$ErrorActionPreference='Stop';" ^
  "$portOk = (Test-NetConnection -ComputerName 127.0.0.1 -Port %PORT% -WarningAction SilentlyContinue).TcpTestSucceeded;" ^
  "if (-not $portOk) { exit 1 }" ^
  "$stale = $false;" ^
  "try {" ^
  "  $h = Invoke-RestMethod -Uri ('http://127.0.0.1:{0}/api/health' -f %PORT%) -TimeoutSec 3;" ^
  "  if ($h.codigo_desactualizado) { $stale = $true }" ^
  "} catch {" ^
  "  $code = $null;" ^
  "  try { $code = [int]$_.Exception.Response.StatusCode } catch {}" ^
  "  if ($code -eq 404) { $stale = $true }" ^
  "  else { Start-Process ('http://127.0.0.1:{0}' -f %PORT%); exit 0 }" ^
  "}" ^
  "if ($stale) {" ^
  "  Write-Host 'Codigo desactualizado detectado (o /api/health inexistente), reiniciando servidor...';" ^
  "  Get-NetTCPConnection -LocalPort %PORT% -State Listen -ErrorAction SilentlyContinue | ForEach-Object { try { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue } catch {} };" ^
  "  Start-Sleep -Seconds 1;" ^
  "  exit 1" ^
  "} else {" ^
  "  Start-Process ('http://127.0.0.1:{0}' -f %PORT%); exit 0" ^
  "}"
if not errorlevel 1 exit /b 0

if not exist "%PYW%" (
  where pythonw >nul 2>&1
  if errorlevel 1 (
    msg * [SISTEMA DE VACACIONES] No se encontro pythonw ni venv local.
    exit /b 1
  )
  set PYW=pythonw
)

cd /d "%ROOT%"
echo [%DATE% %TIME%] Iniciando servidor... > "%LOG%"
start "" /min "%PYW%" "%SRV%" 1>>"%LOG%" 2>&1

powershell -NoProfile -Command "$deadline=(Get-Date).AddSeconds(30); $ok=$false; while((Get-Date) -lt $deadline) { try { $tcp=[System.Net.Sockets.TcpClient]::new(); $tcp.Connect('127.0.0.1',%PORT%); $tcp.Close(); $ok=$true; break } catch {} Start-Sleep -Milliseconds 400 }; if ($ok) { Start-Sleep -Seconds 1; Start-Process 'http://127.0.0.1:%PORT%' } else { msg * '[SISTEMA DE VACACIONES] El servidor no arranco en 30s. Revisar run_sistema.log'; Start-Process notepad '%LOG%' }"

endlocal
