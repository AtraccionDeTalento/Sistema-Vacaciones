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

powershell -NoProfile -Command "if ((Test-NetConnection -ComputerName 127.0.0.1 -Port %PORT% -WarningAction SilentlyContinue).TcpTestSucceeded) { Start-Process 'http://127.0.0.1:%PORT%'; exit 0 } else { exit 1 }"
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
