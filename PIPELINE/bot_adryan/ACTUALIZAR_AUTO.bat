@echo off
chcp 65001 >nul
title Bot Adryan - Descargar y Actualizar Vacaciones
cd /d "%~dp0"

set "PYEXE=C:\Users\jlopezp\AppData\Local\Programs\Python\Python313\python.exe"

rem  --forzar  = corre el pipeline aunque no cambien los datos
rem  --solo-bot = solo descarga, no procesa
"%PYEXE%" actualizar_todo.py %*

exit /b %ERRORLEVEL%
