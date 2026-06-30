@echo off
chcp 65001 >nul
title Vacaciones USIL - Instalar dependencias
cd /d "%~dp0motor"

set "PYEXE=C:\Users\jlopezp\AppData\Local\Programs\Python\Python313\python.exe"
if not exist "%PYEXE%" set "PYEXE=python"

echo Instalando dependencias del pipeline con:
echo   %PYEXE%
echo.
"%PYEXE%" -m pip install --upgrade -r requirements.txt
echo.
echo Listo. Si no hubo errores arriba, el pipeline ya puede ejecutarse.
pause
