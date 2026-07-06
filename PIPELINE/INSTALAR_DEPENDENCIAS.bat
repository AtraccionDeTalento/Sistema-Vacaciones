@echo off
chcp 65001 >nul
title Vacaciones USIL - Instalar dependencias
cd /d "%~dp0motor"

set "PYEXE=C:\Users\jlopezp\AppData\Local\Programs\Python\Python313\python.exe"
if not exist "%PYEXE%" set "PYEXE=python"

echo Instalando dependencias del pipeline con:
echo   %PYEXE%
echo.
rem --no-user fuerza instalacion global (no en el perfil del usuario): asi
rem pywin32/xlwings no dependen de como cada PC resuelva su AppData, que es
rem justo lo que causo el fallo "engines.active" en una instalacion previa.
"%PYEXE%" -m pip install --upgrade --no-user -r requirements.txt
echo.
echo Registrando pywin32 (necesario para que xlwings encuentre Excel por COM)...
for %%I in ("%PYEXE%") do set "PYDIR=%%~dpI"
if exist "%PYDIR%Scripts\pywin32_postinstall.py" (
    "%PYEXE%" "%PYDIR%Scripts\pywin32_postinstall.py" -install
) else (
    echo   (no se encontro pywin32_postinstall.py; se omite este paso)
)
echo.
echo Listo. Si no hubo errores arriba, el pipeline ya puede ejecutarse.
pause
