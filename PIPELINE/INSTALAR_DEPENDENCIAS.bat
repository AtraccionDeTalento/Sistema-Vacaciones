@echo off
chcp 65001 >nul
title Vacaciones USIL - Instalar dependencias
cd /d "%~dp0motor"

rem Detecta el Python disponible en esta maquina (la ruta cambia de PC en PC).
set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not exist "%PYEXE%" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not exist "%PYEXE%" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if not exist "%PYEXE%" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
if not exist "%PYEXE%" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python39\python.exe"
if not exist "%PYEXE%" for /f "delims=" %%P in ('where python 2^>nul') do if not exist "%PYEXE%" if exist "%%P" set "PYEXE=%%P"
if not exist "%PYEXE%" (
    echo No se encontro ningun Python instalado en esta maquina.
    echo Instala Python 3.10+ desde https://www.python.org/downloads/ y vuelve a correr esto.
    pause
    exit /b 1
)

echo Instalando dependencias del pipeline con:
echo   %PYEXE%
echo.
rem Un pip viejo (el que trae Python de fabrica) no encuentra wheels precompilados
rem de paquetes recientes y trata de compilarlos desde codigo fuente (falla sin
rem Visual Studio instalado). Actualizarlo primero evita ese fallo.
"%PYEXE%" -m pip install --upgrade --no-user pip
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
echo Descargando el navegador Chromium para Playwright (bot de Adryan)...
"%PYEXE%" -m playwright install chromium
echo.
echo Listo. Si no hubo errores arriba, el pipeline y el bot de Adryan ya pueden ejecutarse.
echo Recuerda que ademas necesitas Microsoft Excel de escritorio instalado en esta PC
echo (el pipeline lo automatiza por COM; eso no se puede empaquetar en un instalador).
pause
