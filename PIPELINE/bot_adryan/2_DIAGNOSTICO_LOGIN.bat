@echo off
chcp 65001 >nul
title Diagnostico de login Adryan (un solo intento)
cd /d "%~dp0"

rem Detecta el Python disponible en esta maquina (la ruta cambia de PC en PC).
set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
if not exist "%PYEXE%" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not exist "%PYEXE%" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
if not exist "%PYEXE%" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
if not exist "%PYEXE%" set "PYEXE=%LOCALAPPDATA%\Programs\Python\Python39\python.exe"
if not exist "%PYEXE%" for /f "delims=" %%P in ('where python 2^>nul') do if not exist "%PYEXE%" if exist "%%P" set "PYEXE=%%P"
if not exist "%PYEXE%" (
    echo No se encontro un Python instalado con Playwright en esta maquina.
    echo Instala Python y corre: pip install playwright ^&^& playwright install chromium
    pause
    exit /b 1
)

echo ============================================================
echo   DIAGNOSTICO DE LOGIN A ADRYAN - UN SOLO INTENTO
echo   (no hace reintentos: si Adryan esta limitando por volumen
echo    de intentos, esto no lo empeora)
echo ============================================================
echo.

"%PYEXE%" diagnostico_login.py

echo.
pause
