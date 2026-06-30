@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title Sistema de Vacaciones USIL - Actualizar

cd /d "%~dp0motor"

rem --- Elegir el Python que tenga las dependencias (xlwings, openpyxl, pywin32) ---
set "PYEXE=C:\Users\jlopezp\AppData\Local\Programs\Python\Python313\python.exe"
if not exist "%PYEXE%" (
  where py >nul 2>nul && set "PYEXE=py -3"
)
if not exist "%PYEXE%" if not defined PYEXE set "PYEXE=python"

echo ============================================================
echo   ACTUALIZANDO SISTEMA DE VACACIONES USIL
echo   Toma el ultimo VACRptMotivo_*.xlsx de Descargas,
echo   actualiza la base + tablas dinamicas y guarda una
echo   copia fechada en la carpeta SALIDAS.
echo ============================================================
echo.

"%PYEXE%" pipeline.py %*
set "RC=%ERRORLEVEL%"

echo.
echo ============================================================
if "%RC%"=="0" (
  echo   LISTO. Revisa la carpeta SALIDAS y el log en  logs\.
) else (
  echo   TERMINO CON CODIGO %RC%. Revisa el log en  logs\  para el detalle.
)
echo ============================================================
echo.
pause
endlocal
