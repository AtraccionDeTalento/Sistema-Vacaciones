@echo off
title Diagnostico de Actualizacion - Sistema de Vacaciones USIL
color 0B

if not exist "%~dp0DIAGNOSTICO_ACTUALIZACION.ps1" (
    echo [ERROR] Falta el archivo DIAGNOSTICO_ACTUALIZACION.ps1
    echo         Debe estar en la MISMA carpeta que este .bat.
    echo.
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0DIAGNOSTICO_ACTUALIZACION.ps1"

echo.
pause
