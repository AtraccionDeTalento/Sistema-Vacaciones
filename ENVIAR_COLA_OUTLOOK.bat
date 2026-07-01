@echo off
title Enviar Cola de Correos - Sistema Vacaciones
color 0B
cd /d "%~dp0"

echo ========================================================
echo    ENVIAR CORREOS PENDIENTES DE ALERTAS (Outlook COM)
echo ========================================================
echo.
echo Procesara todos los archivos JSON en alertas_cola\in\
echo y los enviara via Outlook de escritorio.
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] No se encontro .venv\Scripts\python.exe
    echo         Ejecuta primero INSTALAR_SISTEMA_VACACIONES.bat
    pause
    exit /b 1
)

if not exist "enviar_cola_outlook.py" (
    echo [ERROR] No se encontro enviar_cola_outlook.py
    pause
    exit /b 1
)

echo Iniciando envio...
echo.
.venv\Scripts\python.exe enviar_cola_outlook.py
echo.
echo ========================================================
pause
