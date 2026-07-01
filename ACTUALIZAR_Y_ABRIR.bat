@echo off
title Actualizador del Sistema de Vacaciones
color 0B

echo ========================================================
echo       ACTUALIZANDO SISTEMA DE VACACIONES (NUBE)
echo ========================================================
echo.
cd /d "%~dp0"

echo [1/3] Conectando con GitHub para descargar cambios...
git pull origin main
if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo.
    echo ERROR: Hubo un problema al descargar la actualizacion.
    echo Asegurate de tener internet y permisos.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [2/3] Verificando e instalando nuevas dependencias...
if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
    pip install -r requirements.txt --disable-pip-version-check --quiet
) else (
    echo [ADVERTENCIA] No se encontro el entorno virtual .venv. Saltando dependencias.
)

echo.
echo [3/3] Actualizacion exitosa. Iniciando el sistema...
timeout /t 3 /nobreak >nul

:: Ejecutar el script original de apertura
start "" "ABRIR_SISTEMA_VACACIONES.vbs"
exit
