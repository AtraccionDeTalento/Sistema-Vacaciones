@echo off
title Actualizador del Sistema de Vacaciones
color 0B

echo ========================================================
echo       ACTUALIZANDO SISTEMA DE VACACIONES (GITHUB)
echo ========================================================
echo.
cd /d "%~dp0"

echo [1/4] Descargando cambios de GitHub...
git pull origin main
if %ERRORLEVEL% NEQ 0 (
    color 0C
    echo.
    echo ERROR: No se pudo descargar la actualizacion. Revisa internet y permisos.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [2/4] Actualizando dependencias Python...
if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
    pip install -r requirements.txt --disable-pip-version-check --quiet
) else (
    echo [ADVERTENCIA] No se encontro .venv. Saltando dependencias.
)

echo.
echo [3/4] Sincronizando app de Business Partners (OneDrive)...
set DIST=%~dp0dist_electron\win-unpacked
if exist "%DIST%\servidor.py" (
    copy /Y servidor.py "%DIST%\servidor.py" >nul
    copy /Y index_vacaciones.html "%DIST%\index_vacaciones.html" >nul
    copy /Y _bp_map.py "%DIST%\_bp_map.py" >nul
    copy /Y assets\js\app_completo.js "%DIST%\assets\js\app_completo.js" >nul
    copy /Y assets\js\pipeline_vac.js "%DIST%\assets\js\pipeline_vac.js" >nul
    copy /Y assets\css\styles.css "%DIST%\assets\css\styles.css" >nul
    copy /Y PIPELINE\bot_adryan\bot_adryan.py "%DIST%\PIPELINE\bot_adryan\bot_adryan.py" >nul
    copy /Y PIPELINE\bot_adryan\bot_maestro.py "%DIST%\PIPELINE\bot_adryan\bot_maestro.py" >nul
    echo    OK: archivos sincronizados. Los BPs recibiran cambios al reiniciar su app.
) else (
    echo    ADVERTENCIA: No se encontro dist_electron\win-unpacked, saltando sync de BPs.
)

echo.
echo [4/4] Actualizacion exitosa. Iniciando el sistema...
timeout /t 2 /nobreak >nul
call run_sistema.bat
