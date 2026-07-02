@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
title Instalador - Sistema de Vacaciones USIL

set ROOT=%~dp0
set VENV=%ROOT%.venv
set PIP=%VENV%\Scripts\pip.exe
set PY=%VENV%\Scripts\python.exe
set REQ=%ROOT%requirements.txt

echo ============================================================
echo  SISTEMA DE VACACIONES USIL - Instalador
echo ============================================================
echo.

:: ── 1. Verificar Python 3.11+ ────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo [!] Python no encontrado. Intentando instalar con winget...
    winget install -e --id Python.Python.3.11 --silent
    if errorlevel 1 (
        echo [ERROR] No se pudo instalar Python automaticamente.
        echo         Instala Python 3.11+ desde https://www.python.org/downloads/
        echo         y vuelve a ejecutar este instalador.
        pause
        exit /b 1
    )
    echo [OK] Python instalado.
)

for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo [OK] Python %PYVER% encontrado.

:: ── 2. Crear entorno virtual ─────────────────────────────────
if not exist "%VENV%\Scripts\python.exe" (
    echo [..] Creando entorno virtual .venv ...
    python -m venv "%VENV%"
    if errorlevel 1 (
        echo [ERROR] No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
    echo [OK] Entorno virtual creado.
) else (
    echo [OK] Entorno virtual ya existe.
)

:: ── 3. Instalar dependencias ──────────────────────────────────
echo [..] Instalando dependencias (puede tardar unos minutos)...
"%PIP%" install --upgrade pip --quiet
"%PIP%" install -r "%REQ%" --quiet
if errorlevel 1 (
    echo [ERROR] Fallo al instalar dependencias. Revisa requirements.txt.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas.

:: ── 4. Configurar pa_config.json si no existe ─────────────────
if not exist "%ROOT%pa_config.json" (
    if exist "%ROOT%pa_config.plantilla.json" (
        copy "%ROOT%pa_config.plantilla.json" "%ROOT%pa_config.json" >nul
        echo [OK] pa_config.json creado desde plantilla.
        echo [!!] IMPORTANTE: Abre pa_config.json y configura tu correo y contrasena de aplicacion.
    ) else (
        echo [WARN] No se encontro pa_config.plantilla.json. Crea pa_config.json manualmente.
    )
) else (
    echo [OK] pa_config.json ya existe.
)

:: ── 5. Crear acceso directo en el Escritorio ─────────────────
set SHORTCUT=%USERPROFILE%\Desktop\Sistema de Vacaciones USIL.lnk
powershell -NoProfile -Command ^
  "$s=(New-Object -COM WScript.Shell).CreateShortcut('%SHORTCUT%');" ^
  "$s.TargetPath='%ROOT%run_sistema.bat';" ^
  "$s.WorkingDirectory='%ROOT%';" ^
  "$s.Description='Sistema de Vacaciones USIL';" ^
  "$s.IconLocation='%ROOT%icon.png,0';" ^
  "$s.Save()" 2>nul
echo [OK] Acceso directo creado en el Escritorio.

echo.
echo ============================================================
echo  Instalacion completada con exito!
echo  Doble clic en "Sistema de Vacaciones USIL" del Escritorio
echo  para abrir la aplicacion.
echo ============================================================
echo.
pause
