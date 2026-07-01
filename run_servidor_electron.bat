@echo off
setlocal enabledelayedexpansion
set ROOT=%~dp0
set PORT=5002

:: Este script arranca el servidor flask (servidor.py) pero NO abre el navegador.
:: De eso se encarga Electron.

echo Verificando dependencias de Python...
if exist "%ROOT%\.venv\Scripts\python.exe" (
    echo Usando entorno virtual .venv
    "%ROOT%\.venv\Scripts\python.exe" "%ROOT%\servidor.py"
) else (
    echo No hay .venv, intentando usar python global...
    pip install -r "%ROOT%\requirements.txt" --disable-pip-version-check --quiet
    python "%ROOT%\servidor.py"
)
