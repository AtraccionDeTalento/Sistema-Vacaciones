@echo off
title Sistema de Vacaciones USIL - App
color 0B
echo ========================================================
echo       INICIANDO SISTEMA COMO APLICACION DE ESCRITORIO
echo ========================================================
echo.
cd /d "%~dp0"

echo [1/2] Verificando motor grafico (Electron)...
if not exist "node_modules\electron" (
    echo Instalando dependencias de Electron por primera vez...
    call npm install
)

echo.
echo [2/2] Abriendo App...
call npm start
exit
