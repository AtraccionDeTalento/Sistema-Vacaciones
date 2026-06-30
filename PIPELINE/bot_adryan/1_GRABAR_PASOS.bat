@echo off
chcp 65001 >nul
title Grabar pasos de Adryan (calibracion del bot)
cd /d "%~dp0"

set "PYEXE=C:\Users\jlopezp\AppData\Local\Programs\Python\Python313\python.exe"

echo ============================================================
echo   GRABADOR DE PASOS - BOT ADRYAN
echo ------------------------------------------------------------
echo   Se abrira una ventana de Chrome controlada por Playwright
echo   y una ventana "Playwright Inspector" con el codigo.
echo.
echo   QUE HACER (una sola vez):
echo     1. En la barra de direcciones del Chrome que se abre,
echo        escribe la URL de Adryan y entra.
echo     2. Loguearte (usuario y contrasena).
echo     3. Navega hasta el reporte de vacaciones.
echo     4. Scrollea / toca los botones tal cual lo haces a mano.
echo     5. DESCARGA el reporte (VACRptMotivo_*.xlsx).
echo     6. Cierra la ventana de Chrome cuando termines.
echo.
echo   El codigo grabado quedara en:  pasos_grabados.py
echo   Pasame ese archivo y yo armo el bot definitivo.
echo ============================================================
echo.
pause

"%PYEXE%" -m playwright codegen --channel chrome --target python -o pasos_grabados.py

echo.
echo ============================================================
echo   Grabacion guardada en  pasos_grabados.py
echo ============================================================
pause
