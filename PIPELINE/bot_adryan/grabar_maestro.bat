@echo off
echo ============================================
echo   GRABADOR DE PASOS - MAESTRO DE PERSONAL
echo ============================================
echo.
echo Se abrira Adryan en Chrome con el grabador de Playwright.
echo Navega hasta descargar el Maestro de Personal.
echo Playwright grabara cada clic y lo guardara en:
echo   pasos_maestro.py
echo.
echo Cuando termines, cierra el navegador.
echo ============================================
echo.

cd /d "%~dp0"

"C:\Users\jlopezp\AppData\Local\Programs\Python\Python313\python.exe" -m playwright codegen --target python -o "%~dp0pasos_maestro.py" --channel chrome https://adryancloudusil.sapia.com.pe/

echo.
echo ============================================
echo   GRABACION COMPLETADA
echo   Archivo generado: pasos_maestro.py
echo ============================================
pause
