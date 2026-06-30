# Sistema de Vacaciones USIL

## Instalacion

1. Ejecuta INSTALAR_SISTEMA_VACACIONES.bat.
2. El instalador verificara Python 3.11+, creara .venv, instalara dependencias y dejara un acceso directo en el escritorio.

## Inicio diario

1. Ejecuta ABRIR_SISTEMA_VACACIONES.vbs o el acceso directo Sistema de Vacaciones USIL del escritorio.
2. El script levantara el servidor si no esta corriendo y abrira la aplicacion en el navegador.

## Empaquetado en .exe

1. Instala Inno Setup en una maquina de compilacion.
2. Abre SISTEMA_VACACIONES_USIL.iss y compila el proyecto.
3. El resultado saldra en la carpeta dist como un instalador .exe.

## Dependencias instaladas

- Flask
- Waitress
- Pandas
- Openpyxl

## Notas

- Si Python no esta instalado, el instalador intentara instalar Python 3.11 con winget.
- Si winget no existe, instala Python manualmente y vuelve a ejecutar el instalador.
- El sistema se sirve en http://127.0.0.1:5002.