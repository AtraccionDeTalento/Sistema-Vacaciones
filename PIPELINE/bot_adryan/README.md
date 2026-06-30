# Bot Adryan — descarga automática del crudo de vacaciones

Automatiza el paso manual de **entrar a Adryan, loguearse, fijar fechas, buscar y descargar**
el `VACRptMotivo_*.xlsx`. Usa **Playwright** controlando tu **Chrome** instalado.

```
[cada hora]  ->  bot_adryan.py  ->  descarga VACRptMotivo_*.xlsx a Descargas
             ->  actualizar_todo.py decide si cambio  ->  corre pipeline.py  ->  dashboard se refresca
```

## Archivos

| Archivo | Qué es |
|---|---|
| `bot_adryan.py` | El bot. Reusa sesión, navega al reporte, fija fechas (pickadate por JS), busca y descarga. |
| `actualizar_todo.py` | Orquestador: corre el bot y, **solo si los datos cambiaron**, lanza el pipeline. |
| `ACTUALIZAR_AUTO.bat` | Lanzador para la tarea programada (acepta `--forzar`, `--solo-bot`). |
| `guardar_password.py` | Guarda/lee la contraseña **cifrada con DPAPI** (`cred_adryan.bin`). |
| `config_bot.json` | Toda la configuración (URL, usuario, fechas, headless, etc.). |
| `1_GRABAR_PASOS.bat` | Re-grabar los pasos si Adryan cambia (genera `pasos_grabados.py`). |
| `cred_adryan.bin` | Contraseña cifrada (solo tu usuario de Windows la descifra). **No compartir.** |
| `sesion_adryan.json` | Sesión/cookies guardadas. **No compartir.** |
| `logs\` | Un log por día del bot y del orquestador; capturas `error_*.png` si algo falla. |

## Uso manual

```bat
:: descargar + actualizar (solo procesa si cambio)
ACTUALIZAR_AUTO.bat

:: forzar el pipeline aunque no cambien los datos
ACTUALIZAR_AUTO.bat --forzar

:: solo descargar, sin procesar
ACTUALIZAR_AUTO.bat --solo-bot

:: ver el bot trabajando (no headless)
python bot_adryan.py --visible
```

## Automático (cada hora)

Tarea programada de Windows: **"Vacaciones USIL - Bot Adryan"**, modo *Solo interactivo*
(corre solo cuando estás logueado, porque el pipeline abre Excel). Para verla/cambiarla:

```bat
schtasks /Query  /TN "Vacaciones USIL - Bot Adryan" /FO LIST
schtasks /Change /TN "Vacaciones USIL - Bot Adryan" /SC HOURLY /MO 2   :: cada 2 horas
schtasks /Delete /TN "Vacaciones USIL - Bot Adryan" /F                  :: quitarla
```
O por interfaz: **Programador de tareas** → Biblioteca → "Vacaciones USIL - Bot Adryan".

## Desde el dashboard

Botón **🔄 Actualizar vacaciones** → **⤓ Descargar de Adryan y actualizar**: el servidor
corre el bot y luego el pipeline, mostrando el progreso en vivo.

## Si Adryan cambia (se rompe el bot)

1. Mira la captura del error en `logs\error_*.png`.
2. Si cambiaron botones/campos: corre `1_GRABAR_PASOS.bat`, rehaz los pasos a mano,
   y compara `pasos_grabados.py` con los selectores de `bot_adryan.py`.
3. Si cambió la contraseña: corre `python guardar_password.py`.
4. Si cambió el trimestre/fechas: edita `fecha_inicio` y `fecha_termino` en `config_bot.json`.

## Notas de seguridad

- La contraseña **no** se guarda en texto plano: va cifrada con DPAPI (solo tu cuenta de
  Windows, en esta PC, la descifra).
- `cred_adryan.bin` y `sesion_adryan.json` son sensibles. Como esta carpeta está en OneDrive,
  no compartas el enlace de esta carpeta con nadie.
