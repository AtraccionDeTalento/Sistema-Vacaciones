# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Qué es

Sistema de Vacaciones USIL (People Analytics). Dashboard local de RRHH que cruza la tabla maestra de personal con el reporte de vacaciones, calcula el cumplimiento de la meta por colaborador / supervisor / Business Partner (HRBP), y dispara alertas por correo y Teams. Todo el texto del proyecto (código, UI, docs) está en español.

Corre **localmente** en `http://127.0.0.1:5002` sobre Waitress. No hay base de datos: toda la persistencia es en archivos (Excel `.xlsx`, JSON, y caché `pickle`).

## Comandos

```bash
# Arrancar el servidor (idempotente: si el puerto 5002 ya responde, solo abre el navegador)
run_sistema.bat
# o directo:
.venv/Scripts/python.exe servidor.py        # http://127.0.0.1:5002

# Instalación inicial (verifica Python 3.11+, crea .venv, instala deps, crea acceso directo)
INSTALAR_SISTEMA_VACACIONES.bat

# Dependencias del servidor: flask, waitress, pandas, openpyxl (ver requirements.txt)

# Vaciar la cola de correos pendientes vía Outlook de escritorio (COM, sin SMTP)
python enviar_cola_outlook.py               # envía alertas_cola/in/*.json
python enviar_cola_outlook.py --dry-run     # muestra qué enviaría sin enviar
```

No hay suite de tests ni linter configurados. Empaquetado a `.exe` vía Electron (`main.js` + `electron-builder`, carpeta de salida `dist_electron/win-unpacked/`; también existe un `.iss` de Inno Setup histórico).

## Despliegue e instalaciones — LEE ESTO antes de tocar una copia fuera de esta carpeta

**Única fuente de verdad**: este repo (`C:\SISTEMA DE VACACIONES` en la máquina de desarrollo), rama `main` en GitHub (`AtraccionDeTalento/Sistema-Vacaciones`). Todo cambio de código va aquí primero, se commitea y se **pushea** — un commit local sin pushear no existe para efectos de las 5 PCs cliente ni de ningún script de actualización (todos bajan de GitHub, no de disco local).

**Nunca clones este repo por segunda vez en la misma máquina** (ej. a `Documents/` o `Downloads/`). Ya pasó una vez: un clon en `Documents\SISTEMA DE VACACIONES` divergió con 3 commits propios nunca subidos — uno de ellos un fix de seguridad real (contraseña en texto plano) que quedó sin aplicar por semanas mientras el repo "oficial" seguía expuesto en GitHub. Si necesitas otra copia de trabajo, usa `git worktree` o simplemente trabaja en este mismo folder.

**Las 5 PCs cliente corren un build empaquetado de Electron** (`dist_electron/win-unpacked/`), no este repo con `.git`. La copia "maestra" que se usa como plantilla para instalar/actualizar esas PCs vive en `C:\Users\EQUIPO\Documents\SISTEMA DE VACACIONES\dist_electron\win-unpacked` — tras cualquier fix real en este repo, sincronízala copiando los archivos de `_ARCHIVOS_ACTUALIZABLES` (ver abajo) y actualizando su `.version_commit`.

**Cómo se actualiza el código en una PC cliente** (de más a menos automático):
1. **Botón "⬇️ Actualizar sistema"** en la barra superior del dashboard (`/api/sistema/actualizar` en `servidor.py`) — baja los archivos de GitHub vía HTTPS, los verifica con SHA-256, y reinicia el proceso Python **en caliente** (spawn de un proceso nuevo que reintenta el bind al puerto 5002, luego `os._exit()` del viejo — **no uses `os.execv` para esto en Windows, no es confiable, se probó y deja el servidor muerto**). Es el mecanismo recomendado: funciona sin depender de Electron.
2. **Auto-actualizador de Electron** (`main.js` → `verificarActualizacion`) — corre solo al abrir la app. **Limitación real y permanente**: `main.js` queda empaquetado dentro de `resources/app.asar` al construir el `.exe`; aunque baje una copia nueva de sí mismo a disco, esa copia nunca se ejecuta (Electron sigue corriendo el `main.js` horneado en el instalador). Por eso `main.js` **no** está en la lista de archivos auto-actualizables — solo puede cambiar reconstruyendo y redistribuyendo el `.exe`. Sí actualiza correctamente todo lo demás (`servidor.py`, `index_vacaciones.html`, `assets/`, `PIPELINE/`).
3. **`.bat` manuales** (para cuando 1 y 2 no corrieron o hace falta forzar): `ACTUALIZAR_Y_ABRIR_TOTAL.bat` (punto de entrada recomendado, reemplaza abrir el `.exe` directo), `REPARAR_TOTAL.bat` (mata procesos viejos + fuerza + verifica por hash en memoria — usar cuando "ya se actualizó pero el problema sigue"), `FORZAR_ACTUALIZACION.bat`, `DIAGNOSTICO_ACTUALIZACION.bat`.

**Lista de archivos distribuidos** — vive duplicada en 3 lugares que deben mantenerse en sync manualmente (no hay una fuente única todavía): `main.js` (`ARCHIVOS_ACTUALIZABLES`), `servidor.py` (`_ARCHIVOS_ACTUALIZABLES`), y el array `$archivos`/`$files` dentro de cada `.bat` de arriba. Si agregas un archivo nuevo al proyecto que debe llegar a las PCs cliente, agrégalo a los 3.

**Diagnóstico multi-PC**: `DIAGNOSTICO_DATOS.bat` (compara hashes de Excel/código entre PCs con metas distintas) y `DIAGNOSTICO_CELDA_CORRUPTA.bat` (encuentra la fila exacta de `BASE GENERAL` que infla un avance sospechoso) — ambos hablan con `/api/diagnostico/*` en `servidor.py`.

### Pipeline de Excel (subproyecto independiente en `PIPELINE/`)

Proceso separado, con su **propio Python y dependencias** (necesita Excel instalado + `xlwings`/`pywin32`). Toma el reporte crudo de Adryan (`VACRptMotivo_*.xlsx` de Descargas) y actualiza el Excel objetivo "como un humano" por COM, preservando pivotes/gráficos/formato.

```bat
PIPELINE\ACTUALIZAR_VACACIONES.bat          # un clic
cd PIPELINE\motor & python pipeline.py      # normal
python pipeline.py --dry-run                # valida sin abrir Excel
python pipeline.py --oculto                 # Excel invisible
```

Toda la configuración del pipeline está en `PIPELINE/motor/config.json` (rutas, columnas, reglas) — no se edita código para reconfigurar. El servidor invoca el pipeline con el intérprete de `config.json → integracion_front.python_pipeline` (distinto del `.venv` del servidor). Solo se permite una corrida a la vez.

## Arquitectura

**Backend monolítico — `servidor.py` (~14.5k líneas, 66 rutas Flask).** Un solo archivo que es a la vez controlador HTTP, motor analítico (pandas) y dispatcher de notificaciones. Todo el trabajo nuevo de API ocurre aquí. Convención de nombres: helpers internos con prefijo `_`, rutas con `@app.route('/api/...')`.

**Frontend SPA vanilla JS** (sin frameworks):
- `index_vacaciones.html` — única página, servida en `/`.
- `assets/js/app_completo.js` — toda la lógica del dashboard.
- `assets/js/pipeline_vac.js` — panel de "Avance de meta" + botón "Actualizar vacaciones".
- `assets/css/styles.css`.

Flask se sirve con `static_folder='.'`, así que `index_vacaciones.html`, `assets/`, etc. se sirven directo desde la raíz.

### Datos (todo en archivos)

- `DATAS/PersonalMaestroReporte_*.xlsx` — tabla maestra de personal (universo completo, organigrama, cadena de mando). El sistema toma el archivo **más reciente** por patrón de nombre; los `_sanitizado` son los buenos.
- `DATA SENSIBLE/Reporte Vacaciones Objetivo_*.xlsx` — metas y días de vacaciones por trimestre. `VACACIONES_DATA_FILE` apunta al trimestre vigente (con fallback al `Reporte Vacaciones Objetivo*` más reciente). Este es **el archivo que lee el dashboard**; el pipeline publica aquí su salida.
- `DATAS/__cache__/*.pkl` — DataFrames serializados con `pickle`. La invalidación es por **mtime** del Excel: si tocas un Excel, la siguiente lectura recomputa. Si cambia la versión de Python, hay que borrar el caché. `_cargar_con_cache()` es el patrón central.
- `confirmaciones_vacaciones.json`, `log_envios.json` — estado liviano transaccional, protegido con `threading.Lock`.

Detección de columnas/hojas es **tolerante**: por nombre de cabecera normalizado (sin tildes/mayúsculas), no por posición (`_norm`, `_col`, `_detectar_header_row`, `_resolver_hoja_excel`). Excel abierto por el usuario se lee copiando primero a temp (`_safe_read_excel`) para evitar `PermissionError`.

### Notificaciones (3 canales)

1. **Outlook envío directo (COM)** — **canal principal del frontend**. `_enviar_correo_smtp` (nombre histórico: ya NO usa SMTP; todo el código SMTP real fue eliminado) envía el correo en silencio vía Outlook de escritorio (`mail.Send()`) desde la cuenta del usuario con sesión iniciada — desde el panel se percibe como envío directo. Requiere Outlook abierto. Usa `pythoncom.CoInitialize()` porque Waitress atiende cada request en un hilo distinto. Los flags `enviar_smtp`/`smtp_*` de API y UI conservan el nombre histórico pero significan "enviar vía Outlook COM". `enviar_cola_outlook.py --revisar` abre la cola como borradores (Display) en vez de auto-enviar.
2. **Cola "File Drop" → Power Automate** (respaldo, sin UI visible): escribe `{uuid}.json` en `alertas_cola/in/` (dentro de OneDrive). Power Automate, `enviar_cola_outlook.py` o el endpoint `/api/cola-pa/enviar-outlook-ahora` lo recogen y envían. Ver `_guardar_json_cola`, `_cola_pa_loop`, `_aplicar_rutas_cola`. SMTP AUTH y el flujo PA desde la cuenta de Talento están bloqueados por TI, por eso dejó de ser el canal principal. La mecánica de envío no se expone en la UI a propósito: debe percibirse como automatización.
3. **Teams webhook** (`_post_teams_webhook`, `_build_webhook_payload`).

Las alertas se agrupan por supervisor/HRBP y se compilan a HTML con `_build_html_jefe`. Hay modos de prueba (`teams_testing_mode`, `vacaciones_test_email`, allowlist/blocklist) para no spamear destinatarios reales.

### Configuración — `pa_config.json`

Config central del servidor (junto a `servidor.py`). Contiene **secretos** (contraseña SMTP, webhook de Teams) además del mapeo sección/área → supervisor, fechas límite de campaña, y flags de modo prueba. Se lee/escribe con `_buscar_pa_config` / `_guardar_pa_config` (patch parcial). No commitear cambios que expongan los secretos.

**Nunca hardcodees ni commitees credenciales reales** (contraseñas, tokens, cookies de sesión), ni en código "de respaldo"/fallback ni en archivos de texto sueltos tipo `CREDENTIALS.txt`. Ya pasó: la contraseña de Adryan vivió en texto plano en `PIPELINE/bot_adryan/CREDENTIALS.txt` (trackeado en git, pusheado a GitHub) y en un fallback hardcodeado en `guardar_password.py` durante semanas. Los secretos van en `pa_config.json` / `cred_adryan.bin` (cifrado DPAPI) / `config_bot.json` — **todos gitignored**. Antes de cualquier commit que toque `PIPELINE/bot_adryan/` o manejo de credenciales, correr `git grep` por contraseñas conocidas.

### Reglas de negocio clave (cuidado al tocar cálculos)

- **El "Avance de meta" (`_kpis_vacaciones()`) se calcula 100% en Python, nunca leyendo una celda de fórmula de Excel.** Antes leía `BASE GENERAL!Q1`/`T1` (SUMs sobre una fórmula `VLOOKUP` que arma `pipeline.py` por COM al escribir el Excel). Se demostró con casos reales que esa fórmula puede quedar mal recalculada tras una corrida del pipeline (Excel COM asíncrono, `RefreshAll` a medio terminar) y producir un avance corrupto (90%+) **incluso leyendo el Excel "en vivo" sin caché de por medio** — la celda simplemente queda guardada mal. Ahora `_meta_vac_data()` suma los días de forma independiente desde la hoja `base` (respetando `_CORTE_REGISTRO_VAC`) en la misma pasada que arma el universo "con meta", y expone el total amplio (incluye cesados/colegio, para calzar con el Excel crudo) como `dias_meta_bruto_total`/`dias_gozados_bruto_total`. **No vuelvas a leer `BASE GENERAL!Q1`/`T1`/`R_Cumplimiento!E9` como fuente de un KPI mostrado al usuario** — son celdas de fórmula, no fuente de verdad.
- **"Obligatorio" ≠ "tiene meta".** Una persona es de salida **obligatoria** solo si `Vac. Vencidas > 0` o `Vac. Pendiente > 0` (días legales). Los **truncos se pagan, no se gozan**, así que no obligan aunque haya meta. Esta regla está validada celda por celda contra el archivo original (0 diferencias). Ver `PIPELINE/README_PIPELINE.md §3` y `_build_analisis.py`.
- **Mapeo Área/Departamento → Business Partner** vive en `_bp_map.py` (derivado del organigrama de Talento y Cultura); devuelve un HRBP o `REVISAR`. `_build_analisis.py` es un script standalone que genera el workbook de auditoría `DATA SENSIBLE/ANALISIS_BP_*.xlsx`.
- Cadena de mando y equipos se resuelven recursivamente desde la maestra (`_obtener_cadena_mando_recursiva`, `_obtener_equipo_recursivo`).

## Documentación

`docs/` tiene material extenso en español: `ARQUITECTURA.md` (diagramas mermaid del flujo de datos y notificaciones), `API_REFERENCE.md`, `MANUAL_TECNICO.md`, `MANUAL_USUARIO.md`, `INVENTARIO_TECNICO.md`. `PIPELINE/README_PIPELINE.md` documenta el subproyecto de Excel en detalle.
