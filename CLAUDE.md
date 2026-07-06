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

No hay suite de tests ni linter configurados. Empaquetado a `.exe` vía Inno Setup (`SISTEMA_VACACIONES_USIL.iss`).

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

### Reglas de negocio clave (cuidado al tocar cálculos)

- **"Obligatorio" ≠ "tiene meta".** Una persona es de salida **obligatoria** solo si `Vac. Vencidas > 0` o `Vac. Pendiente > 0` (días legales). Los **truncos se pagan, no se gozan**, así que no obligan aunque haya meta. Esta regla está validada celda por celda contra el archivo original (0 diferencias). Ver `PIPELINE/README_PIPELINE.md §3` y `_build_analisis.py`.
- **Mapeo Área/Departamento → Business Partner** vive en `_bp_map.py` (derivado del organigrama de Talento y Cultura); devuelve un HRBP o `REVISAR`. `_build_analisis.py` es un script standalone que genera el workbook de auditoría `DATA SENSIBLE/ANALISIS_BP_*.xlsx`.
- Cadena de mando y equipos se resuelven recursivamente desde la maestra (`_obtener_cadena_mando_recursiva`, `_obtener_equipo_recursivo`).

## Documentación

`docs/` tiene material extenso en español: `ARQUITECTURA.md` (diagramas mermaid del flujo de datos y notificaciones), `API_REFERENCE.md`, `MANUAL_TECNICO.md`, `MANUAL_USUARIO.md`, `INVENTARIO_TECNICO.md`. `PIPELINE/README_PIPELINE.md` documenta el subproyecto de Excel en detalle.
