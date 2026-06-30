# Pipeline de Vacaciones USIL

Automatiza el proceso manual que hace César: toma el reporte crudo de **Adryan**
(`VACRptMotivo_*.xlsx`), lo limpia, lo vuelca a la estructura objetivo, **actualiza la
base y las tablas dinámicas**, recalcula el cumplimiento y guarda una **copia fechada**.
"Toca el Excel como un humano" usando Excel por COM (xlwings), por lo que **conserva
pivotes, enlaces externos, gráficos y formato**.

---

## 1. Uso rápido

### Opción A — Un clic (.bat)
1. Descarga de Adryan el reporte (queda en `Descargas` como `VACRptMotivo_<fecha>_<hora>.xlsx`).
2. Doble clic en **`ACTUALIZAR_VACACIONES.bat`**.
3. Al terminar, el archivo actualizado queda en **`SALIDAS\`** con fecha y hora.

### Opción B — Botón dentro de Excel (macro VBA)
1. Excel → `Alt+F11` → Archivo → Importar → `motor\Lanzador_Excel.bas`.
2. Pestaña *Programador* → Insertar → Botón → asigna la macro `ActualizarVacaciones`.
3. Un clic en el botón corre exactamente lo mismo que el `.bat`.

### Línea de comandos (avanzado)
```bat
cd motor
python pipeline.py                 :: corrida normal (Excel visible segun config)
python pipeline.py --oculto        :: Excel invisible
python pipeline.py --no-cerrar     :: deja Excel abierto al terminar
python pipeline.py --dry-run       :: solo prepara/valida datos, NO abre Excel
python pipeline.py --crudo "RUTA"  :: fuerza un crudo especifico
```

---

## 2. Qué hace, paso a paso

```
ADRYAN  ->  VACRptMotivo_*.xlsx  (crudo, todo texto, con metadatos y fila "SIN GRUPO")
   |
   |  1. Toma el crudo MAS RECIENTE de Descargas
   |  2. Detecta la fila de cabecera ("Matrícula") y conserva solo filas con
   |     matricula de 10 digitos  ->  elimina metadatos y filas de grupo
   |  3. Convierte tipos:  Dias y Mes Pago -> numero ;  Fecha Ingreso/Inicio/Termino
   |     -> fecha real ;  Matricula/Año Pago/Periodo -> texto (conserva ceros)
   v
hoja 'base'   (se limpia y se vuelca; preserva la columna OBSERVACION)
   |
   |  4. Reajusta la fuente del pivote y RefreshAll (tabla dinamica DB_Vac)
   |  5. Calcula la tabla espejo L:S en Python (Suma de Dias por Matricula x Mes Pago)
   v
hoja 'DB_Vac_07_05_2026'  (pivote visual + espejo L:S + columna de reconciliacion)
   |
   |  6. Rellena formulas:
   |       Meta?       = IF(Objetivo>0, "Sí","No")
   |       Obligatorio = IF(O(Vencidas>0 ; Pendiente>0), "Sí","No")   <-- regla verificada
   |       Registradas = VLOOKUP al espejo (Total general)
   |       Meta% = Registradas/Objetivo ;  Dias Restantes = Total - Registradas
   |       Totales de la fila 1 (SUM dinamico)
   v
hoja 'BASE GENERAL'  (1 fila por empleado)
   |
   |  7. RefreshAll otra vez  ->  R_Cumplimiento toma la nueva BASE GENERAL
   v
hoja 'R_Cumplimiento'  (avance por HRBP; excluye al colegio via filtro del pivote)
   |
   v
SALIDAS\Reporte ... - actualizado_AAAAMMDD_HHMM.xlsx   +   log en logs\
```

**Reconciliación automática:** al final lista las personas que **registraron vacaciones
pero no están en BASE GENERAL** (nuevos ingresos posteriores a la tabla maestra). Es el
mismo cruce manual que hace César ("se está escapando alguien"). Salen en el log y en la
columna de reconciliación del espejo (marca `FALTA EN BG`).

---

## 3. La regla de "Obligatorio" (importante)

> **Obligatorio = "Sí"** solo si la persona tiene **Vac. Vencidas > 0** o **Vac. Pendiente > 0**.
> Si solo tiene **truncos** (o nada), es **"No"**, aunque tenga meta.

Motivo: los truncos se **pagan**, no se gozan, así que no obligan a salir de vacaciones.
Esta regla se validó **celda por celda contra las 1,778 filas** del archivo original:
**0 diferencias**. Antes la columna estaba escrita a mano; ahora es una fórmula que se
recalcula sola.

---

## 4. Estructura de carpetas

```
PIPELINE\
├─ ACTUALIZAR_VACACIONES.bat      <- lanzador de un clic
├─ INSTALAR_DEPENDENCIAS.bat      <- instala librerias (solo la 1a vez / PC nueva)
├─ README_PIPELINE.md             <- este archivo
├─ Reporte ... Segundo Trimestre 2026.xlsx   <- PLANTILLA objetivo (no se toca)
├─ motor\
│   ├─ pipeline.py                <- motor principal (orquesta Excel por COM)
│   ├─ vac_lib.py                 <- lectura/limpieza del crudo + conversiones
│   ├─ config.json                <- TODA la configuracion (rutas, columnas, reglas)
│   ├─ requirements.txt           <- dependencias
│   └─ Lanzador_Excel.bas         <- macro VBA (boton dentro de Excel)
├─ SALIDAS\                       <- copias fechadas resultantes (output)
├─ RESPALDOS\                     <- (reservada)
└─ logs\                          <- un log por corrida
```

---

## 5. Configuración (`motor\config.json`)

Todo se ajusta sin tocar el código. Lo más útil:

| Clave | Para qué |
|---|---|
| `entrada.carpetas_descarga` | Dónde buscar el crudo. Puedes poner varias carpetas. |
| `entrada.patron_crudo` | Patrón del archivo (`VACRptMotivo_*.xlsx`). |
| `salida.objetivo_plantilla` | El archivo objetivo base. |
| `salida.usar_salida_mas_reciente_como_fuente` | `true` = cada corrida parte de la salida anterior (acumula OBSERVACION y ajustes manuales). |
| `salida.prefijo_salida` | Nombre de las copias fechadas. |
| `excel.visible` | `true` = ves Excel trabajando; `false` = silencioso. |
| `esquema_base` | Las 17 columnas, su orden y su tipo (texto/fecha/entero). |
| `base_general.regla_obligatorio` | La fórmula de Obligatorio (editable). |

> **Mapeo por nombre, no por posición:** las columnas del crudo se ubican por su **nombre
> de cabecera** (sin importar acentos, mayúsculas ni el orden). Si Adryan cambia el orden
> de columnas, el pipeline sigue funcionando.

---

## 6. Mantenimiento y notas

- **Crecimiento de la base (límite 2000):** las fórmulas y la fuente del pivote vienen a
  2000 filas. El pipeline **auto-extiende** los rangos según los registros. Aun si el
  pivote visual quedara corto, los **números de cumplimiento son correctos** porque el
  espejo se calcula en Python desde todos los datos.
- **Nuevo trimestre:** cambia `salida.objetivo_plantilla` por la plantilla del nuevo
  trimestre y, si cambian los meses, el pipeline los detecta solos (usa los `Mes Pago`
  presentes en el crudo).
- **Enlaces externos `[1]` y `[2]`** (tabla maestra y reporte de truncas): se abren sin
  actualizar, conservando sus valores. No se rompen.
- **El original nunca se modifica**: siempre se trabaja sobre una copia fechada.
- **Dependencias:** Python con `openpyxl`, `pandas`, `xlwings`, `pywin32`, y Excel
  instalado. Si faltan, corre `INSTALAR_DEPENDENCIAS.bat`.

## 7. Problemas comunes

| Síntoma | Causa / solución |
|---|---|
| "No se encontró crudo" | No hay `VACRptMotivo_*.xlsx` en Descargas, o cambió el patrón. |
| "ModuleNotFoundError: xlwings" | Corre `INSTALAR_DEPENDENCIAS.bat` o ajusta `PYEXE` en el `.bat`. |
| Excel queda abierto/colgado | Cierra procesos `EXCEL.EXE` huérfanos y reintenta. |
| Avance se ve raro | Revisa el log en `logs\` y la lista de "FALTA EN BG" (nuevos ingresos). |

---

## 8. Integración con el dashboard (botón + KPIs en tiempo real)

El dashboard de vacaciones (`servidor.py`, puerto 5002) tiene dos botones arriba a la derecha:

- **📊 Avance de meta** — abre un panel con el avance hacia la meta: global (49% toda la base
  y 48% sin colegio) y **una barra por cada Business Partner** (HRBP), ordenadas del más
  atrasado al más avanzado. Lee `/api/vacaciones/avance` (BASE GENERAL W1 + R_Cumplimiento).
- **🔄 Actualizar vacaciones** — corre el pipeline. Al pulsarlo:

1. Dispara el pipeline en el servidor (toma el último `VACRptMotivo` de Descargas).
2. Muestra el **progreso en vivo** (barra + pasos: leyendo crudo → volcando base →
   refrescando dinámica → recalculando → publicando).
3. Al terminar, **publica** la salida al archivo que lee el dashboard y **anima el avance
   de meta** de antes → después (con el delta en puntos y los días registrados nuevos).

### Cómo funciona por dentro
- El pipeline, además de la copia fechada en `SALIDAS\`, **publica** una copia en
  `DATA SENSIBLE\Reporte Vacaciones Objetivo_Segundo Trimestre 2026.xlsx` (el archivo exacto
  que lee el front), respaldando antes el archivo vivo en `RESPALDOS\`. Como el servidor
  cachea por *mtime*, los KPIs se refrescan solos en la siguiente lectura.
- El pipeline también escribe `PIPELINE\estado_pipeline.json` con los KPIs (meta, registrado,
  avance, % cumplimiento, nuevos ingresos). El servidor lo lee al instante (sin tocar OneDrive).

### Endpoints añadidos a servidor.py
| Método | Ruta | Para qué |
|---|---|---|
| POST | `/api/vacaciones/pipeline/run` | Inicia la actualización (devuelve `job_id`). |
| GET  | `/api/vacaciones/pipeline/estado/<job_id>` | Progreso en vivo + KPIs antes/después. |
| GET  | `/api/vacaciones/kpis` | KPIs actuales (para prellenar el panel). |

El servidor corre el pipeline con el Python definido en
`config.json → integracion_front.python_pipeline` (el que tiene xlwings), aunque el servidor
mismo use su `.venv`. Solo se permite **una actualización a la vez**.

### Archivos del front
- `index_vacaciones.html` — botón + panel modal + estilos (añadidos).
- `assets/js/pipeline_vac.js` — lógica del panel (nuevo, autocontenido).
- `servidor.py` — bloque "INTEGRACION CON EL PIPELINE" antes del arranque.

> Para **desactivar** la publicación al archivo vivo: en `config.json` pon
> `integracion_front.publicar_por_defecto = false` (o corre el pipeline con `--no-publicar`).
