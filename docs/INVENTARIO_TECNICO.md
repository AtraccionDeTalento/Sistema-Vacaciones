# Inventario Técnico - Sistema de Vacaciones USIL

## 1. Información General
* **Proyecto:** Sistema de Vacaciones USIL (People Analytics)
* **Objetivo:** Automatizar la gestión, análisis y notificación (seguimiento) del goce vacacional de los colaboradores, enviando alertas personalizadas a las jefaturas de Talento y Cultura a través de múltiples canales (Correo, Teams, Power Automate).
* **Stack Tecnológico:** Python (Flask, Waitress, Pandas), HTML5, Vanilla CSS, Vanilla JavaScript.

## 2. Estructura de Carpetas

```text
SISTEMA DE VACACIONES/
├── .venv/                      # Entorno virtual de Python
├── assets/                     # Archivos estáticos del frontend
│   ├── css/
│   │   └── styles.css          # Hoja de estilos principal (Vanilla CSS)
│   └── js/
│       └── app_completo.js     # Lógica del frontend (Vanilla JS)
├── DATA SENSIBLE/              # Carpeta para reportes confidenciales
│   └── Reporte Vacaciones...   # Excels con la información vacacional y objetivos
├── DATAS/                      # Datos maestros e intermedios
│   ├── __cache__/              # Archivos temporales serializados (Pickle)
│   └── PersonalMaestro...      # Excel con la data maestra de colaboradores
├── docs/                       # Documentación técnica y de usuario
├── static/                     # Archivos estáticos adicionales
└── [Archivos Raíz]
```

## 3. Dependencias (requirements.txt)
* **flask (>=3.0.0):** Framework backend para la API y servidor web.
* **waitress (>=3.0.0):** Servidor WSGI para despliegue en producción (Windows).
* **pandas (>=2.2.3):** Para la manipulación, limpieza y análisis cruzado de los archivos Excel.
* **openpyxl (>=3.1.5):** Motor para leer y procesar archivos `.xlsx`.

## 4. Archivos y Componentes Principales
### 4.1. Archivos Raíz
* `servidor.py`: Archivo monolítico (~350KB) que contiene TODA la lógica del backend (Rutas Flask, procesamiento Pandas, integraciones, lógica de negocio).
* `index_vacaciones.html`: SPA (Single Page Application) que sirve como interfaz gráfica del sistema.
* `confirmaciones_vacaciones.json` y `log_envios.json`: Archivos que actúan como "base de datos" local para registrar transacciones y estados.
* `run_sistema.bat`: Script de inicio que arranca el servidor en el puerto 5002 usando Waitress/Python.
* `SISTEMA_VACACIONES_USIL.iss`, `INSTALAR_SISTEMA_VACACIONES.bat/ps1`: Scripts para la instalación local y empaquetado del software en las máquinas de los usuarios.

### 4.2. Módulos y Flujo Principal de Ejecución (`servidor.py`)
El backend no está dividido en módulos. Todas las funciones están en `servidor.py`:
* **Procesamiento de Datos:** Funciones como `_cargar_maestro_universo`, `_fusionar_universo_maestro`, `cargar_datos`. Utilizan Pandas para cruzar la data de "Personal Maestro" con los reportes de "Vacaciones".
* **Caché:** Funciones como `_cargar_con_cache` usan `pickle` para evitar la sobrecarga de leer los Excels en cada reinicio.
* **API REST:** Rutas como `api_init`, `api_resumen`, `api_jefes_equipo_arbol` sirven datos al frontend.
* **Integraciones y Notificaciones:** 
  * Power Automate: Deposita JSONs en una carpeta local/OneDrive (`alertas_cola/in`).
  * Teams: Envía webhooks (`_enviar_teams_webhook_supervisor`).
  * Correo: Usa SMTP (`_enviar_correo_smtp`).

## 5. Deuda Técnica y Riesgos de Mantenimiento

1. **Arquitectura Monolítica Crítica:** `servidor.py` tiene más de 13,000 líneas de código. Esto lo hace altamente susceptible a errores en cadena, dificulta el testeo, y genera un riesgo altísimo de mantenibilidad (God Object Anti-pattern).
2. **Dependencia de Archivos Locales (Sin Base de Datos Real):** Se usan `.json` para logs y `.pkl` para caché. Esto puede causar problemas de concurrencia y corrupción de datos si dos usuarios intentan operar simultáneamente.
3. **Hardcoding de Rutas y Nombres:** La lógica de lectura depende de nombres de archivo específicos o patrones (`Reporte Vacaciones Objetivo*.xlsx`). Si el proveedor de datos cambia el formato de nombre, el sistema fallará.
4. **Acoplamiento Fuerte con Excel:** Todo el flujo depende de la estructura exacta de las columnas de los Excels. Si la matriz añade/elimina columnas, el backend lanzará excepciones de Pandas.
5. **Caché Frágil:** El uso de `pickle` para cachear DataFrames puede generar incompatibilidades si se actualiza la versión de Pandas o Python.

## 6. Puntos de Entrada
* **Web/Usuario:** `index_vacaciones.html` (Servido por Flask o directamente).
* **Backend:** Ejecución de `run_sistema.bat` -> `servidor.py` -> Inicia servidor en `127.0.0.1:5002`.
* **Cola de Procesamiento:** Directorio `alertas_cola/in` donde Power Automate vigila los archivos `.json` que genera el sistema para automatizaciones externas.
