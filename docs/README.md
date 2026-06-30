# Sistema de Vacaciones USIL

Plataforma automatizada diseñada para la gestión, análisis cruzado y notificación del estado vacacional de los colaboradores en la organización. El sistema permite al área de Talento y Cultura enviar alertas proactivas a las jefaturas utilizando múltiples canales de comunicación (Correo, Teams, Power Automate).

## 🚀 Inicio Rápido

### Requisitos Previos
* **Sistema Operativo:** Windows 10/11 (Optimizado para entorno corporativo).
* **Python:** Python 3.10+ instalado.
* **Datos:** Archivos Excel de *Personal Maestro* y *Reporte Vacaciones Objetivo* alojados en los directorios correspondientes.

### Instalación
1. Ejecutar el archivo `INSTALAR_SISTEMA_VACACIONES.bat` o el instalador `.iss`.
2. El instalador creará un entorno virtual en la carpeta `.venv` y descargará las dependencias (`requirements.txt`).
3. El sistema configurará automáticamente la red local para exponer el puerto `5002`.

### Arranque
Ejecute el archivo `run_sistema.bat`. Este script validará el entorno, levantará el servidor WSGI (Waitress) en segundo plano y abrirá automáticamente el navegador por defecto en `http://127.0.0.1:5002`.

## 📂 Estructura del Proyecto

```text
SISTEMA DE VACACIONES/
├── .venv/                      # Entorno virtual aislado de Python
├── assets/                     # Recursos estáticos de la interfaz
│   ├── css/styles.css          # Estilos Vanilla CSS
│   └── js/app_completo.js      # Lógica transaccional frontend
├── DATA SENSIBLE/              # Origen de datos confidenciales vacacionales
├── DATAS/                      # Base maestra y almacenamiento de caché
├── docs/                       # Documentación técnica corporativa
├── index_vacaciones.html       # Interfaz principal (SPA)
├── servidor.py                 # Core monolítico del backend (Flask/Pandas)
├── requirements.txt            # Declaración de dependencias
├── confirmaciones_vacaciones.json # Base de datos transaccional de estados
└── log_envios.json             # Historial de auditoría de comunicaciones
```

## 🛠 Stack Tecnológico
* **Backend:** Python, Flask, Waitress.
* **Procesamiento de Datos:** Pandas, Openpyxl, Pickle.
* **Frontend:** HTML5, CSS3 (Vanilla), JavaScript ES6.
* **Integración:** Webhooks (Teams), Interfaz de Cola (Power Automate).

## 📄 Licencia y Propiedad
Este código es propiedad exclusiva de **Universidad San Ignacio de Loyola (USIL)** - Área de People Analytics. No se autoriza su distribución externa.
