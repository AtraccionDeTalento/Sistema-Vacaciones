# Manual Técnico para Desarrolladores

Este manual provee instrucciones para el mantenimiento del código fuente, depuración de errores y despliegue del Sistema de Vacaciones.

## 1. Configuración del Entorno de Desarrollo

### Instalación Manual
Si no se utiliza el instalador compilado, siga estos pasos:
1. Clonar el repositorio.
2. Crear un entorno virtual: `python -m venv .venv`
3. Activar entorno: `.venv\Scripts\activate`
4. Instalar librerías: `pip install -r requirements.txt`

### Variables de Configuración
El sistema utiliza un archivo físico opcional llamado `pa_config.json` en la raíz (generado dinámicamente o editable manualmente). Las variables críticas incluyen:
* `alertas_cola_dir`: Ruta absoluta al directorio de OneDrive para Power Automate.
* `teams_webhook_url`: Endpoint de Teams para alertas.
* `smtp_email` / `smtp_password`: Credenciales (en caso de fallback SMTP).
* `session_secret`: Llave criptográfica para las sesiones web.

## 2. Anatomía de `servidor.py`

Debido a que `servidor.py` es un archivo monolítico, está lógicamente seccionado mediante comentarios en bloque.

### 2.1 Sección: Manejo de Archivos Excel
Funciones clave: `_safe_read_excel`, `_cargar_con_cache`, `cargar_datos`.
* **Cuidado:** `_safe_read_excel` copia temporalmente el archivo a la carpeta `%TEMP%` antes de leerlo con Pandas. Esto soluciona el `PermissionError` clásico de Windows cuando el usuario tiene el Excel abierto.

### 2.2 Sección: Lógica de Negocio y Jerarquías
Funciones clave: `_obtener_equipo_recursivo`, `api_jefes_equipo_arbol`.
* **Recursividad:** El sistema calcula las jerarquías orgánicas iterando sobre el DataFrame maestro usando el campo de "ID Supervisor" para armar el árbol de dependencias hacia abajo.

### 2.3 Sección: Integraciones (Queue)
Funciones clave: `_aplicar_rutas_cola`, `_liberar_todo_pa_ahora`.
* **File Drop:** El sistema define varias carpetas: `/in` (donde Power Automate lee), `/pendientes` (donde se retrasan los envíos por N segundos), y `/archivados`. El hilo principal en segundo plano mueve archivos de `/pendientes` a `/in` cuando se vence el timer de retraso (usualmente 60 segundos, permitiendo cancelar envíos en el front).

## 3. Anatomía del Frontend (`app_completo.js`)

Es un script de más de 100KB que controla todo el DOM de la aplicación SPA.
* **Manejo de Estado:** Utiliza variables globales al inicio del archivo (ej. `window.STATE`).
* **Wizard:** La lógica de envío masivo está manejada por la función `wizardGo(step)`. Esta función valida el DOM de cada vista, oculta paneles (usando `.visible`) y muestra el progreso.
* **Autocompletado:** El buscador de jefaturas individuales consulta localmente una lista de objetos parseada durante el `/api/init`.

## 4. Troubleshooting (Resolución de Problemas)

| Síntoma | Causa Probable | Solución |
|---------|----------------|----------|
| No inicia / Puerto ocupado | Proceso `pythonw.exe` colgado en background. | Matar el proceso desde el Administrador de tareas o ejecutar `taskkill /F /IM pythonw.exe`. |
| Errores 500 al arrancar | Los excels cambiaron de nombre o columnas. | Revisar `run_sistema.log`. Asegurar que los nombres en la carpeta `DATA SENSIBLE` coincidan con el patrón esperado. Eliminar contenido de `DATAS/__cache__`. |
| Power Automate no envía correos | OneDrive local no está sincronizando. | Verificar que el ícono de la nube de OneDrive en la barra de tareas esté conectado y subiendo los archivos depositados en `alertas_cola/in`. |

## 5. Procedimiento de Depuración
Para auditar la ejecución de Pandas, buscar el decorador `@medir_tiempo` en `servidor.py` y aplicarlo a la función que se sospecha es lenta. El tiempo de ejecución se imprimirá en consola.
