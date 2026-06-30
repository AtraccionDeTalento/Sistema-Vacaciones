# API Reference - Sistema de Vacaciones

Esta documentación lista los principales endpoints RESTful expuestos por `servidor.py` y consumidos por la SPA frontend.

Todas las respuestas son en formato `application/json`.

## Endpoints de Inicialización y Estado

### `GET /api/init`
* **Descripción:** Endpoint inicial de carga rápida (ping) para verificar que el servidor local está operativo.
* **Respuesta (200 OK):**
  ```json
  { "status": "ok", "version": "1.0", "config": {...} }
  ```

### `GET /api/resumen`
* **Descripción:** Retorna las estadísticas agregadas principales (KPIs) resultantes del cruce del maestro de empleados con las vacaciones.
* **Parámetros GET:** Ninguno.
* **Respuesta (200 OK):**
  ```json
  {
    "kpis": {
      "total": 1200,
      "con_vacaciones": 450,
      "con_saldo": 300,
      "sin_vacaciones": 450
    },
    "detalle_alerta": [...]
  }
  ```

## Endpoints de Directorio y Estructura Organizacional

### `GET /api/jefes_equipo_arbol`
* **Descripción:** Devuelve la jerarquía organizacional procesada. Contiene la lista estructurada de todos los supervisores que pueden ser contactados.
* **Respuesta (200 OK):**
  ```json
  {
    "status": "ok",
    "data": [
      {
        "supervisor_id": "0001",
        "nombre": "Juan Perez",
        "gerencia": "Tecnología",
        "equipo_directo_count": 5
      }
    ]
  }
  ```

### `GET /api/personas_autocomplete`
* **Descripción:** Devuelve el directorio plano indexado para el campo de autocompletado en el formulario de "Envío Individual".

## Endpoints de Transaccionalidad

### `POST /api/enviar_a_supervisor`
* **Descripción:** Desencadena el procesamiento y generación de las alertas de correo para un líder de equipo. Deposita el payload en la cola de Power Automate.
* **Cuerpo de la Petición (JSON):**
  ```json
  {
    "id_supervisor": "0001",
    "plantilla": "urgente",
    "modo_prueba": false,
    "mensaje_personalizado": "..."
  }
  ```
* **Respuesta (200 OK):**
  ```json
  { "status": "ok", "enqueued": true, "transaction_id": "uuid-..." }
  ```

### `GET /api/cola-pa/ultima-pendiente`
* **Descripción:** Efectúa un polling (Short Polling) para saber si existen mensajes pendientes de envío retenidos por el temporizador de retraso.
* **Uso:** Determina si el botón "Cancelar" debe mostrarse en la UI.

### `POST /api/cola-pa/cancelar-todos`
* **Descripción:** Purga el directorio `alertas_cola/pendientes` hacia la carpeta `cancelados`, previniendo que los correos sean emitidos.

## Endpoints de Utilidad

### `GET /api/diagnostico_datos`
* **Descripción:** Devuelve el "Health Check" de los archivos de Excel requeridos. Avisa al usuario si falta el archivo Maestro o de Datos Sensibles, y muestra su última fecha de modificación.
