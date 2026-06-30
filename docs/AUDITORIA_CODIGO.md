# Auditoría de Código y Deuda Técnica

Este documento presenta una revisión crítica de la calidad del código, las deficiencias arquitectónicas (Deuda Técnica) y las recomendaciones estratégicas para asegurar el mantenimiento futuro del Sistema de Vacaciones.

## 1. Riesgos de Arquitectura Identificados

### 1.1 El Monolito `servidor.py` (God Object)
* **Hallazgo:** Toda la lógica del backend reside en un solo archivo de casi 350 KB y más de 13,000 líneas. Se mezclan responsabilidades de controladores web (Flask), servicios de dominio, manejo de archivos OS, consumo de webhooks y procesamiento masivo con Pandas.
* **Riesgo:** Altísima probabilidad de causar regresiones. Modificar una regla de validación de Excel podría romper inadvertidamente el flujo de webhooks debido a posibles side-effects en variables globales. El código es indescifrable para un desarrollador nuevo en poco tiempo.
* **Severidad:** CRÍTICA.

### 1.2 Persistencia Basada en Caché y Archivos Temporales
* **Hallazgo:** Se utiliza la librería `pickle` para almacenar el resultado de Pandas en disco en la carpeta `__cache__`.
* **Riesgo:** `pickle` no es seguro contra inyección de código si el directorio se ve comprometido. Adicionalmente, las estructuras de `pickle` se rompen fácilmente ante actualizaciones menores de las versiones de la librería de `pandas` o Python.
* **Severidad:** ALTA.

### 1.3 Condición de Carrera en Archivos `.json` (State Management)
* **Hallazgo:** Archivos como `confirmaciones_vacaciones.json` y `log_envios.json` actúan como bases de datos mutables. Se utilizan `threading.Lock()` para manejarlos.
* **Riesgo:** Si el sistema escalara para funcionar bajo múltiples instancias de `Waitress` en distintos procesos (workers > 1), el bloqueo a nivel de hilos fallaría irremediablemente, corrompiendo el archivo JSON.
* **Severidad:** MEDIA (Mitigado temporalmente por el uso estricto en entorno local con 1 solo worker principal).

### 1.4 Dependencia Estrecha a Formatación de Excel
* **Hallazgo:** El procesamiento depende de que se mantenga el nombre de columnas específicas, de que no hayan filas "combinadas" erróneas, y de nombres de archivo como `Reporte Vacaciones Objetivo*.xlsx`.
* **Riesgo:** Las estructuras de datos manuales fallan. Si la persona que extrae el reporte de Oracle/SAP cambia el orden o nombre de las cabeceras, el sistema lanzará un `KeyError` interno, volviéndose inoperativo sin ofrecer feedback claro en el front.
* **Severidad:** ALTA.

## 2. Recomendaciones y Plan de Refactorización

Para transformar esta prueba de concepto en un software de nivel corporativo, se deben seguir estos pasos:

### Fase 1: Modularización (Prioridad 1)
Desacoplar el "God Object" (`servidor.py`) usando el patrón MVC o Hexagonal:
1. Crear carpeta `routes/` (o blueprints de Flask) para contener las definiciones de endpoints.
2. Crear carpeta `services/` para aislar la lógica de pandas (procesamiento).
3. Crear carpeta `integrations/` para las llamadas de Teams y escritura hacia Power Automate.

### Fase 2: Inserción de Base de Datos (Prioridad 2)
Reemplazar la arquitectura *File-based* por un motor relacional embebido:
* Utilizar **SQLite** y un ORM como SQLAlchemy.
* Beneficio: Transacciones seguras, integridad de datos, queries rápidas (sin necesidad de cargar en RAM gigantescos DataFrames cada vez), y adiós a la inestabilidad del `pickle`.

### Fase 3: Hardening del Frontend (Prioridad 3)
* Dividir el gigantesco archivo `app_completo.js` en componentes modulares ES6 o transicionar la SPA a un empaquetador ligero como `Vite`.
* Añadir validaciones robustas y manejo de errores visible cuando la API de Flask retorne un código `500`.

## 3. Conclusión de Auditoría
El sistema cumple operativamente sus objetivos de negocio de manera altamente efectiva y creativa (aprovechando Power Automate como pasarela de correos), pero su base de código actual es frágil. Se requiere un esfuerzo de estabilización arquitectónica antes de delegar el sistema a ingenieros junior u otro departamento de IT.
