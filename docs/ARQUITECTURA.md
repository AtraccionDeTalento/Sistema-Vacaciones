# Arquitectura del Sistema de Vacaciones

Este documento detalla la arquitectura de software del Sistema de Vacaciones, orientada a explicar los componentes, patrones de diseño y flujo de datos de extremo a extremo.

## 1. Visión General de la Arquitectura

El sistema utiliza un patrón arquitectónico monolítico de tipo **Cliente-Servidor (SPA + REST API)**, con persistencia basada enteramente en el sistema de archivos (File-based Persistence) y un mecanismo de mensajería asíncrona mediante un "File Drop" (Cola de archivos) para integraciones externas.

```mermaid
graph TD
    subgraph Frontend [Capa de Presentación - Cliente]
        UI[index_vacaciones.html]
        JS[app_completo.js]
        UI -->|Interacciones| JS
    end

    subgraph Backend [Capa de Servicios - Servidor Local]
        Flask[API REST Flask]
        Waitress[Waitress WSGI]
        Pandas[Pandas Data Engine]
        
        Waitress --> Flask
        Flask <--> Pandas
    end

    subgraph FileSystem [Capa de Datos y Persistencia]
        ExcelMaestro[(DATAS/Maestro.xlsx)]
        ExcelVacaciones[(DATA SENSIBLE/Vacaciones.xlsx)]
        PickleCache[(__cache__/*.pkl)]
        JSONLogs[(log_envios.json)]
    end

    subgraph Integraciones [Capa de Integración Externa]
        Teams[Teams Webhook API]
        OneDrive[(alertas_cola/in)]
        PowerAutomate[Microsoft Power Automate]
    end

    JS <-->|HTTP/REST| Waitress
    Pandas <-->|Lectura/Escritura| ExcelMaestro
    Pandas <-->|Lectura/Escritura| ExcelVacaciones
    Pandas <-->|Deserialización| PickleCache
    Flask <-->|Append| JSONLogs
    Flask -->|POST| Teams
    Flask -->|Genera .json| OneDrive
    OneDrive -->|Trigger| PowerAutomate
```

## 2. Descripción de Componentes

### 2.1 Capa de Presentación
La interfaz gráfica es una Single Page Application (SPA). En lugar de usar frameworks como React o Angular, emplea Vanilla JS con manipulación directa del DOM. El estado de la aplicación reside en el cliente durante la sesión.

### 2.2 Capa de Procesamiento (Backend)
Contenida íntegramente en `servidor.py` (~350KB). Realiza tres funciones fundamentales:
1. **Controlador HTTP:** Define las rutas usando los decoradores de Flask.
2. **Motor Analítico:** Utiliza `pandas` para ingerir archivos masivos de Excel, realizar joins (cruce entre Maestro y Vacaciones) y calcular los estados (Elegibles, Con Saldo, etc.).
3. **Dispatcher de Notificaciones:** Compila plantillas HTML y orquesta la salida de mensajes por SMTP nativo, Teams Webhooks y la cola de Power Automate.

### 2.3 Capa de Datos (File-based)
Al no existir un RDBMS (como PostgreSQL o SQL Server), todo el estado reside en el disco:
* **Lectura Pesada:** Archivos `.xlsx`. Para optimizar arranques, el sistema usa `pickle` para serializar en binario los DataFrames resultantes.
* **Transacciones Livianas:** Operaciones ACID-like simuladas mediante archivos `.json` (`confirmaciones_vacaciones.json`) utilizando bloqueos de hilos (`threading.Lock`) para evitar colisiones concurrentes.

## 3. Diagramas de Secuencia

### 3.1 Flujo de Ingesta y Caché

```mermaid
sequenceDiagram
    participant User as Usuario
    participant Server as servidor.py
    participant Disk as Sistema de Archivos

    User->>Server: Inicia Servidor (run_sistema.bat)
    Server->>Disk: Verifica `DATAS/__cache__/*.pkl`
    alt Caché Inválido o Mtime Alterado
        Server->>Disk: Lee Excels (Maestro y Vacaciones)
        Disk-->>Server: Retorna Datos (Lento)
        Server->>Server: Computa DataFrames (Pandas)
        Server->>Disk: Escribe nuevos .pkl
    else Caché Válido
        Disk-->>Server: Retorna binario deserializado (Rápido)
    end
    Server-->>User: Servidor Listo en :5002
```

### 3.2 Flujo de Envío de Notificaciones

```mermaid
sequenceDiagram
    participant JS as app_completo.js
    participant API as Flask API
    participant Logs as log_envios.json
    participant Queue as alertas_cola/in
    participant PA as Power Automate

    JS->>API: POST /api/enviar_a_supervisor (Payload)
    API->>API: Compila plantilla HTML/JSON
    API->>Logs: Append registro de transacción
    API->>Queue: Escribe archivo {uuid}.json
    API-->>JS: 200 OK (Envío Programado)
    
    Note over Queue, PA: Proceso Asíncrono
    PA->>Queue: File Drop Trigger detecta .json
    PA->>PA: Parsea JSON y extrae destinatarios
    PA->>UsuarioFinal: Envía correo corporativo Office 365
    PA->>Queue: Mueve a /archivados
```

## 4. Decisiones de Diseño Críticas
1. **Desacoplamiento asíncrono vía "File Drop":** Se optó por escribir `.json` en un directorio de OneDrive (`alertas_cola/in`) en lugar de integrar el Graph API de Microsoft directamente. Esto simplifica enormemente la autenticación, delegando la seguridad de Office 365 a Power Automate.
2. **Caché en Pickle:** Reduce el tiempo de arranque de minutos a segundos, pero requiere limpieza de caché si la versión de Python cambia.
