# Portafolio de Proyecto: Sistema de Vacaciones Inteligente

## Resumen Ejecutivo

El **Sistema de Gestión y Alerta de Vacaciones** es una solución corporativa desarrollada para el área de Talento y Cultura (People Analytics) de la Universidad San Ignacio de Loyola (USIL). Este software resuelve el problema crítico de la planificación deficiente de los períodos vacacionales, lo cual generaba pasivos laborales y afectaba la salud organizacional.

Mediante la automatización de flujos de datos complejos y el uso inteligente de plataformas corporativas (Teams y Office 365), la aplicación empodera a los HR Business Partners (HRBP) para hacer seguimiento proactivo a las jefaturas en toda la institución.

## 1. Problema de Negocio

Previo al desarrollo del sistema, la gestión vacacional padecía las siguientes deficiencias:
1. **Descentralización y opacidad:** Los datos habitaban en pesados archivos de nómina difíciles de leer y cruzar manualmente.
2. **Cuellos de botella de comunicación:** Las alertas de vencimiento de vacaciones hacia los líderes requerían escribir decenas de correos individualizados.
3. **Falta de métricas en tiempo real:** No existía un lugar unificado donde el equipo directivo pudiera ver la tasa de cumplimiento trimestral.

## 2. La Solución Entregada

Se construyó una **aplicación web de análisis de datos** que opera en los equipos locales del equipo de Talento y Cultura.

### Características Principales:
* **Dashboard Analítico (KPIs):** Segmenta automáticamente al personal en categorías de acción (Elegibles, Con Saldo Crítico, etc.).
* **Generación de Jerarquías Dinámicas:** Reconstruye el organigrama de la empresa on-the-fly basándose en el reporte maestro para agrupar a cada empleado con su jefe directo.
* **Motor de Plantillas de Comunicación:** Genera resúmenes personalizados por jefatura, detallando exactamente qué miembros de su equipo están en riesgo de perder o acumular vacaciones indebidamente.
* **Integración Multicanal Automatizada:** Con un solo clic, el sistema se conecta a **Power Automate** para lanzar correos corporativos auténticos (sorteando problemas de spam) y emite notificaciones silenciosas a **Microsoft Teams**.

## 3. Valor Agregado e Impacto

1. **Ahorro de Tiempo Extraordinario:** Un proceso de revisión, cruce de Excels y redacción de correos que antes tomaba días, ahora se ejecuta masivamente en un asistente de 3 clics en menos de un minuto.
2. **Cumplimiento Legal y Pasivos:** Garantiza el control estricto sobre los saldos vacacionales, evitando multas o pagos indemnizatorios obligatorios por vacaciones truncas o vencidas según la ley peruana.
3. **Poka-Yoke Informático:** El sistema cuenta con mecanismos de visualización previa, "modo de prueba", y cancelación en diferido (cola de 60 segundos) que garantizan cero correos enviados por error humano.

## 4. Tecnologías Involucradas

* Procesamiento pesado de datos: **Python (Pandas)**
* Backend RESTful: **Flask & Waitress**
* Frontend Ágil: **JavaScript ES6, Single Page Application**
* Integraciones Corporativas: **Microsoft Power Automate (RPA), Microsoft Teams Webhooks**

Este proyecto evidencia la madurez en la creación de puentes entre ecosistemas analíticos pesados (Data Analytics) y soluciones de software de usuario final orientadas a procesos de Recursos Humanos.
