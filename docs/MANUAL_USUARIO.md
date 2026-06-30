# Manual de Usuario - Sistema de Vacaciones

Bienvenido al Sistema de Gestión de Vacaciones. Esta herramienta centraliza el análisis del goce vacacional de la organización y facilita el envío de notificaciones proactivas a los líderes de equipo.

## 1. Acceso al Sistema

1. Busque el acceso directo **"Abrir Sistema de Vacaciones"** en su escritorio.
2. Si el sistema no está en ejecución, esto lanzará el servidor en segundo plano. Espere unos 5 segundos.
3. Se abrirá automáticamente su navegador web predeterminado (Google Chrome, Edge, etc.) en la dirección local del aplicativo.

## 2. Tablero de Control (Dashboard)

En la pantalla principal observará cuatro grandes tarjetas de resumen (KPIs):
* **Total Colaboradores:** La nómina activa completa.
* **Con Vacaciones Trimestre:** Personal que ya tiene programados o gozados sus días libres.
* **Con Saldo Vacacional:** Empleados que tienen un remanente acumulado.
* **Sin Vacaciones Trimestre:** Empleados que están en alerta por falta de planificación.

> **Tip:** Hacer clic en cualquiera de estas tarjetas abrirá un modal (ventana emergente) con la lista detallada de las personas que componen ese número.

## 3. Envío Masivo (Wizard de 3 Pasos)

La herramienta principal es el panel **"Envío de Mensajes a Jefaturas"**.

### Paso 1: ¿A quiénes enviar?
* Utilice los filtros desplegables para acotar la lista (Por HRBP, Gerencia o Área).
* Haga clic en los botones "Activar/Desactivar" para marcar a los líderes que desea notificar. Puede usar el botón de "✔ Todos" si filtró por un área en particular.
* Haga clic en "Siguiente".

### Paso 2: ¿Qué mensaje?
* Seleccione la plantilla de comunicación adecuada al escenario. Las opciones pueden incluir "Recordatorios estándar", "Alertas urgentes por vencimiento" o "Felicitaciones".
* En la parte inferior, verá una **Vista Previa** de cómo le llegará el correo al líder de equipo.
* Si desea alterar el mensaje, use la plantilla "Personalizado" o edite el campo directamente en pantalla.

### Paso 3: Confirmar y Enviar
* Verifique el recuento total de destinatarios.
* Es obligatorio **marcar la casilla de confirmación** ("Confirmo que revisé..."). Esto habilita el botón de enviar.
* Al presionar "🚀 Enviar a X jefes", el sistema programará los correos.

## 4. Envío Individual

Si solo necesita notificar a un único líder (por una solicitud particular o una contingencia), desplácese hacia el panel **"Envío Individual"**.
1. Escriba el nombre del jefe. El sistema autocompletará la búsqueda.
2. Verifique su correo electrónico.
3. Haga clic en **"Previsualizar mensaje"** para confirmar que el reporte adjunto de su equipo es exacto.
4. Haga clic en **"Enviar mensaje"**.

## 5. Cancelación de Envíos (Cola Pendiente)

El sistema cuenta con un "Poka-Yoke" (prevención de errores). Cuando envía notificaciones masivas, estas se quedan en "Espera" durante 60 segundos antes de ser despachadas a Power Automate.
Si nota un error inmediatamente después de enviar, haga clic en el botón rojo **"✕ Cancelar pendiente"** en la pantalla principal para detener el proceso antes de que los correos salgan de la bandeja.
