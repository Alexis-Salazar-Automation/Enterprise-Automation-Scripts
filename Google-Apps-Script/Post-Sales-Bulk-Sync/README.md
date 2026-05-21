# Post-Sales Bulk Data Synchronizer (Google Apps Script)

## 📝 Descripción
Este sistema automatiza la extracción y consolidación de la cartera de postventa, la cual se encuentra fragmentada en más de 50 proyectos y 200 condominios. Antes de esta automatización, el proceso requería **5 horas diarias** de trabajo manual; actualmente se realiza de forma desatendida en **40 minutos**.

## 🚀 Desafíos Técnicos y Soluciones

### 1. Optimización de la API (Payload Hacking)
Al analizar el comportamiento del endpoint de la empresa, identifiqué que el envío de valores `null` en los parámetros de filtrado permitía la descarga masiva por proyecto completo en lugar de realizar peticiones individuales por cada uno de los 200 condominios, reduciendo drásticamente el consumo de recursos y el tiempo de ejecución.

### 2. Gestión de Tiempos de Ejecución (Quotas Bypass)
Dado que el volumen de datos excede el límite de ejecución de 6 minutos de Google Apps Script (Standard Quota), implementé una lógica de **auto-reinicio**:
- El script monitorea su propio tiempo de ejecución.
- Antes de alcanzar el límite, guarda el progreso actual y programa un nuevo disparador (Trigger) para continuar desde el último registro procesado.

### 3. Integración y Notificación
- **Autenticación Automática:** Gestión de tokens para la comunicación con la API corporativa.
- **Reporting:** Envío de notificaciones automáticas vía email confirmando la integridad de la base consolidada cada mañana a las 7:00 AM.

## 🛠️ Tecnologías
- **Google Apps Script** (V8 Engine)
- **REST APIs** (JSON Payload manipulation)
- **Time-driven Triggers**