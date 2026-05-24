# Mass Notification & Engagement Engine (Google Apps Script)

## 📝 Descripción
Motor de comunicación masiva omnicanal diseñado para el envío automatizado de notificaciones vía Email y WhatsApp. El sistema procesa miles de gestiones semanales integrando una lógica de balanceo de carga y recuperación de errores en tiempo real.

## 🚀 Capacidades Técnicas Avanzadas

### 1. Motor de Procesamiento Dual
Diseñé el sistema con dos filosofías de ejecución seleccionables por el usuario:
- **🛡️ Modo Seguro (Sequential):** Envío uno a uno con validación profunda de respuesta. Ideal para detectar errores específicos de contacto (ej. correos inexistentes).
- **🚀 Modo Urgente (Batch):** Envío en lotes de 200 registros por petición. Optimizado para velocidad máxima cuando la integridad de los datos es alta.

### 2. Algoritmo de Auto-Sanación (Self-Healing)
Implementé una lógica de redundancia que entra en acción si la API rechaza una notificación por discrepancias en las variables de la plantilla. El script intenta automáticamente diferentes combinaciones de campos (fallbacks) para garantizar el envío sin intervención humana.

### 3. Extracción y Cruce de Datos (ETL)
- **Cruce Automático:** El sistema extrae en segundos los IDs de cobranza y vivienda de una base maestra de +60,000 registros para preparar los lotes de envío.
- **Validación de Plantillas:** Gestión dinámica de catálogos para asegurar que las notificaciones utilicen la identidad visual correcta.

## 🛠️ Tecnologías
- **Google Apps Script** (V8 Engine)
- **REST APIs** (JSON Payloads complejos)
- **JavaScript Asíncrono** (Interfaz de usuario y polling)