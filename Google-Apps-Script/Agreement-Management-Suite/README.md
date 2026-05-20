# Agreement & Negotiation Management Suite (Google Apps Script)

## 📝 Descripción
Sistema integral desarrollado en **Google Apps Script** para la estandarización, seguimiento y cierre de convenios de pago y negociaciones financieras. Diseñado para manejar carteras masivas de clientes (+60,000 registros) con una interfaz de usuario fluida y procesos automatizados de auditoría.

## 🚀 Funcionalidades Técnicas Destacadas

### 1. Optimización de Consultas (High-Volume Data)
- **Filtrado en Cascada:** Implementación de lógica de búsqueda que reduce el tiempo de respuesta del servidor de 3 minutos a menos de 5 segundos mediante la indexación local de listas maestras.
- **Buscador Avanzado:** Motor de búsqueda con soporte para múltiples criterios (referencia, analista, desarrollo) y filtros dinámicos.

### 2. Automatización de Documentos y Auditoría
- **Generador de PDF (Acuses):** Creación automática de acuses de recibo en formato PDF mediante la manipulación de plantillas de Google Sheets y envío automatizado por correo electrónico (Blob conversion).
- **Hoja de Vida del Lote (Logging):** Registro detallado de cada evento (cambios de estatus, recordatorios enviados, asignaciones) para garantizar la trazabilidad de cada trámite.

### 3. Seguridad y Control de Concurrencia
- **Race Condition Prevention:** Uso de `LockService` para gestionar el acceso concurrente de múltiples analistas, garantizando la integridad de los folios únicos y los datos en el Concentrado.
- **Validación de Roles:** Menús dinámicos que se adaptan según el nivel de permiso del usuario (Editor vs. Lector).

### 4. Ciclo de Vida Automatizado (Batch Processing)
- **Triggers Diarios:** Procesos programados para la sincronización de carteras remotas, actualización de estatus de vencimiento y envío de reportes de desempeño gerencial.

## 🛠️ Tecnologías
- **Google Apps Script** (V8 Engine)
- **HTML5/CSS3/JavaScript** (Frontend de los formularios)
- **Bootstrap 4.5 & Select2** (UI/UX componentes)
- **Gmail API & Drive API** (Distribución y almacenamiento)