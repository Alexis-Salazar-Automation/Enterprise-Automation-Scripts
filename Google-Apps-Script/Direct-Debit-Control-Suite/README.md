# Direct Debit Control & Validation Suite (Google Apps Script)

## 📝 Descripción
Solución corporativa para el control, validación y auditoría de domiciliaciones bancarias. El sistema asegura que las instrucciones de cargo automático coincidan exactamente con la base de datos maestra de la empresa, previniendo errores de cobro y morosidad.

## 🚀 Funcionalidades Técnicas

### 1. Auditoría y Conciliación Automática (Data Integrity)
- **Validación Cross-Reference:** El sistema compara diariamente los registros locales contra la Cartera Maestra para detectar discrepancias en estatus (ej: Domiciliación activa en papel pero inactiva en sistema).
- **Control de Vigencia:** Lógica avanzada para determinar el registro vigente por lote basado en la estampa de tiempo más reciente, evitando duplicidad de gestiones.

### 2. Dashboard de KPIs por Correo
- Generación de reportes diarios en HTML con métricas clave:
    - **Salud de Cartera (%):** Porcentaje de registros sin discrepancias.
    - **Monto Activo Mensual:** Suma total de cargos proyectados.
    - **Alertas de Morosidad Crítica:** Identificación inmediata de clientes con >= 2 meses de adeudo a pesar de tener domiciliación activa.

### 3. Procesamiento y UX
- **Filtro en Cascada:** Optimización de búsqueda sobre +60,000 registros.
- **Motor de PDF:** Generación masiva de acuses de entrega con firma digital y envío automatizado.
- **Seguridad de Datos:** Implementación de truncado de números de tarjeta para cumplimiento de políticas de privacidad, permitiendo la visualización solo a usuarios autorizados.

## 🛠️ Tecnologías
- **Google Apps Script** (V8 Engine)
- **HTML5/JavaScript**
- **Bootstrap 4.5**
- **Intl API** (Para formateo de moneda y fechas)