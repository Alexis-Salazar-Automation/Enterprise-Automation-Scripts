# Portfolio Sync Engine (Python & GitHub Actions)

## 🚀 Descripción
Motor de sincronización y orquestación de datos diseñado para actuar como la **infraestructura base** del área. El sistema automatiza la extracción de la cartera de clientes desde una API privada, centralizando la información en Google Drive y Google Sheets.

**Impacto y Valor de Negocio:**
- **Eliminación del Error Humano:** Garantiza la disponibilidad de datos al 100%, eliminando la dependencia de ejecuciones manuales propensas a olvidos.
- **Ecosistema Integrado:** Actúa como el proveedor de datos (Data Provider) para todas las demás automatizaciones del departamento.
- **Trazabilidad Total:** Implementación de reportes de variaciones diarios que permiten identificar cambios críticos en la cartera de forma inmediata.

## 🛠️ Stack Tecnológico
- **Lenguaje:** Python 3.10
- **Librerías de Datos:** Pandas, Numpy.
- **Integraciones:** Google Sheets API, Google Drive API, REST APIs (Requests).
- **Infraestructura:** GitHub Actions (CI/CD) para orquestación programada (CRON).

## ⚙️ Arquitectura del Proceso (ETL)
1. **Extracción:** Autenticación mediante Handshake JWT y descarga masiva de registros JSON.
2. **Transformación:** Limpieza de datos nulos, normalización de tipos y mapeo de campos de negocio.
3. **Carga:** Sincronización mediante `gspread` optimizada para alto volumen de filas.
4. **Resiliencia:** Implementación de política de reintentos (Backoff) para tolerar inestabilidad en la API.
5. **Notificación:** Generación de reporte HTML dinámico comparando la cartera de "ayer" vs "hoy".

## 🤖 Automatización
El script se ejecuta de forma desatendida mediante un Workflow de GitHub Actions dos veces al día, garantizando que el equipo de operaciones siempre trabaje con datos frescos.