# Portfolio Sync Engine (Python & GitHub Actions)

## 🚀 Descripción
Bot de automatización diseñado para el ciclo de vida de datos financieros. El sistema extrae información de una API privada, procesa los datos con **Pandas**, sincroniza una hoja de cálculo en la nube y genera respaldos históricos automáticos.

**Impacto:** Sustitución de un proceso manual de 5 horas por una ejecución automatizada de 45 segundos.

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