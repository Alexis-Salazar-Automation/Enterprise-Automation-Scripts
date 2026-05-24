# Telephony Analytics & Auditor (Google Apps Script)

## 📝 Descripción
Motor de extracción y análisis diseñado para auditar el desempeño de las gestiones telefónicas. El sistema consolida datos de múltiples plataformas (App propietaria vs. Call Center oficial), aplica lógica de calificación de llamadas y genera tableros de control automáticos.

## 🚀 Capacidades de Análisis de Datos (BI)

### 1. Motor de Calificación por Regex
Implementé un algoritmo de procesamiento de lenguaje natural simple (RegEx) para analizar las observaciones de los ejecutivos y calificar automáticamente si una llamada fue **Efectiva** o **No Efectiva**, basándose en palabras clave y umbrales de duración (segundos).

### 2. Consolidación y Dashboards Dinámicos
- **ETL Automatizado:** El script extrae registros diarios vía API, limpia la información y la clasifica por origen.
- **Visualización Interactiva:** Generación automática de Tablas Dinámicas y Slicers en Google Sheets para permitir al equipo gerencial filtrar el desempeño por ejecutivo o por rango de fechas.

### 3. Reporteo Ejecutivo (Email Delivery)
Al finalizar la extracción, el sistema construye un reporte corporativo en **HTML/CSS** que incluye:
- KPIs de adopción de la nueva App.
- Tasas de efectividad comparativas.
- Identificación de "Top Performers" del periodo.
- Archivo Excel (.xlsx) adjunto generado dinámicamente mediante el motor de Google.

## 🛠️ Tecnologías
- **Google Apps Script** (Backend & SMTP)
- **Regular Expressions (RegEx)** (Data Mining)
- **PivotTable API** (Generación de reportes)
- **HTML5 / CSS3** (Diseño de correos corporativos)