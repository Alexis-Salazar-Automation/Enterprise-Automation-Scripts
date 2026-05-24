# Bulk Status & Assignment Manager (Google Apps Script)

## 📝 Descripción
Herramienta de nivel empresarial para la gestión masiva de carteras. Permite realizar actualizaciones simultáneas de estatus y reasignaciones de ejecutivos sobre miles de registros mediante una arquitectura de procesamiento por lotes (Chunking) y una interfaz de monitoreo en tiempo real.

## 🚀 Innovaciones Técnicas

### 1. Arquitectura ETL Unificada
- **Extracción:** Cruce dinámico de datos entre la hoja de trabajo del usuario y la base de datos maestra de 60,000 registros.
- **Transformación:** Lógica de normalización de nombres y IDs de estatus mediante catálogos centralizados.
- **Carga:** Inserción controlada en la base de datos operativa.

### 2. Resiliencia y Manejo de Errores (Smart Retry)
Para evitar caídas por saturación del servidor (API 429/503), implementé:
- **Batch Processing:** Envío de datos en bloques de 200 registros.
- **Protocolo de Cooldown:** El sistema detecta cuando el servidor está saturado y activa automáticamente una pausa de 30 segundos, informando al usuario mediante un cronómetro visual en la terminal.

### 3. Interfaz de Monitoreo (Real-Time Terminal)
Desarrollé una interfaz basada en Canvas e Inyección de HTML para simular una terminal de comandos. Esta terminal permite al usuario:
- Ver el progreso real de las peticiones a la API.
- Visualizar errores específicos por bloque de datos.
- Recibir alertas de seguridad y estado del backend.

## 🛠️ Tecnologías
- **Google Apps Script** (Lógica de control y API Fetch)
- **HTML5 Canvas & JavaScript** (Interfaz de monitoreo)
- **CacheService** (Gestión de estados del bot)