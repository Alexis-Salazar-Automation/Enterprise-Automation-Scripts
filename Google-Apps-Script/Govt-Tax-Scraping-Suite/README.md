# Government Tax Automation & Web Scraping Suite (Google Apps Script)

## 📝 Descripción
Colección de motores de extracción de datos (scrapers) diseñados para automatizar la consulta masiva de adeudos prediales en diferentes municipios. La suite procesa activamente más de **25,000 registros mensuales**, eliminando el error humano y reduciendo drásticamente los tiempos de auditoría fiscal.

## 🚀 Diversidad de Estrategias Técnicas

Cada municipio presentó un desafío de seguridad distinto, los cuales fueron resueltos con las siguientes técnicas:

### 1. Gestión de Cookies y Sesión Azure (León/Lagos)
- **El Reto:** El servidor utiliza protección de sesión basada en infraestructura Azure con cookies temporales y tokens volátiles.
- **La Solución:** Implementación de un sistema de **Handshake** inicial para capturar el encabezado `Set-Cookie`. Se utilizó `CacheService` para persistir la sesión por 15 minutos, reduciendo las peticiones de autenticación a la mitad y optimizando las cuotas de Google.

### 2. Autenticación Bearer & OAuth2 (El Marqués)
- **El Reto:** Acceso restringido mediante tokens dinámicos generados por un servidor de identidad.
- **La Solución:** Desarrollo de un flujo de autenticación automático que obtiene y refresca el `access_token` mediante peticiones POST (x-www-form-urlencoded), permitiendo consultas ininterrumpidas en lotes grandes.

### 3. Resiliencia y Control de Tráfico (Corregidora)
- **El Reto:** El servidor bloquea peticiones rápidas (Error 429 - Rate Limit).
- **La Solución:** Implementación de **Retroceso Exponencial (Exponential Backoff)** y esperas aleatorias (Jitter) para simular comportamiento humano y evitar el baneo de la IP de Google.

## 🛠️ Stack Tecnológico
- **Google Apps Script** (Lógica de Backend)
- **Asynchronous JavaScript** (Terminales de monitoreo en tiempo real)
- **CacheService API** (Gestión de sesiones)
- **RegEx** (Limpieza y normalización de datos catastrales)