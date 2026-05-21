const HOJA_ACTIVA = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏢 Bot Predial El Marqués')
      .addItem('🔍 Consultar Adeudos (Monitor Avanzado)', 'abrirMonitorDeProgreso')
      .addToUi();
}

function abrirMonitorDeProgreso() {
  var html = HtmlService.createHtmlOutputFromFile('Monitor')
      .setWidth(800)
      .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Terminal de Consulta de Adeudos');
}

function getMarquesTokenCached() {
  var cache = CacheService.getUserCache();
  var token = cache.get('token_predial_marques');
  if (!token) {
    var url = "https://www.elmarquesdigital.gob.mx/auth/realms/Bus-Servicios-Marques-Prod/protocol/openid-connect/token";
    var payload = "client_id=[ID_CLIENTE]&client_secret=[ID_CLIENTE_SECRETO]&username=[USERNAME]&password=[YOUR PASSWORD]";
    
    var opciones = { 
      "method": "post", 
      "contentType": "application/x-www-form-urlencoded", 
      "payload": payload, 
      "muteHttpExceptions": true,
      "headers": {
        // Disfrazamos la petición para parecer un navegador Chrome en Windows
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept-Language": "es-MX,es-419;q=0.9,es;q=0.8"
      }
    };
    
    var respuesta = UrlFetchApp.fetch(url, opciones);
    var json = JSON.parse(respuesta.getContentText());
    if (json.access_token) {
      token = json.access_token;
      // Guardamos el token en caché por 250 segundos
      cache.put('token_predial_marques', token, 250); 
    } else {
      throw new Error("Fallo de autenticación con el servidor del Municipio.");
    }
  }
  return token;
}

/**
 * Lee las claves de la Columna A y el status de la Col B. 
 * Solo devuelve las que no han sido procesadas (Permite reanudar si se corta).
 */
function obtenerClavesParaProcesar() {
  var ultimaFila = HOJA_ACTIVA.getLastRow();
  if (ultimaFila < 2) return[];
  
  // Traemos Clave (A) y Status (B)
  var datos = HOJA_ACTIVA.getRange(2, 1, ultimaFila - 1, 2).getValues();
  var pendientes =[];
  
  for (var i = 0; i < datos.length; i++) {
    var clave = datos[i][0];
    var status = datos[i][1];
    
    // Si hay clave y el status está vacío (o fue error de script previo), lo metemos a pendientes
    if (clave && (!status || status === "ERROR SCRIPT")) {
      pendientes.push({ fila: i + 2, clave: clave.toString().trim() });
    }
  }
  return pendientes;
}

/**
 * Procesa UNA SOLA clave con sistema de auto-reintento para mayor seguridad en listas largas.
 */
function procesarUnaClave(claveCatastral) {
  var maxIntentos = 3; 
  
  for (var i = 0; i < maxIntentos; i++) {
    try {
      var token = getMarquesTokenCached();
      var url = `https://www.elmarquesdigital.gob.mx/ServiceBusEMDRest/v1/api/predial/adeudo?bimestre=6&clave_catastral=${claveCatastral}&periodo=2026`;
      
      var opciones = { 
        "method": "get", 
        "headers": { 
          "Authorization": "Bearer " + token,
          // Disfrazamos la petición
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
          "Accept": "application/json, text/plain, */*",
          "Accept-Language": "es-MX,es-419;q=0.9,es;q=0.8"
        }, 
        "muteHttpExceptions": true 
      };
      
      var respuesta = UrlFetchApp.fetch(url, opciones);
      var codigoHttp = respuesta.getResponseCode();
      
      // SISTEMA ANTI-BLOQUEO: Manejo del error 429
      if (codigoHttp === 429) {
        Utilities.sleep(5000); 
        throw new Error("429 Rate Limit - Servidor saturado");
      }

      // NUEVO: Manejo 401 (Token inválido o expirado antes de tiempo)
      if (codigoHttp === 401 || codigoHttp === 403) {
        CacheService.getUserCache().remove('token_predial_marques'); // Borramos el token viejo
        token = getMarquesTokenCached(); // Forzamos a pedir uno nuevo fresco
        throw new Error("Token caducado, forzando reintento..."); // El loop (maxIntentos) hará el resto
      }

      var json = JSON.parse(respuesta.getContentText());

      // ESPERA ALEATORIA: Entre 800ms y 1500ms
      var tiempoEspera = Math.floor(Math.random() * (1500 - 800 + 1)) + 800;
      Utilities.sleep(tiempoEspera); 

      var conceptos = { "IMPUESTO": 0, "REZAGO_ACTUAL": 0, "REZAGO_ANTERIOR": 0, "RECARGOS": 0, "ACTUALIZACION": 0, "MULTAS": 0, "REDONDEO": 0 };
      
      if ((json.status == "201" || json.status == 200) && json.body && json.body.conceptos) {
        json.body.conceptos.forEach(function(c) {
          if (c.descripcion.includes("IMPUESTO PREDIAL AÑO ACTUAL")) conceptos["IMPUESTO"] += c.importe;
          if (c.descripcion.includes("REZAGO ACTUAL")) conceptos["REZAGO_ACTUAL"] += c.importe;
          if (c.descripcion.includes("REZAGO AÑOS ANTERIORES")) conceptos["REZAGO_ANTERIOR"] += c.importe;
          if (c.descripcion.includes("RECARGOS")) conceptos["RECARGOS"] += c.importe;
          if (c.descripcion.includes("ACTUALIZACION")) conceptos["ACTUALIZACION"] += c.importe;
          if (c.descripcion.includes("MULTAS")) conceptos["MULTAS"] += c.importe;
          if (c.descripcion.includes("REDONDEO")) conceptos["REDONDEO"] += c.importe;
        });
      }
      
      return { status: json.status, mensaje: json.mensajeVentanilla || json.message, datos: json.body, conceptosDesglosados: conceptos };
      
    } catch (e) {
      if (i === maxIntentos - 1) { 
        return { status: 'ERROR SCRIPT', mensaje: e.message, datos: {}, conceptosDesglosados: {} };
      }
      // NUEVO: Retroceso Exponencial (2s, luego 4s, luego 8s...)
      var esperaExponencial = Math.pow(2, i + 1) * 1000;
      Utilities.sleep(esperaExponencial); 
    }
  }
}

/**
 * Prepara los encabezados a partir de la Columna B y formatea columnas de moneda
 */
function configurarHoja() {
  var encabezados =[
    "Status API", "Mensaje", "Total a Pagar", "Periodo Adeudo",
    "Impuesto Actual", "Rezago Actual", "Rezago Anterior", "Recargos", 
    "Actualización", "Multas", "Redondeo", "RFC", "Dirección Predio"
  ];
  
  HOJA_ACTIVA.getRange("A1").setValue("Clave Catastral").setFontWeight("bold").setBackground("#4a86e8").setFontColor("white");
  HOJA_ACTIVA.getRange(1, 2, 1, encabezados.length).setValues([encabezados]).setFontWeight("bold").setBackground("#4a86e8").setFontColor("white");
  
  // Formatear columnas enteras a moneda de una vez para optimizar
  HOJA_ACTIVA.getRange("D:D").setNumberFormat("$#,##0.00");
  HOJA_ACTIVA.getRange("F:L").setNumberFormat("$#,##0.00");
}

/**
 * Guarda un paquete (chunk) de resultados. 
 * Pega cada resultado en la fila exacta que le corresponde (por si las filas no son consecutivas).
 */
function guardarChunkDiscontinuo(chunk) {
  if (!chunk || chunk.length === 0) return;
  chunk.forEach(function(item) {
    // Escribe a partir de la Columna 2 (B) en la fila específica de esta clave
    HOJA_ACTIVA.getRange(item.fila, 2, 1, item.valores.length).setValues([item.valores]);
  });
  return true;
}