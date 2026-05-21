function getHojaActiva() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Adeudos_Predial");
  if (!hoja) {
    throw new Error("❌ No se encontró la pestaña llamada 'Adeudos_Predial'. Por favor, créala o renómbrala.");
  }
  return hoja;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏢 Bot Predial León')
      .addItem('🔍 Consultar Adeudos (Monitor Avanzado)', 'abrirMonitorLagos')
      .addSeparator() // Una línea divisoria para que se vea ordenado
      .addItem('📊 Ver Peticiones Restantes Hoy', 'mostrarCuotaAlert')
      .addToUi();
}


function abrirMonitorLagos() {
  var html = HtmlService.createHtmlOutputFromFile('MonitorLagos').setWidth(800).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Terminal Lagos de Moreno');
}

/**
 * Función Maestra para León
 * Incluye Caché de Sesión para reducir peticiones a la mitad (Ahorro de Cuota de Google).
 */
function procesarClaveLagos(claveCatastral) {
  var maxIntentos = 3;
  const headersBase = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "es-ES,es;q=0.9",
    "Origin": "https://pagoenlineagp.azurewebsites.net",
    "Referer": "https://pagoenlineagp.azurewebsites.net/",
    "Connection": "keep-alive"
  };

  for (var i = 0; i < maxIntentos; i++) {
    try {
      var cache = CacheService.getUserCache();
      var sesionGuardada = cache.get("sesion_leon");
      var token, cookiesParaEnviar;

      if (sesionGuardada) {
        var sesion = JSON.parse(sesionGuardada);
        token = sesion.token;
        cookiesParaEnviar = sesion.cookies;
      } else {
        var resInicio = UrlFetchApp.fetch("https://pagoenlineaapi-c6b4fza2f9a8emea.mexicocentral-01.azurewebsites.net/api/Autentication/GetToken?id=14053", {"method": "get", "headers": headersBase, "muteHttpExceptions": true});
        
        if (resInicio.getResponseCode() !== 200) {
           throw new Error("Fallo en API de Tokens (Mantenimiento del servidor).");
        }
        var cookiesSet = resInicio.getAllHeaders()['Set-Cookie'];
        cookiesParaEnviar = Array.isArray(cookiesSet) ? cookiesSet.join('; ') : cookiesSet;
        var jsonToken = JSON.parse(resInicio.getContentText());
        token = jsonToken.token.token;
        cache.put("sesion_leon", JSON.stringify({ token: token, cookies: cookiesParaEnviar }), 900);
      }

      var urlConsulta = "https://pagoenlineaapi-c6b4fza2f9a8emea.mexicocentral-01.azurewebsites.net/PagoEnLinea//consulta";
      var clienteObj = {
        "clientes": "C68AE520-B6ED-4EE4-BB6F-C6EEBCFFE225", "bbva": false, "banbajio": true, "codi": true, "oxxo": true, "consulta": true,
        "color1": "#f08b03", "color2": "#f08b03", "idCliente": 56, "clave": "14053", "nombre": "MUNICIPIO DE LAGOS DE MORENO JALISCO", "rfc": "MLM630725HU4",
        "calle": "JUAREZ", "noExterior": "SN", "localidad": "LAGOS DE MORENO", "municipio": "LAGOS DE MORENO", "estado": "JALISCO", "pais": "MEXICO", "codigoPostal": 47400, "regimenFiscal": "603"
      };
      var payloadConsulta = {
        "Campos": { "correo": "ale@gmail.com", "cuenta": claveCatastral },
        "Cliente": clienteObj,
        "Departamento": { "clientes": clienteObj.clientes, "iddepartamento": "01", "nombre": "" }
      };
      var headersConsulta = Object.assign({}, headersBase, {
        "Authorization": "Bearer " + token,
        "Content-Type": "application/json; charset=utf-8",
        "Cookie": cookiesParaEnviar
      });

      var resConsulta = UrlFetchApp.fetch(urlConsulta, { "method": "post", "headers": headersConsulta, "payload": JSON.stringify(payloadConsulta), "muteHttpExceptions": true });
      var codeConsulta = resConsulta.getResponseCode();
      
      // LÓGICA REFORZADA: Si el token falla, lanzamos un error que el 'catch' manejará.
      if (codeConsulta === 401 || codeConsulta === 403) {
        throw new Error("TOKEN_INVALIDO");
      }

      var textoConsulta = resConsulta.getContentText();
      if (!textoConsulta) throw new Error("Azure respondió vacío en Consulta");
      var jsonConsulta = JSON.parse(textoConsulta);

      if (!jsonConsulta.successfully || !jsonConsulta.data || !jsonConsulta.data.oidConsulta) {
        return { status: 'NO_DATA', mensaje: (jsonConsulta.data && jsonConsulta.data.message) ? jsonConsulta.data.message : 'Sin adeudos', datos: {}, conceptosDesglosados: {} };
      }

      var oidEnc = jsonConsulta.data.oidEncabezado;
      var oidCon = jsonConsulta.data.oidConsulta;
      var urlFinal = `https://pagoenlineaapi-c6b4fza2f9a8emea.mexicocentral-01.azurewebsites.net/PagoEnLinea//EstadoCuenta?Token=${token}&IdMunicipio=${clienteObj.clientes}&IdDepartamento=01&OidEncabezado=${oidEnc}&OidConsulta=${oidCon}`;
      
      var resFinal = UrlFetchApp.fetch(urlFinal, { "method": "get", "headers": { "User-Agent": headersBase["User-Agent"], "Cookie": cookiesParaEnviar, "Authorization": "Bearer " + token }, "muteHttpExceptions": true });
      var jsonFinal = JSON.parse(resFinal.getContentText());

      Utilities.sleep(Math.floor(Math.random() * (1200 - 600 + 1)) + 600);

      if (jsonFinal.successfully && jsonFinal.data) {
        const conceptos = { "IMPUESTO": 0, "REZAGO": 0, "ACTUALIZACION": 0, "RECARGOS": 0, "GASTOS": 0, "OTROS": 0 };
        let añosEncontrados = [];

        if (jsonFinal.data.detalle) {
          jsonFinal.data.detalle.forEach(function(c) {
            const imp = parseFloat(c.importe) || 0; const desc = c.descripcion.toUpperCase();
            if (desc.includes("PREDIAL NORMAL")) conceptos.IMPUESTO += imp;
            else if (desc.includes("REZAGO")) conceptos.REZAGO += imp;
            else if (desc.includes("ACTUALIZACION")) conceptos.ACTUALIZACION += imp;
            else if (desc.includes("RECARGOS")) conceptos.RECARGOS += imp;
            else if (desc.includes("GASTOS")) conceptos.GASTOS += imp;
            else conceptos.OTROS += imp;
            if (c.vigencia) añosEncontrados.push(parseInt(c.vigencia));
            const coincidencia = desc.match(/(\d{4})-(\d{4})/);
            if (coincidencia) { añosEncontrados.push(parseInt(coincidencia[1])); añosEncontrados.push(parseInt(coincidencia[2])); }
          });
        }
        let periodoTexto = añosEncontrados.length > 0 ? `${Math.min(...añosEncontrados)}06 al ${Math.max(...añosEncontrados)}06` : "N/A";

        return { status: 'OK', mensaje: 'Consulta Exitosa', periodo: periodoTexto, datos: jsonFinal.data, conceptosDesglosados: conceptos };
      } else { throw new Error("Respuesta final inválida"); }

    } catch (e) {
      // MANEJO DE ERRORES INTELIGENTE
      if (e.message === "TOKEN_INVALIDO") {
        cache.remove("sesion_leon"); // Borramos la sesión mala
        if (i === maxIntentos - 1) { // Si falló 3 veces seguidas con token nuevo
          return { status: 'ERROR SCRIPT', mensaje: 'Fallo de autenticación persistente. Servidor ocupado.', datos: {}, conceptosDesglosados: {} };
        }
        // No hacemos nada más, dejamos que el loop for() reintente y pida un token nuevo.
      } else {
        // Para cualquier otro tipo de error, si es el último intento, lo reportamos.
        if (i === maxIntentos - 1) {
          return { status: 'ERROR SCRIPT', mensaje: e.message, datos: {}, conceptosDesglosados: {} };
        }
      }
      Utilities.sleep(Math.pow(2, i + 1) * 2000); // Pausa más larga para darle tiempo al servidor.
    }
  }
}

/**
 * Lee la hoja "Adeudos_Predial" en lugar de variables globales.
 */
function obtenerClavesParaProcesarLagos() {
  var hoja = getHojaActiva();
  var ultimaFila = hoja.getLastRow();
  
  if (ultimaFila < 2) return [];
  
  var datos = hoja.getRange(2, 1, ultimaFila - 1, 2).getValues();
  var pendientes = [];
  
  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0] && (!datos[i][1] || datos[i][1] === "ERROR SCRIPT")) {
      pendientes.push({ fila: i + 2, clave: datos[i][0].toString().trim() });
    }
  }
  return pendientes;
}

/**
 * Configura encabezados y da formato financiero a rangos específicos.
 */
function configurarHojaLagos() {
  var hoja = getHojaActiva();
  var encabezados = [
    "Status API", "Mensaje", "Total a Pagar", "Periodo Adeudo", 
    "Contribuyente", "Dirección", "Observaciones", "Descuento", 
    "Impuesto", "Rezago", "Actualización", "Recargos", "Gastos", "Otros"
  ];
  
  hoja.getRange("A1").setValue("Clave Catastral").setFontWeight("bold").setBackground("#f08b03").setFontColor("white");
  hoja.getRange(1, 2, 1, encabezados.length).setValues([encabezados]).setFontWeight("bold").setBackground("#f08b03").setFontColor("white");
  
  hoja.getRange("C2:C").setNumberFormat("$#,##0.00"); 
  hoja.getRange("H2:N").setNumberFormat("$#,##0.00"); 
}

/**
 * Guarda los registros en Google Sheets sin perder conexión.
 */
function guardarChunkDiscontinuo(chunk) {
  if (!chunk || chunk.length === 0) return;
  var hoja = getHojaActiva();
  chunk.forEach(function(item) {
    hoja.getRange(item.fila, 2, 1, item.valores.length).setValues([item.valores]);
  });
  SpreadsheetApp.flush(); // Obliga a Google a guardar la info en el archivo real
  return true;
}
function obtenerCuotaRestante() {
  // Esta es la función que le pregunta a Google el límite real
  return MailApp.getRemainingDailyQuota(); // <--- ✅ CORRECTO
}

function mostrarCuotaAlert() {
  var cuota = obtenerCuotaRestante();
  var mensaje = "📊 Tienes " + cuota.toLocaleString() + " peticiones web restantes para las próximas 24 horas.";
  
  // Si la cuota es alta (más de 50,000), le damos un mensaje de confianza
  var extra = cuota > 50000 ? "\n\n✅ Tienes cuota de sobra para procesar miles de registros." : "";
  
  SpreadsheetApp.getUi().alert("Estado de la Cuenta Google", mensaje + extra, SpreadsheetApp.getUi().ButtonSet.OK);
}
