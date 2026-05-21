/**
 * Función de ayuda para obtener la hoja correcta siempre,
 * sin importar de dónde llame la función.
 */
function getHojaActiva() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏢 Bot Predial Corregidora')
      .addItem('🔍 Consultar Adeudos (Monitor Avanzado)', 'abrirMonitorCorregidora')
      .addToUi();
}

function abrirMonitorCorregidora() {
  var html = HtmlService.createHtmlOutputFromFile('MonitorCorregidora')
      .setWidth(800)
      .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Terminal de Consulta de Adeudos - Corregidora');
}

function obtenerClavesParaProcesarCorregidora() {
  var hoja = getHojaActiva();
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  
  var datos = hoja.getRange(2, 1, ultimaFila - 1, 2).getValues();
  var pendientes = [];
  var LIMITE_POR_SESION = 500; 
  
  for (var i = 0; i < datos.length; i++) {
    var clave = datos[i][0];
    var status = datos[i][1];
    
    if (clave && (!status || status === "ERROR SCRIPT")) {
      pendientes.push({ fila: i + 2, clave: clave.toString().trim() });
    }
    
    if (pendientes.length >= LIMITE_POR_SESION) {
      break; 
    }
  }
  return pendientes;
}

function procesarClaveCorregidora(claveCatastral) {
  var maxIntentos = 3; 

  const hoy = new Date();
  const dia = ('0' + hoy.getDate()).slice(-2);
  const mes = ('0' + (hoy.getMonth() + 1)).slice(-2); 
  const anio = hoy.getFullYear();
  const fechaFormato = `${dia}${mes}${anio}`;

  for (var i = 0; i < maxIntentos; i++) {
    try {
      var url = `https://www.corregidoraenlinea.gob.mx:9990/api/predial/na/formality/paying/?pCveCat=${claveCatastral}&pPer=202606&pFec=${fechaFormato}`;
      
      var opciones = { 
        "method": "get", 
        "headers": { 
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }, 
        "muteHttpExceptions": true,
        "followRedirects": true,
        "connectTimeout": 60000 
      };
      
      var respuesta = UrlFetchApp.fetch(url, opciones);
      var codigoHttp = respuesta.getResponseCode();
      
      if (codigoHttp === 429) {
        Utilities.sleep(5000); 
        throw new Error("429 Rate Limit - Servidor saturado");
      }

      var json = JSON.parse(respuesta.getContentText());

      //var tiempoEspera = Math.floor(Math.random() * (1500 - 800 + 1)) + 800;
      var tiempoEspera = Math.floor(Math.random() * (4500 - 2500 + 1)) + 2500;
      Utilities.sleep(tiempoEspera); 

      if (json.items && json.items.length > 0) {
        const item = json.items[0];
        const conceptos = { "IMPUESTO": 0, "REZAGO_ACTUAL": 0, "REZAGO_ANTERIOR": 0, "RECARGOS": 0, "ACTUALIZACION": 0, "MULTAS": 0, "REDONDEO": 0 };
        
        if(item.fte_ing && Array.isArray(item.fte_ing)) {
          item.fte_ing.forEach(function(c) {
            if (c.desc_fte.includes("IMPUESTO PREDIAL")) conceptos["IMPUESTO"] += c.importe;
            if (c.desc_fte.includes("REZAGO PREDIAL ACTUAL")) conceptos["REZAGO_ACTUAL"] += c.importe;
            if (c.desc_fte.includes("REZAGO PREDIAL EJERCICIOS ANTERIORES")) conceptos["REZAGO_ANTERIOR"] += c.importe;
            if (c.desc_fte.includes("RECARGOS")) conceptos["RECARGOS"] += c.importe;
            if (c.desc_fte.includes("ACTUALIZACIÓN") || c.desc_fte.includes("ACTUALIZACION")) conceptos["ACTUALIZACION"] += c.importe;
            if (c.desc_fte.includes("MULTAS")) conceptos["MULTAS"] += c.importe;
            if (c.desc_fte.includes("AJUSTE POR REDONDEO")) conceptos["REDONDEO"] += c.importe;
          });
        }
        
        return { status: 'OK', mensaje: 'Adeudo encontrado', datos: item, conceptosDesglosados: conceptos };
      } else {
        return { status: 'NO_DATA', mensaje: 'Sin adeudos o clave no encontrada.', datos: {}, conceptosDesglosados: {} };
      }
      
    } catch (e) {
      if (i === maxIntentos - 1) { 
        return { status: 'ERROR SCRIPT', mensaje: e.message, datos: {}, conceptosDesglosados: {} };
      }
      var esperaExponencial = Math.pow(2, i + 1) * 1000; 
      Utilities.sleep(esperaExponencial); 
    }
  }
}

function configurarHojaCorregidora() {
  var hoja = getHojaActiva();
  var encabezados = [
    "Status API", "Mensaje", "Total Adeudo", "Periodo Inicial", "Periodo Final", 
    "Nombre Contribuyente", "RFC", "Dirección Predio", "Dirección Fiscal",
    "Impuesto", "Rezago Actual", "Rezago Anterior", "Recargos", 
    "Actualización", "Multas", "Redondeo", "Valor Terreno", "Valor Construcción"
  ];
  
  hoja.getRange("A1").setValue("Clave Catastral").setFontWeight("bold").setBackground("#d9534f").setFontColor("white");
  hoja.getRange(1, 2, 1, encabezados.length).setValues([encabezados]).setFontWeight("bold").setBackground("#d9534f").setFontColor("white");
  
  // Modificado: Formato a rangos cerrados a partir de la fila 2
  hoja.getRange("C2:C").setNumberFormat("$#,##0.00");
  hoja.getRange("J2:R").setNumberFormat("$#,##0.00");
}

function guardarChunkDiscontinuo(chunk) {
  if (!chunk || chunk.length === 0) return;
  var hoja = getHojaActiva();
  chunk.forEach(function(item) {
    hoja.getRange(item.fila, 2, 1, item.valores.length).setValues([item.valores]);
  });
  SpreadsheetApp.flush(); 
  return true;
}