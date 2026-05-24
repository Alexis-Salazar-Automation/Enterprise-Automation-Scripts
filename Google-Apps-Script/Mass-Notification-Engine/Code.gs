// ==========================================
// 🚀 CONFIGURACIÓN GLOBAL Y LISTAS FIJAS
// ==========================================
const ID_ARCHIVO_CARTERA = '[ID_Hoja_calculo]';
const NOMBRE_HOJA_CARTERA = '[Sheet_Name]';

const CATALOGO_PLANTILLAS_CORREO =[
  Calatogo de platillas valido para correo
];

const CATALOGO_PLANTILLAS_WHATSAPP =[
  Calatogo de platillas valido para whatsapp
];

// ==========================================
// 🤖 MENÚ DE INTERFAZ
// ==========================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📢 Notificador Masivo')
      .addItem('1️⃣ Fase 1: Preparar Hoja y Extraer Lotes', 'fase1PrepararLotes')
      .addItem('🚀 Fase 2: Enviar Notificaciones', 'fase2PreCheck')
      .addSeparator()
      .addItem('🧹 Fase 3: Archivar y Limpiar', 'fase3LimpiarYArchivar')
      .addToUi();
}

// ==========================================
// 🔑 TOKEN DINÁMICO
// ==========================================
function obtenerTokenGPH() {
  var cache = CacheService.getUserCache();
  var tokenGuardado = cache.get("gph_token");
  if (tokenGuardado) return tokenGuardado;

  var correoActivo = Session.getActiveUser().getEmail() || "[Correo_Valido]";
  var nombreDerivado = correoActivo.split('@')[0].toUpperCase();
  var credencialesDinamicas = { "googleId": "[ID_GOOGLE]", "email": correoActivo, "name": nombreDerivado };

  try {
    var resp = UrlFetchApp.fetch("[URL_ENDPOINT]", {
      "method": "post", "contentType": "application/json",
      "payload": JSON.stringify({"data": credencialesDinamicas}), "muteHttpExceptions": true
    });
    var json = JSON.parse(resp.getContentText());
    if (json.status === 1) {
      cache.put("gph_token", json.data.token, 1800); 
      return json.data.token;
    }
    return null;
  } catch(e) { return null; }
}

// ==========================================
// 🖥️ LANZADORES DE INTERFAZ
// ==========================================
function fase1PrepararLotes() {
  // ✅ CORRECCIÓN APLICADA AQUÍ: Se llama a la función correcta
  lanzarMatrix('fase1');
}

function fase2PreCheck() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaLocal = libro.getSheetByName("Envios_Masivos") || libro.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  var pCorreo = hojaLocal.getRange("B1").getValue().toString().trim();
  var pWhats = hojaLocal.getRange("D1").getValue().toString().trim();
  
  if (!pCorreo && !pWhats) return ui.alert("Error", "Debes seleccionar al menos una plantilla.", ui.ButtonSet.OK);

  var datos = hojaLocal.getRange(5, 1, Math.max(1, hojaLocal.getLastRow() - 4), 9).getValues();
  var totalSeleccionados = 0;
  for (var i = 0; i < datos.length; i++) {
    var lote = datos[i][0].toString().trim(), nombre = datos[i][1].toString().trim();
    var medio = datos[i][7].toString().trim(), estatus = datos[i][8].toString().trim();
    if (!lote || nombre === "NO ENCONTRADO" || !medio || estatus === "✅ Correo | ✅ Whats") continue;
    
    if ((medio.includes("CORREO") || medio === "AMBOS") && !estatus.includes("✅ Correo")) totalSeleccionados++;
    if ((medio.includes("WHATSAPP") || medio === "AMBOS") && !estatus.includes("✅ Whats")) totalSeleccionados++;
  }

  if (totalSeleccionados === 0) return ui.alert("Información", "No hay registros pendientes para enviar.", ui.ButtonSet.OK);

  // 🛡️ LANZAMOS EL SELECTOR DE MOTOR
  var html = HtmlService.createHtmlOutput(SELECTOR_HTML.replace('{{TOTAL}}', totalSeleccionados)).setWidth(500).setHeight(360);
  ui.showModalDialog(html, '⚙️ Motor de Procesamiento');
}

// Esta función es llamada desde Fase 1 o desde el Modal HTML para abrir Matrix
function lanzarMatrix(fase) {
  var htmlFinal = MATRIX_HTML.replace('{{FASE}}', fase);
  var htmlObj = HtmlService.createHtmlOutput(htmlFinal).setWidth(700).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(htmlObj, 'Sistema Automatizado...');
}

// ==========================================
// ⚙️ BACKEND - FASE 1
// ==========================================
function fase1Backend() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaLocal = libro.getActiveSheet();

  if (hojaLocal.getName() !== "Envios_Masivos") hojaLocal.setName("Envios_Masivos");

  if (hojaLocal.getRange("A1").getValue() !== "PLANTILLA CORREO:") {
    hojaLocal.clear(); 
    hojaLocal.getRange("A1").setValue("PLANTILLA CORREO:").setFontWeight("bold").setBackground("#cfe2f3");
    hojaLocal.getRange("C1").setValue("PLANTILLA WHATS:").setFontWeight("bold").setBackground("#cfe2f3");
    hojaLocal.getRange("A2").setValue("VIGENCIA:").setFontWeight("bold").setBackground("#cfe2f3");
    hojaLocal.getRange("B1:B2").setBackground("#fff2cc");
    hojaLocal.getRange("D1").setBackground("#fff2cc");
    hojaLocal.getRange("B2").setValue(new Date()).setNumberFormat("dd/mm/yyyy");
    
    var cabeceras =["Lote (Vivienda)", "Nombre Cliente", "ID Cobranza", "ID Vivienda", "Condominio", "Proyecto", "Ref. Mantenimiento", "Medio", "Estatus Envío"];
    hojaLocal.getRange(4, 1, 1, cabeceras.length).setValues([cabeceras]).setFontWeight("bold").setBackground("#4c1130").setFontColor("white");
    hojaLocal.setFrozenRows(4);
  }

  var reglaCorreo = SpreadsheetApp.newDataValidation().requireValueInList(CATALOGO_PLANTILLAS_CORREO, true).build();
  hojaLocal.getRange("B1").setDataValidation(reglaCorreo);
  var reglaWhats = SpreadsheetApp.newDataValidation().requireValueInList(CATALOGO_PLANTILLAS_WHATSAPP, true).build();
  hojaLocal.getRange("D1").setDataValidation(reglaWhats);

  var hojaMaestra = SpreadsheetApp.openById(ID_ARCHIVO_CARTERA).getSheetByName(NOMBRE_HOJA_CARTERA);
  var datosMaestros = hojaMaestra.getDataRange().getValues();
  var enc = datosMaestros[0];
  
  var diccCartera = {};
  for (var i = 1; i < datosMaestros.length; i++) {
    var viviendaVal = datosMaestros[i][enc.indexOf("vivienda")] ? datosMaestros[i][enc.indexOf("vivienda")].toString().trim() : "";
    if (viviendaVal) {
      diccCartera[viviendaVal] = {
        idCobranza: datosMaestros[i][enc.indexOf("id_cobranza")], nombre: datosMaestros[i][enc.indexOf("clienteNombre")],
        referencia: datosMaestros[i][enc.indexOf("RefMantenimiento")], idVivienda: datosMaestros[i][enc.indexOf("idVivienda")],
        condominio: datosMaestros[i][enc.indexOf("ncondominio")], proyecto: datosMaestros[i][enc.indexOf("nproyecto")]
      };
    }
  }

  var ultimaFila = hojaLocal.getLastRow();
  if (ultimaFila < 5) return { done: true, pct: 100, msg: "Interfaz preparada. Pega los lotes y repite la Fase 1.", titulo: "Fase 1 Lista" };
  
  var lotesUsuario = hojaLocal.getRange(5, 1, ultimaFila - 4, 1).getValues();
  var resultadosExtraidos = lotesUsuario.map(function(row) {
    var loteIngresado = row[0].toString().trim();
    if (!loteIngresado) return["", "", "", "", "", ""];
    var d = diccCartera[loteIngresado];
    return d ?[d.nombre, d.idCobranza, d.idVivienda, d.condominio, d.proyecto, d.referencia] :["NO ENCONTRADO", "-", "-", "-", "-", "-"];
  });

  hojaLocal.getRange(5, 2, resultadosExtraidos.length, 6).setValues(resultadosExtraidos);
  var listaMedios =["1 - CORREO", "3 - WHATSAPP", "AMBOS"];
  var reglaMedios = SpreadsheetApp.newDataValidation().requireValueInList(listaMedios, true).build();
  hojaLocal.getRange(5, 8, resultadosExtraidos.length, 1).setDataValidation(reglaMedios);
  
  return { done: true, pct: 100, msg: "Cruce completado. Ejecuta la Fase 2.", titulo: "Fase 1 Completada" };
}


// ==========================================
// 🛡️ BACKEND - MODO SEGURO (1 POR 1 + AUTO-SANACIÓN)
// ==========================================
function fase2BackendSeguro() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaLocal = libro.getSheetByName("Envios_Masivos");
  
  var pCorreo = hojaLocal.getRange("B1").getValue().toString().trim();
  var pWhats = hojaLocal.getRange("D1").getValue().toString().trim();
  var vig = hojaLocal.getRange("B2").getValue();

  var idPlantillaCorreo = pCorreo ? pCorreo.split(" - ")[0] : null;
  var idPlantillaWhats = pWhats ? pWhats.split(" - ")[0] : null;
  var vigenciaTexto = (vig instanceof Date) ? Utilities.formatDate(vig, Session.getScriptTimeZone(), "dd-MM-yyyy") : vig.toString().trim().replace(/\//g, "-");

  var datos = hojaLocal.getRange(5, 1, hojaLocal.getLastRow() - 4, 9).getValues();
  var pendientes =[], procesadosExitososOConErrorRevisado = 0;

  for (var i = 0; i < datos.length; i++) {
    var viviendaCompleta = datos[i][0].toString().trim(), nombreCl = datos[i][1].toString().trim();
    var medioElegido = datos[i][7].toString().trim(), estatusActual = datos[i][8].toString().trim();

    if (!viviendaCompleta || nombreCl === "NO ENCONTRADO" || !medioElegido) continue;
    
    var basePayload = {
      "idVivienda": parseInt(datos[i][3]), "idCobranza": datos[i][2].toString().trim(), "nombre": nombreCl,
      "campos":[
        { "N": "CLIENTE", "V": nombreCl }, { "N": "LOTE", "V": viviendaCompleta.split("-").pop() },
        { "N": "CONDOMINIO", "V": datos[i][4].toString().trim() }, { "N": "PROYECTO", "V": datos[i][5].toString().trim() },
        { "N": "VIGENCIA", "V": vigenciaTexto }, { "N": "REFERENCIA_MANT", "V": datos[i][6].toString().trim() }
      ]
    };

    if (medioElegido.includes("CORREO") || medioElegido === "AMBOS") {
      if (!estatusActual.includes("✅ Correo") && !estatusActual.includes("❌ Correo")) {
        pendientes.push({ tipo: "1", fila: i + 5, payload: basePayload });
      } else { procesadosExitososOConErrorRevisado++; }
    }
    if (medioElegido.includes("WHATSAPP") || medioElegido === "AMBOS") {
      if (!estatusActual.includes("✅ Whats") && !estatusActual.includes("❌ Whats")) {
        pendientes.push({ tipo: "3", fila: i + 5, payload: basePayload });
      } else { procesadosExitososOConErrorRevisado++; }
    }
  }

  var totalGeneral = procesadosExitososOConErrorRevisado + pendientes.length;
  if (pendientes.length === 0) return { done: true, pct: 100, msg: "Todos los envíos fueron gestionados.", titulo: "Proceso Terminado" };

  var token = obtenerTokenGPH();
  if (!token) return { done: true, pct: 0, msg: "Tu usuario no fue autorizado.", titulo: "Error de Servidor" };

  var urlNotificacion = "[URL_ENDPOINT]";
  
  // MODO SEGURO: 5 registros 1x1
  var loteProcesar = pendientes.slice(0, 5); 

  for (var j = 0; j < loteProcesar.length; j++) {
    var item = loteProcesar[j];
    var idPlantillaAUsar = (item.tipo === "1") ? idPlantillaCorreo : idPlantillaWhats;
    var payloadFinal = { "clientes": JSON.stringify([item.payload]), "plantilla": idPlantillaAUsar, "lotes": "false", "medio": item.tipo };

    try {
      var resp = UrlFetchApp.fetch(urlNotificacion, { method: "post", headers: { "authorization": token }, payload: payloadFinal, muteHttpExceptions: true });
      var resJson = JSON.parse(resp.getContentText());
      var celda = hojaLocal.getRange(item.fila, 9);
      var textoPrevio = celda.getValue().toString();
      
      if (resp.getResponseCode() === 200 && resJson.status === 1) {
         var txtExito = (item.tipo === "1") ? "✅ Correo" : "✅ Whats";
         celda.setValue(textoPrevio ? textoPrevio + " | " + txtExito : txtExito);
      } else {
         var errServidor = resJson.mensaje || "Error Desconocido";

         if (errServidor.includes("insertar las notificaciones 1")) {
            var fallbackResult = intentarAutoSanacion(item, token, urlNotificacion, idPlantillaAUsar);
            if (fallbackResult.success) {
               celda.setValue(textoPrevio ? textoPrevio + " | " + fallbackResult.txt : fallbackResult.txt);
               continue; 
            } else { errServidor = fallbackResult.err; }
         }

         if (errServidor.includes("No se encontró información") || errServidor.includes("vacio")) {
           errServidor = (item.tipo === "1") ? "Sin correo en sistema" : "Sin celular en sistema";
         }
         var txtFallo = (item.tipo === "1") ? "❌ Correo: " + errServidor : "❌ Whats: " + errServidor;
         celda.setValue(textoPrevio ? textoPrevio + " | " + txtFallo : txtFallo);
      }
    } catch (e) {
      var celda = hojaLocal.getRange(item.fila, 9);
      var textoPrevio = celda.getValue().toString();
      var txtFallo = (item.tipo === "1") ? "❌ Correo: Error Red" : "❌ Whats: Error Red";
      celda.setValue(textoPrevio ? textoPrevio + " | " + txtFallo : txtFallo);
    }
    Utilities.sleep(500); 
  }
  
  SpreadsheetApp.flush(); 
  var avanceFinal = procesadosExitososOConErrorRevisado + loteProcesar.length;
  var pct = Math.floor((avanceFinal / totalGeneral) * 100);

  return { done: false, pct: pct, msg: "Transmitiendo " + avanceFinal + " de " + totalGeneral + " en Modo Seguro...", titulo: "Procesando" };
}


// ==========================================
// 🚀 BACKEND - MODO URGENTE (LOTES MASIVOS DE 200)
// ==========================================
function fase2BackendMasivo() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaLocal = libro.getSheetByName("Envios_Masivos");
  
  var pCorreo = hojaLocal.getRange("B1").getValue().toString().trim();
  var pWhats = hojaLocal.getRange("D1").getValue().toString().trim();
  var vig = hojaLocal.getRange("B2").getValue();

  var idPlantillaCorreo = pCorreo ? pCorreo.split(" - ")[0] : null;
  var idPlantillaWhats = pWhats ? pWhats.split(" - ")[0] : null;
  var vigenciaTexto = (vig instanceof Date) ? Utilities.formatDate(vig, Session.getScriptTimeZone(), "dd-MM-yyyy") : vig.toString().trim().replace(/\//g, "-");

  var datos = hojaLocal.getRange(5, 1, hojaLocal.getLastRow() - 4, 9).getValues();
  var colaCorreo = [], colaWhats = [], procesados = 0;

  for (var i = 0; i < datos.length; i++) {
    var viviendaCompleta = datos[i][0].toString().trim(), nombreCl = datos[i][1].toString().trim();
    var medioElegido = datos[i][7].toString().trim(), estatusActual = datos[i][8].toString().trim();

    if (!viviendaCompleta || nombreCl === "NO ENCONTRADO" || !medioElegido) continue;
    
    var basePayload = {
      "idVivienda": parseInt(datos[i][3]), "idCobranza": datos[i][2].toString().trim(), "nombre": nombreCl,
      "campos":[
        { "N": "CLIENTE", "V": nombreCl }, { "N": "LOTE", "V": viviendaCompleta.split("-").pop() },
        { "N": "CONDOMINIO", "V": datos[i][4].toString().trim() }, { "N": "PROYECTO", "V": datos[i][5].toString().trim() },
        { "N": "VIGENCIA", "V": vigenciaTexto }, { "N": "REFERENCIA_MANT", "V": datos[i][6].toString().trim() }
      ]
    };

    if (medioElegido.includes("CORREO") || medioElegido === "AMBOS") {
      if (!estatusActual.includes("✅ Correo") && !estatusActual.includes("❌ Correo")) {
        colaCorreo.push({ fila: i + 5, payload: basePayload });
      } else { procesados++; }
    }
    if (medioElegido.includes("WHATSAPP") || medioElegido === "AMBOS") {
      if (!estatusActual.includes("✅ Whats") && !estatusActual.includes("❌ Whats")) {
        colaWhats.push({ fila: i + 5, payload: basePayload });
      } else { procesados++; }
    }
  }

  var pendientesTotal = colaCorreo.length + colaWhats.length;
  var totalGeneral = procesados + pendientesTotal;
  
  if (pendientesTotal === 0) return { done: true, pct: 100, msg: "Todos los envíos masivos fueron gestionados.", titulo: "Proceso Terminado" };

  var token = obtenerTokenGPH();
  if (!token) return { done: true, pct: 0, msg: "Tu usuario no fue autorizado.", titulo: "Error de Servidor" };

  var urlNotificacion = "[URL_ENDPOINT]";
  
  // MODO URGENTE: Tomar hasta 200 de la cola para mandarlos en 1 solo golpe
  var loteProcesar;
  var tipoEnvio;
  var idPlantillaAUsar;

  if (colaCorreo.length > 0) {
    loteProcesar = colaCorreo.slice(0, 200);
    tipoEnvio = "1";
    idPlantillaAUsar = idPlantillaCorreo;
  } else {
    loteProcesar = colaWhats.slice(0, 200);
    tipoEnvio = "3";
    idPlantillaAUsar = idPlantillaWhats;
  }

  var clientesArray = loteProcesar.map(function(item) { return item.payload; });
  var payloadFinal = { "clientes": JSON.stringify(clientesArray), "plantilla": idPlantillaAUsar, "lotes": "false", "medio": tipoEnvio };

  try {
    var resp = UrlFetchApp.fetch(urlNotificacion, { method: "post", headers: { "authorization": token }, payload: payloadFinal, muteHttpExceptions: true });
    var resJson = JSON.parse(resp.getContentText());
    
    var txtResultado = "";
    if (resp.getResponseCode() === 200 && resJson.status === 1) {
       txtResultado = (tipoEnvio === "1") ? "✅ Correo" : "✅ Whats";
    } else {
       var errServidor = resJson.mensaje || "Error Desconocido";
       if (errServidor.includes("No se encontró información") || errServidor.includes("vacio")) {
         errServidor = (tipoEnvio === "1") ? "Sin correo en sistema" : "Sin celular en sistema";
       }
       txtResultado = (tipoEnvio === "1") ? "❌ Correo: " + errServidor : "❌ Whats: " + errServidor;
    }

    // Aplicar a los 200 de golpe
    loteProcesar.forEach(function(item) {
       var celda = hojaLocal.getRange(item.fila, 9);
       var textoPrevio = celda.getValue().toString();
       if (!textoPrevio.includes(txtResultado)) {
         celda.setValue(textoPrevio ? textoPrevio + " | " + txtResultado : txtResultado);
       }
    });

  } catch (e) {
    var txtFallo = (tipoEnvio === "1") ? "❌ Correo: Error Red/Timeout" : "❌ Whats: Error Red/Timeout";
    loteProcesar.forEach(function(item) {
       var celda = hojaLocal.getRange(item.fila, 9);
       var textoPrevio = celda.getValue().toString();
       if (!textoPrevio.includes(txtFallo)) {
         celda.setValue(textoPrevio ? textoPrevio + " | " + txtFallo : txtFallo);
       }
    });
  }
  
  SpreadsheetApp.flush(); 
  Utilities.sleep(1000); // Pequeña pausa de 1 seg para que el servidor trague el mega paquete
  
  var avanceFinal = procesados + loteProcesar.length;
  var pct = Math.floor((avanceFinal / totalGeneral) * 100);
  var labelMedio = (tipoEnvio === "1") ? "Correos" : "WhatsApp";

  return { done: false, pct: pct, msg: "Lote Urgente: Enviando bloque de " + loteProcesar.length + " " + labelMedio + "...", titulo: "Modo Urgente Activo" };
}


// ==========================================
// 🩹 FUNCIÓN AUXILIAR: EL MOTOR DE AUTO-SANACIÓN (Solo para el modo seguro)
// ==========================================
function intentarAutoSanacion(item, token, urlNotificacion, idPlantillaAUsar) {
   var combinaciones = [
      ["CLIENTE", "LOTE", "CONDOMINIO", "PROYECTO", "REFERENCIA_MANT"], 
      ["CLIENTE", "LOTE", "CONDOMINIO", "PROYECTO", "VIGENCIA"], 
      ["CLIENTE", "LOTE", "CONDOMINIO", "PROYECTO"] 
   ];

   for (var i = 0; i < combinaciones.length; i++) {
       var camposFiltrados = item.payload.campos.filter(function(c) { return combinaciones[i].indexOf(c.N) !== -1; });
       var payloadPrueba = JSON.parse(JSON.stringify(item.payload)); 
       payloadPrueba.campos = camposFiltrados;

       var payloadFinal = { "clientes": JSON.stringify([payloadPrueba]), "plantilla": idPlantillaAUsar, "lotes": "false", "medio": item.tipo };

       try {
         var resp = UrlFetchApp.fetch(urlNotificacion, { method: "post", headers: { "authorization": token }, payload: payloadFinal, muteHttpExceptions: true });
         var resJson = JSON.parse(resp.getContentText());
         
         if (resp.getResponseCode() === 200 && resJson.status === 1) {
             return { success: true, txt: (item.tipo === "1" ? "✅ Correo" : "✅ Whats") };
         }
         var errMsg = resJson.mensaje || "";
         if (!errMsg.includes("insertar las notificaciones 1")) {
             return { success: false, err: errMsg };
         }
       } catch(e) {}
   }
   return { success: false, err: "Ocurrió un error al insertar las notificaciones (Faltan variables en la macro)" };
}

// ==========================================
// 🧹 FASE 3: ARCHIVAR Y LIMPIAR
// ==========================================
function fase3LimpiarYArchivar() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaPrincipal = libro.getSheetByName("Envios_Masivos");
  var ui = SpreadsheetApp.getUi();

  if (!hojaPrincipal) return ui.alert("Error", "No se encontró la hoja principal.", ui.ButtonSet.OK);
  var ultimaFila = hojaPrincipal.getLastRow();
  if (ultimaFila < 5) return ui.alert("Aviso", "No hay datos para archivar.", ui.ButtonSet.OK);

  var hojaHistorico = libro.getSheetByName("Historico_envios");
  if (!hojaHistorico) {
    hojaHistorico = libro.insertSheet("Historico_envios");
    var cabeceras =["Lote (Vivienda)", "Nombre Cliente", "ID Cobranza", "ID Vivienda", "Condominio", "Proyecto", "Ref. Mantenimiento", "Medio", "Estatus Envío", "Correo de Envío"];
    hojaHistorico.getRange(1, 1, 1, cabeceras.length).setValues([cabeceras]).setFontWeight("bold").setBackground("#4c1130").setFontColor("white");
    hojaHistorico.setFrozenRows(1);
  }

  var datos = hojaPrincipal.getRange(5, 1, ultimaFila - 4, 9).getValues();
  var datosAArchivar =[];
  var correoUsuario = Session.getActiveUser().getEmail() || "Usuario Script";

  for (var i = 0; i < datos.length; i++) {
    if (datos[i][0].toString().trim() !== "" && datos[i][8].toString().trim() !== "") {
      var filaCopia = datos[i].slice();
      filaCopia.push(correoUsuario);   
      datosAArchivar.push(filaCopia);
    }
  }

  if (datosAArchivar.length > 0) {
    hojaHistorico.getRange(hojaHistorico.getLastRow() + 1, 1, datosAArchivar.length, 10).setValues(datosAArchivar);
  }

  hojaPrincipal.getRange(5, 1, ultimaFila - 4, 9).clearContent();
  ui.alert("✅ Limpieza Exitosa", "Se archivaron " + datosAArchivar.length + " registros históricos. Hoja lista.", ui.ButtonSet.OK);
}


// ==========================================
// 🎨 INTERFACES HTML: SELECTOR Y MATRIX
// ==========================================

const SELECTOR_HTML = `
<!DOCTYPE html>
<html>
<head>
  <style>
     body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f4f7f6; padding: 20px; text-align: center; }
     h2 { color: #333; margin-top: 0; }
     p { color: #555; font-size: 14px; margin-bottom: 25px; }
     .btn { display: block; width: 100%; padding: 15px; margin-bottom: 15px; border: none; border-radius: 6px; font-size: 16px; font-weight: bold; cursor: pointer; transition: 0.2s; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
     .btn-seguro { background: #28a745; color: white; }
     .btn-seguro:hover { background: #218838; }
     .btn-urgente { background: #dc3545; color: white; }
     .btn-urgente:hover { background: #c82333; }
     .sub { font-size: 11px; font-weight: normal; display: block; margin-top: 5px; opacity: 0.9; }
  </style>
</head>
<body>
  <h2>Tienes {{TOTAL}} envíos pendientes</h2>
  <p>Selecciona el motor de procesamiento que deseas utilizar:</p>

  <button class="btn btn-seguro" onclick="lanzar('fase2_seguro')">
     🛡️ MODO SEGURO (1 a 1)
     <span class="sub">Recomendado. Lento pero 100% exacto para detectar errores.</span>
  </button>

  <button class="btn btn-urgente" onclick="lanzar('fase2_masivo')">
     🚀 MODO URGENTE (Lotes de 200)
     <span class="sub">Extra Rápido.(Acepta riesgo de falsos positivos).</span>
  </button>

  <script>
     function lanzar(modo) {
        document.body.innerHTML = '<h3 style="color:#333; margin-top:50px;">Iniciando motor, por favor espera...</h3>';
        google.script.run.withSuccessHandler(function() {
           // Se cerrará automáticamente
        }).lanzarMatrix(modo);
     }
  </script>
</body>
</html>
`;

const MATRIX_HTML = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body, html { margin: 0; padding: 0; width: 100%; height: 100%; background-color: #050505; color: #0f0; font-family: 'Courier New', Courier, monospace; overflow: hidden; }
    canvas { display: block; position: absolute; top: 0; left: 0; z-index: 1; }
    .overlay {
      position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); z-index: 2;
      background: rgba(0, 15, 0, 0.85); border: 2px solid #0f0; padding: 40px; border-radius: 8px;
      text-align: center; width: 85%; max-width: 500px; box-shadow: 0 0 25px rgba(0, 255, 0, 0.5);
      backdrop-filter: blur(3px);
    }
    .title { font-size: 24px; font-weight: bold; margin-bottom: 25px; text-shadow: 0 0 10px #0f0; letter-spacing: 1px; }
    .progress-container { width: 100%; background: #002200; border: 1px solid #0f0; height: 30px; border-radius: 4px; position: relative; margin-bottom: 20px; box-shadow: inset 0 0 10px #000; }
    .progress-bar { width: 0%; height: 100%; background: #0f0; box-shadow: 0 0 15px #0f0; transition: width 0.3s ease; }
    .progress-text { position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); color: #fff; font-weight: bold; text-shadow: 1px 1px 2px #000; font-size: 14px; }
    .status { font-size: 15px; opacity: 0.9; text-shadow: 0 0 3px #0f0; min-height: 20px; }
    .btn-close {
      background: #000; color: #0f0; border: 1px solid #0f0; padding: 12px 25px; font-family: inherit; font-size: 16px;
      cursor: pointer; margin-top: 25px; text-transform: uppercase; font-weight: bold; letter-spacing: 1px;
      box-shadow: 0 0 10px #0f0; transition: all 0.3s ease; display: none; margin-left: auto; margin-right: auto;
    }
    .btn-close:hover { background: #0f0; color: #000; box-shadow: 0 0 20px #0f0; }
  </style>
</head>
<body>
  <canvas id="matrix"></canvas>
  <div class="overlay">
    <div class="title" id="title">Iniciando Enlace...</div>
    <div class="progress-container" id="p-cont">
      <div class="progress-bar" id="bar"></div>
      <div class="progress-text" id="pct">0%</div>
    </div>
    <div class="status" id="status">Preparando conexión con el servidor...</div>
    <button class="btn-close" id="btn-ok" onclick="google.script.host.close()">CERRAR</button>
  </div>

  <script>
    const canvas = document.getElementById('matrix');
    const ctx = canvas.getContext('2d');
    canvas.width = window.innerWidth; canvas.height = window.innerHeight;
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789@#$%^&*()'.split('');
    const fontSize = 14; const columns = canvas.width / fontSize;
    const drops =[]; for(let x=0; x<columns; x++) drops[x] = 1;
    function draw() {
      ctx.fillStyle = 'rgba(0, 0, 0, 0.07)'; ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.fillStyle = '#0F0'; ctx.font = fontSize + 'px monospace';
      for(let i=0; i<drops.length; i++){
        const text = chars[Math.floor(Math.random() * chars.length)];
        ctx.fillText(text, i * fontSize, drops[i] * fontSize);
        if(drops[i] * fontSize > canvas.height && Math.random() > 0.975) drops[i] = 0;
        drops[i]++;
      }
    }
    setInterval(draw, 35);

    const faseActiva = '{{FASE}}';
    document.getElementById('title').innerText = faseActiva === 'fase1' ? 'SISTEMA: EXTRACCIÓN DE CARTERA' : 'SISTEMA: TRANSMISIÓN MASIVA';

    window.onload = function() {
      ejecutarSiguienteLote();
    };

    function ejecutarSiguienteLote() {
      if (faseActiva === 'fase1') {
         google.script.run.withSuccessHandler(onFinish).withFailureHandler(onError).fase1Backend();
      } else if (faseActiva === 'fase2_seguro') {
         google.script.run.withSuccessHandler(procesarRespuestaFase2).withFailureHandler(onError).fase2BackendSeguro();
      } else if (faseActiva === 'fase2_masivo') {
         google.script.run.withSuccessHandler(procesarRespuestaFase2).withFailureHandler(onError).fase2BackendMasivo();
      }
    }

    function procesarRespuestaFase2(res) {
      if (!res) return onError({message: "Interrupción de conexión con Google."});
      
      document.getElementById('bar').style.width = res.pct + '%';
      document.getElementById('pct').innerText = res.pct + '%';
      document.getElementById('status').innerText = res.msg;

      if (!res.done) {
        ejecutarSiguienteLote();
      } else {
        onFinish(res);
      }
    }

    function onFinish(res) {
      document.getElementById('bar').style.width = '100%';
      document.getElementById('pct').innerText = '100%';
      setTimeout(() => {
        document.getElementById('p-cont').style.display = 'none';
        document.getElementById('title').innerText = res.titulo;
        document.getElementById('status').innerText = res.msg;
        document.getElementById('btn-ok').style.display = 'block';
      }, 800);
    }

    function onError(err) {
      document.getElementById('title').innerText = "CONEXIÓN INTERRUMPIDA";
      document.getElementById('title').style.color = "#f00";
      
      let mensajeError = (err && err.message) ? err.message : (typeof err === 'string' ? err : "Micro-corte de red detectado.\\n\\nCierra esta ventana y vuelve a dar clic en 'Fase 2' para continuar el envío donde se quedó.");
      
      document.getElementById('status').innerText = mensajeError;
      document.getElementById('btn-ok').style.display = 'block';
      document.getElementById('btn-ok').style.borderColor = "#f00";
      document.getElementById('btn-ok').style.color = "#f00";
    }
  </script>
</body>
</html>
`;