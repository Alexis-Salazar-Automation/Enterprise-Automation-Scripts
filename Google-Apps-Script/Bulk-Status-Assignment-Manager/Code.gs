const ID_ARCHIVO_CARTERA = 'ID_CARTERA';
const NOMBRE_HOJA_CARTERA = 'NOMBRE_HOJA';

// --- CATÁLOGOS UNIFICADOS ---
const CATALOGO_ESTATUS =[LISTA DE TODOS LOS PROYECTOS VÁLIDOS A DESCARGAR];

// ==========================================
// 🚀 GATILLO MÁGICO: SELECCIÓN MÚLTIPLE SIN ERROR ROJO
// ==========================================
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() === "Operacion_Masiva" && e.range.getColumn() === 6 && e.range.getRow() > 1) {
    var newValue = e.value; var oldValue = e.oldValue;
    if (!newValue) return; 
    if (oldValue && oldValue !== newValue) {
      if (oldValue.indexOf(newValue) === -1) { 
        var combo = oldValue + ", " + newValue; e.range.setValue(combo); 
        var catTemp = CATALOGO_ESTATUS.slice(); catTemp.push(combo);
        e.range.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(catTemp, true).build());
      } else { e.range.setValue(oldValue); }
    }
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🤖 Súper Bot Cartera')
      .addItem('🔍 1. Validar Lotes y Extraer Datos', 'abrirMatrixFase1')
      .addItem('🚀 2. Ejecutar Operación Unificada', 'abrirMatrixFase2')
      .addSeparator()
      .addItem('🕵️‍♂️ 3. Auditoría Forense', 'abrirMatrixFase3')
      .addToUi();
}

// NUEVA VERSIÓN: DETECCIÓN DINÁMICA DE USUARIO
function obtenerToken() {
  // Se obtiene el correo de quien esté usando la hoja (con fallback de seguridad)
  var correoActivo = Session.getActiveUser().getEmail() || "[correo_electronico]";
  // Se deriva un nombre a partir del correo
  var nombreDerivado = correoActivo.split('@')[0].replace(/\./g, ' ').toUpperCase();
  
  var credencialesDinamicas = { "googleId": "0000000000000", "email": correoActivo, "name": nombreDerivado };
  var url = "https://api.your-company.com/index.php";
  
  try {
    var resp = UrlFetchApp.fetch(url, { 
      "method": "post", 
      "contentType": "application/json", 
      "payload": JSON.stringify({"data": credencialesDinamicas}), 
      "muteHttpExceptions": true 
    });
    var jsonResp = JSON.parse(resp.getContentText());
    if (jsonResp && jsonResp.data && jsonResp.data.token) {
      return jsonResp.data.token;
    } else {
      throw new Error("Credenciales rechazadas.");
    }
  } catch(e) { return null; }
}

function getBackendStatus() {
  return CacheService.getUserCache().get('bot_status');
}

function abrirMatrixFase1() { var h = HtmlService.createTemplateFromFile('Matrix'); h.fase = 1; SpreadsheetApp.getUi().showModalDialog(h.evaluate().setWidth(600).setHeight(400), 'Terminal ETL: Unificada'); }
function abrirMatrixFase2() { var h = HtmlService.createTemplateFromFile('Matrix'); h.fase = 2; SpreadsheetApp.getUi().showModalDialog(h.evaluate().setWidth(600).setHeight(400), 'Terminal Ejecución: Unificada'); }
function abrirMatrixFase3() { var h = HtmlService.createTemplateFromFile('Matrix'); h.fase = 3; SpreadsheetApp.getUi().showModalDialog(h.evaluate().setWidth(600).setHeight(400), 'Terminal Auditoría'); }

// ==========================================
// FASE 1: ETL UNIFICADO
// ==========================================
function backendFase1() {
  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Operacion_Masiva");
    if (hoja.getLastRow() <= 1) return { error: "No hay lotes en la columna A." };
    
    var lotesUsuario = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues();
    var hojaMaestra = SpreadsheetApp.openById(ID_ARCHIVO_CARTERA).getSheetByName(NOMBRE_HOJA_CARTERA);
    var datosMaestros = hojaMaestra.getDataRange().getValues();
    var enc = datosMaestros[0];
    
    var diccCartera = {};
    for (var i = 1; i < datosMaestros.length; i++) {
      var loteStr = datosMaestros[i][enc.indexOf("vivienda")] ? datosMaestros[i][enc.indexOf("vivienda")].toString().trim() : "";
      if (loteStr) diccCartera[loteStr] = { nom: datosMaestros[i][enc.indexOf("clienteNombre")], idCob: datosMaestros[i][enc.indexOf("id_cobranza")], est: datosMaestros[i][enc.indexOf("idEstatus")], ase: datosMaestros[i][enc.indexOf("asesorNombre")] };
    }
    
    var resB=[], resC=[], resD=[], resE=[];
    for (var k = 0; k < lotesUsuario.length; k++) {
      var l = lotesUsuario[k][0].toString().trim();
      if (!l) { resB.push([""]); resC.push([""]); resD.push([""]); resE.push([""]); continue; }
      if (diccCartera[l]) { resB.push([diccCartera[l].nom]); resC.push([diccCartera[l].idCob]); resD.push([diccCartera[l].est]); resE.push([diccCartera[l].ase]); } 
      else { resB.push(["NO ENCONTRADO"]); resC.push(["-"]); resD.push(["-"]); resE.push(["-"]); }
    }
    
    hoja.getRange(2, 2, resB.length, 1).setValues(resB); hoja.getRange(2, 3, resC.length, 1).setValues(resC);
    hoja.getRange(2, 4, resD.length, 1).setValues(resD); hoja.getRange(2, 5, resE.length, 1).setValues(resE);
    
    hoja.getRange(2, 6, resB.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CATALOGO_ESTATUS, true).setAllowInvalid(true).build());
    hoja.getRange(2, 7, resB.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CATALOGO_SEDES, true).build());
    hoja.getRange(2, 8, resB.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(CATALOGO_ASESORES, true).build());
    
    return { success: "Datos cruzados correctamente." };
  } catch (e) { return { error: e.message }; }
}

// ==========================================
// FASE 2: EJECUCIÓN UNIFICADA CON REINTENTO INTELIGENTE
// ==========================================
function backendFase2() {
  var cache = CacheService.getUserCache();
  try {
    cache.put('bot_status', 'INICIANDO', 600); 
    
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hojaLocal = libro.getSheetByName("Operacion_Masiva");
    var datos = hojaLocal.getDataRange().getValues();
    if (datos.length <= 1) return { error: "La hoja está vacía." };

    var hojaMaestra = SpreadsheetApp.openById(ID_ARCHIVO_CARTERA).getSheetByName(NOMBRE_HOJA_CARTERA);
    var datosMaestros = hojaMaestra.getDataRange().getValues();
    var enc = datosMaestros[0];
    
    var diccObjetos = {};
    for (var i = 1; i < datosMaestros.length; i++) {
      var loteStr = datosMaestros[i][enc.indexOf("vivienda")] ? datosMaestros[i][enc.indexOf("vivienda")].toString().trim() : "";
      if (loteStr) {
        var objCliente = {};
        for(var c = 0; c < enc.length; c++) { objCliente[enc[c]] = (datosMaestros[i][c] === "" || datosMaestros[i][c] === undefined) ? null : datosMaestros[i][c]; }
        objCliente.isSeleccionado = true; objCliente.id = objCliente.id_asignacion ? parseInt(objCliente.id_asignacion) : null;
        if(objCliente.idVivienda) objCliente.idVivienda = parseInt(objCliente.idVivienda);
        if(objCliente.idEMant) objCliente.idEMant = parseInt(objCliente.idEMant);
        if(objCliente.idestviv) objCliente.idestviv = parseInt(objCliente.idestviv);
        diccObjetos[loteStr] = objCliente;
      }
    }

    // EXTRAE EL CORREO DINÁMICO PARA EL HISTORIAL (FALLBACK AL TUYO POR PRECAUCIÓN)
    var logData =[]; 
    var usuarioEjecutor = Session.getActiveUser().getEmail() || "aux.adm6@gph.mx"; 
    var fechaHora = new Date();
    
    var gruposEstatusV2 = {}; var grupoAsesores = {}; var obsPorAsesor = {};
    
    for (var j = 1; j < datos.length; j++) {
      var lote = datos[j][0].toString().trim(), nombre = datos[j][1].toString().trim();
      var idCobranza = datos[j][2].toString().trim(), estatusViejo = datos[j][3];
      var estNuevo = datos[j][5].toString().trim(), asesorNuevo = datos[j][7].toString().trim(), obs = datos[j][8].toString().trim();
      
      if (lote === "" || nombre.includes("NO ENCONTRADO") || !diccObjetos[lote]) continue;
      var cambioEstatus = (estNuevo !== ""); var cambioAsesor = (asesorNuevo !== "");
      if (!cambioEstatus && !cambioAsesor) continue; 

      if (cambioEstatus) {
        var partesEstatus = estNuevo.split(",");
        var idsList =[];
        for (var x = 0; x < partesEstatus.length; x++) { idsList.push(partesEstatus[x].split(" - ")[0].trim()); }
        var llaveCombinacion = idsList.sort().join(","); 
        if(!gruposEstatusV2[llaveCombinacion]) { gruposEstatusV2[llaveCombinacion] = { "idEstatus": idsList, "id_cobranza": [] }; }
        gruposEstatusV2[llaveCombinacion].id_cobranza.push(idCobranza);
        
        diccObjetos[lote].idEstatus = idsList.join(","); 
        // MODIFICADO: ACCIÓN PENÚLTIMA, USUARIO ÚLTIMO
        logData.push([fechaHora, lote, idCobranza, nombre, estatusViejo, idsList.join(","), "CAMBIO ESTATUS", usuarioEjecutor]);
      }
      if (cambioAsesor) {
        var idAsesorID = asesorNuevo.split(" - ")[0].trim();
        if (!grupoAsesores[idAsesorID]) { grupoAsesores[idAsesorID] =[]; obsPorAsesor[idAsesorID] = obs; }
        grupoAsesores[idAsesorID].push(diccObjetos[lote]);
        // MODIFICADO: ACCIÓN PENÚLTIMA, USUARIO ÚLTIMO
        logData.push([fechaHora, lote, idCobranza, nombre, datos[j][4], asesorNuevo, "REASIGNACIÓN", usuarioEjecutor]);
      }
    }

    var token = obtenerToken(); if (!token) throw new Error("Credenciales rechazadas. Contacta a sistemas.");
    var contEstatus = 0; var contAsig = 0;

    // --- EJECUCIÓN 1: ESTATUS ---
    var combosEstatus = Object.keys(gruposEstatusV2);
    if (combosEstatus.length > 0) {
      cache.put('bot_status', 'UPDATING_STATUS', 600); 
      var urlEstV2 = "https://api.your-company.com/index.php";
      for (var cb = 0; cb < combosEstatus.length; cb++) {
         var comboInfo = gruposEstatusV2[combosEstatus[cb]];
         for (var p = 0; p < comboInfo.id_cobranza.length; p += 200) {
            var rebIds = comboInfo.id_cobranza.slice(p, p + 200);
            var rEst = UrlFetchApp.fetch(urlEstV2, { "method": "post", "contentType": "application/json", "headers": {"authorization": token}, "payload": JSON.stringify({ "idEstatus": comboInfo.idEstatus, "id_cobranza": rebIds }), "muteHttpExceptions": true });
            
            var jsE = {}; try{ jsE = JSON.parse(rEst.getContentText()); }catch(e){}
            if (rEst.getResponseCode() === 200 && jsE.status === 1) contEstatus += rebIds.length; 
            else throw new Error("Fallo Estatus V2. " + rEst.getContentText().substring(0,100));
            Utilities.sleep(1000); 
         }
      }
    }

    // --- EJECUCIÓN 2: ASIGNACIÓN V2 (CON REINTENTO INTELIGENTE) ---
    var asesoresIds = Object.keys(grupoAsesores);
    if (asesoresIds.length > 0) {
      cache.put('bot_status', 'ASSIGNING', 600); 
      var urlAsig = "https://api-cobranza.gphsis.com/index.php/Clientes/crear_asignarClientesV2";
      
      for (var k = 0; k < asesoresIds.length; k++) {
        var idAse = asesoresIds[k]; var listaC = grupoAsesores[idAse];
        for (var m = 0; m < listaC.length; m += 200) {
          var rebAsig = listaC.slice(m, m + 200);
          var pAsig = { "id_usuario": idAse, "clientes": rebAsig, "observaciones": obsPorAsesor[idAse] };
          
          var success = false;
          var intentos = 0;
          var maxIntentos = 4;
          
          while (!success && intentos < maxIntentos) {
            var rAsi = UrlFetchApp.fetch(urlAsig, { "method": "post", "contentType": "application/json", "headers": {"authorization": token}, "payload": JSON.stringify(pAsig), "muteHttpExceptions": true });
            var jResp = {}; try { jResp = JSON.parse(rAsi.getContentText()); } catch(e) {}
            
            if (rAsi.getResponseCode() === 200 && jResp.status === true) { 
              success = true;
              contAsig += rebAsig.length; 
            } else { 
              intentos++;
              if (intentos >= maxIntentos) {
                 throw new Error("Asignación fallida tras " + maxIntentos + " intentos. Servidor ocupado. Resp: " + rAsi.getContentText().substring(0,100));
              }
              
              var tiempoEsperaSegundos = 30; 
              var finCooldown = new Date().getTime() + (tiempoEsperaSegundos * 1000);
              
              cache.put('bot_status', 'COOLDOWN|' + finCooldown, 600);
              Utilities.sleep(tiempoEsperaSegundos * 1000); 
              cache.put('bot_status', 'ASSIGNING', 600);
            }
          }
          Utilities.sleep(1000); 
        }
      }
    }

    if (logData.length > 0) {
      var hojaLog = libro.getSheetByName("Log_Historico");
      hojaLog.getRange(hojaLog.getLastRow() + 1, 1, logData.length, logData[0].length).setValues(logData);
    }
    hojaLocal.getRange(2, 1, hojaLocal.getLastRow(), 9).clearContent();
    hojaLocal.getRange(2, 6, hojaLocal.getLastRow(), 3).clearDataValidations();
    
    cache.remove('bot_status'); 
    return { success: "¡Éxito! " + contEstatus + " Estatus y " + contAsig + " Asignaciones realizadas con Reintento Inteligente." };
  } catch (e) { 
    CacheService.getUserCache().remove('bot_status'); 
    return { error: e.message }; 
  }
}

// ==========================================
// FASE 3: AUDITORÍA FORENSE
// ==========================================
function backendFase3() {
  try {
    var libro = SpreadsheetApp.getActiveSpreadsheet(); var hojaAudit = libro.getSheetByName("Buscador_Auditoria");
    var datosMaestros = SpreadsheetApp.openById(ID_ARCHIVO_CARTERA).getSheetByName(NOMBRE_HOJA_CARTERA).getDataRange().getValues();
    var enc = datosMaestros[0]; var dicc = {}; 
    for (var i = 1; i < datosMaestros.length; i++) {
      var lote = datosMaestros[i][enc.indexOf("vivienda")] ? datosMaestros[i][enc.indexOf("vivienda")].toString().trim() : ""; 
      if(lote) dicc[lote] = datosMaestros[i][enc.indexOf("idVivienda")];
    }
    var lotesA_Buscar = hojaAudit.getRange(2, 1, hojaAudit.getLastRow()-1, 1).getValues(); var token = obtenerToken(); var resultados =[];
    for (var i = 0; i < lotesA_Buscar.length; i++) {
      var lote = lotesA_Buscar[i][0].toString().trim(); if (lote === "" || !dicc[lote]) continue;
      var resp = UrlFetchApp.fetch("https://api.your-company.com/"; + dicc[lote], { "method": "get", "headers": {"authorization": token}, "muteHttpExceptions": true });
      var hist = JSON.parse(resp.getContentText());
      if (Array.isArray(hist) && hist.length > 0) {
        for (var j = 0; j < hist.length; j++) { var d = hist[j]; resultados.push([ lote, d.NombreCliente || "-", d.estatus || "-", d.usrCrea || "-", d.fecha_creacion || "-", d.usrUpdate || "-", d.fecha_modifica || "-", d.observacion || "-" ]); }
      } else { resultados.push([lote, "SIN HISTORIAL", "-", "-", "-", "-", "-", "-"]); }
      Utilities.sleep(300);
    }
    hojaAudit.getRange(2, 1, resultados.length, 8).setValues(resultados); return { success: "Auditoría completa." };
  } catch (e) { return { error: e.message }; }
}


function pruebaDeSeguridadAPI() {
  // 1. Ponemos un correo válido (puede ser el tuyo)
  var correoPrueba = "aux.adm6@gph.mx"; 
  
  // 2. Ponemos un Google ID completamente falso e inventado
  var idFalso = "999999999999999999999"; 
  var nombreFalso = "PRUEBA";

  var payload = { 
    "data": { 
      "googleId": idFalso, 
      "email": correoPrueba, 
      "name": nombreFalso 
    } 
  };

  var url = "https://api-cobranza.gphsis.com/index.php/User/verificar_gmail";

  try {
    var resp = UrlFetchApp.fetch(url, { 
      "method": "post", 
      "contentType": "application/json", 
      "payload": JSON.stringify(payload), 
      "muteHttpExceptions": true 
    });
    
    var textoRespuesta = resp.getContentText();
    var codigoEstado = resp.getResponseCode();
    var jsonResp = JSON.parse(textoRespuesta);
    
    // Evaluamos el resultado
    if (codigoEstado === 200 && jsonResp && jsonResp.data && jsonResp.data.token) {
      // 🚨 EL SISTEMA FUE VULNERADO
      SpreadsheetApp.getUi().alert(
        "🚨 FALLA DE SEGURIDAD CONFIRMADA 🚨\n\n" +
        "El servidor confió ciegamente. Nos dio un Token de acceso válido usando un ID completamente falso.\n\n" +
        "Token obtenido: " + jsonResp.data.token.substring(0, 30) + "..."
      );
    } else {
      // ✅ EL SISTEMA ES SEGURO
      SpreadsheetApp.getUi().alert(
        "✅ EL SISTEMA ES SEGURO ✅\n\n" +
        "El servidor rechazó el acceso porque detectó que el ID no es correcto.\n\n" +
        "Mensaje del servidor: " + textoRespuesta
      );
    }
    
  } catch(e) { 
    SpreadsheetApp.getUi().alert("Error al intentar conectar: " + e.message);
  }
}