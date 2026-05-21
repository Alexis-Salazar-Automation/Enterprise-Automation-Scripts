// =====================================================================
// VARIABLES GLOBALES Y MENÚ
// =====================================================================
const PAYLOAD_LOGIN = {
  "data": {
    "googleId": "PROTECTED_ID",
    "email": "user@example.com",
    "name": "DEVELOPER_NAME",
  }
};

const API_BASE_URL = "https://api.your-company.com/index.php";
const PUBLIC_SPREADSHEET_ID = "YOUR_PUBLIC_SHEET_ID";
const DRIVE_BACKUP_FOLDER_ID = "YOUR_DRIVE_FOLDER_ID";
const RECIPIENT_EMAILS = "admin@example.com,manager@example.com";

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 Automation Bot')
    .addItem('🚀 Open Manual Console', 'mostrarConsola')
    .addSeparator()
    .addItem('⚙️ Force Auto-Start', 'iniciarEjecucionAutomatica')
    .addToUi();
}
function mostrarConsola() {
  var html = HtmlService.createHtmlOutputFromFile('Consola')
      .setTitle('Robot Extractor Nacional PV')
      .setWidth(750)
      .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

// =====================================================================
// FUNCIONES NÚCLEO (API)
// =====================================================================
function loginGAS() {
  var url = "https://api-cobranza.gphsis.com/index.php/User/verificar_gmail";
  var options = { method: "post", contentType: "application/json", payload: JSON.stringify(PAYLOAD_LOGIN), muteHttpExceptions: true };
  var resp = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(resp.getContentText());
  if (json.status === 1) return json.data.token;
  throw new Error("Error en login");
}

function obtenerProyectosGAS(token) {
  var url = "https://api-cobranza.gphsis.com/index.php/Catalogo/PVProyectos";
  var options = { method: "post", contentType: "application/json", headers: { "authorization": token }, muteHttpExceptions: true };
  var resp = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(resp.getContentText());
  return json.proyectos ||[];
}

function procesarProyectoGAS(token, idProy, isFirst) {
  var urlCond = "https://api-cobranza.gphsis.com/index.php/PosVenta/Condominios/" + idProy;
  var urlViv = "https://api-cobranza.gphsis.com/index.php/PosVenta/Viviendasxpersona";
  var optsBase = { method: "post", contentType: "application/json", headers: { "authorization": token }, muteHttpExceptions: true };
  
  optsBase.payload = JSON.stringify({ "idResidencial": idProy });
  var respCond = UrlFetchApp.fetch(urlCond, optsBase);
  var condominios =[];
  try { condominios = JSON.parse(respCond.getContentText()); } catch(e) {}
  
  if (!Array.isArray(condominios) || condominios.length === 0) return 0;

  var todas_viviendas =[];
  
  for (var i = 0; i < condominios.length; i++) {
    if (!condominios[i].idProyCond) continue;
    optsBase.payload = JSON.stringify({ "idProyCond": condominios[i].idProyCond, "npersona": "" });
    var respViv = UrlFetchApp.fetch(urlViv, optsBase);
    try {
      var jsonViv = JSON.parse(respViv.getContentText());
      if (jsonViv.viviendas && Array.isArray(jsonViv.viviendas)) {
        todas_viviendas = todas_viviendas.concat(jsonViv.viviendas);
      }
    } catch(e) {}
  }

  if (todas_viviendas.length === 0) return 0;

  pegarEnHoja(todas_viviendas, isFirst);
  return todas_viviendas.length;
}

/*function pegarEnHoja(viviendasArray, isFirst) {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = libro.getSheetByName("Cartera_PV");
  if (!sheet) sheet = libro.insertSheet("Cartera_PV");

  var props = PropertiesService.getScriptProperties();
  var headers;

  if (isFirst) {
    // 🔥 Guardamos cuántas filas había AYER antes de borrar todo
    var filasAntes = sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 0;
    props.setProperty("FILAS_AYER", filasAntes.toString());

    sheet.clear();
    headers = Object.keys(viviendasArray[0]);
    props.setProperty("CABECERAS", JSON.stringify(headers));
    sheet.appendRow(headers); 
    SpreadsheetApp.flush(); 
  } else {
    var storedHeaders = props.getProperty("CABECERAS");
    headers = storedHeaders ? JSON.parse(storedHeaders) : Object.keys(viviendasArray[0]);
  }

  var datos2D = viviendasArray.map(function(obj) {
    return headers.map(function(h) { return (obj[h] !== undefined && obj[h] !== null) ? obj[h] : ""; });
  });

  var chunkSize = 10000; 
  var maxIntentos = 3;   
  var ultimaFila = sheet.getLastRow();

  for (var i = 0; i < datos2D.length; i += chunkSize) {
    var bloque = datos2D.slice(i, i + chunkSize);
    var exito = false;
    var intentos = 0;

    while (!exito && intentos < maxIntentos) {
      try {
        sheet.getRange(ultimaFila + 1, 1, bloque.length, headers.length).setValues(bloque);
        SpreadsheetApp.flush(); 
        ultimaFila += bloque.length; 
        exito = true;
        Utilities.sleep(1500); 
      } catch(e) {
        intentos++;
        Utilities.sleep(6000); 
        if (intentos >= maxIntentos) throw new Error("Fallo al subir bloque.");
      }
    }
  }
}*/

function pegarEnHoja(viviendasArray, isFirst) {
  // 1. EL LIBRO MAESTRO (Donde corre el script, privado para ti)
  var libroMaster = SpreadsheetApp.getActiveSpreadsheet();
  var hojaMaster = libroMaster.getSheetByName("Cartera_PV");
  if (!hojaMaster) hojaMaster = libroMaster.insertSheet("Cartera_PV");

  // 2. EL LIBRO PÚBLICO (El que le vas a compartir al equipo)
  var idPublico = "1Z6zB6IZGMzQQDIcIEnFz8CTSTbwBAWOX6I2ORM9L6us";
  var libroPublico = SpreadsheetApp.openById(idPublico);
  var hojaPublica = libroPublico.getSheetByName("Cartera_PV");
  if (!hojaPublica) hojaPublica = libroPublico.insertSheet("Cartera_PV");

  var props = PropertiesService.getScriptProperties();
  var headers;

  if (isFirst) {
    // Guardar cuántas filas tenía el Maestro ayer
    var filasAntes = hojaMaster.getLastRow() > 1 ? hojaMaster.getLastRow() - 1 : 0;
    props.setProperty("FILAS_AYER", filasAntes.toString());

    // Limpiar ambas hojas
    hojaMaster.clear();
    hojaPublica.clear();

    headers = Object.keys(viviendasArray[0]);
    props.setProperty("CABECERAS", JSON.stringify(headers));
    
    // Pegar títulos en ambas
    hojaMaster.appendRow(headers); 
    hojaPublica.appendRow(headers);
    SpreadsheetApp.flush(); 
  } else {
    var storedHeaders = props.getProperty("CABECERAS");
    headers = storedHeaders ? JSON.parse(storedHeaders) : Object.keys(viviendasArray[0]);
  }

  var datos2D = viviendasArray.map(function(obj) {
    return headers.map(function(h) { return (obj[h] !== undefined && obj[h] !== null) ? obj[h] : ""; });
  });

  var chunkSize = 10000; 
  var maxIntentos = 3;   
  var ultimaFila = hojaMaster.getLastRow();

  for (var i = 0; i < datos2D.length; i += chunkSize) {
    var bloque = datos2D.slice(i, i + chunkSize);
    var exito = false;
    var intentos = 0;

    while (!exito && intentos < maxIntentos) {
      try {
        // Pegar el bloque en AMBOS archivos en simultáneo
        hojaMaster.getRange(ultimaFila + 1, 1, bloque.length, headers.length).setValues(bloque);
        hojaPublica.getRange(ultimaFila + 1, 1, bloque.length, headers.length).setValues(bloque);
        
        SpreadsheetApp.flush(); 
        ultimaFila += bloque.length; 
        exito = true;
        Utilities.sleep(1500); 
      } catch(e) {
        intentos++;
        Utilities.sleep(6000); 
        if (intentos >= maxIntentos) throw new Error("Fallo al subir bloque en los libros.");
      }
    }
  }
}

// =====================================================================
// RESPALDO EN DRIVE Y ENVÍO DE CORREO HTML
// =====================================================================
function generarRespaldoYEnviarCorreo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Cartera_PV");
  var filasHoy = sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 0;
  
  var props = PropertiesService.getScriptProperties();
  var filasAyer = parseInt(props.getProperty("FILAS_AYER") || filasHoy);
  var diff = filasHoy - filasAyer;

  // 1. Exportar la hoja actual como archivo Excel (.xlsx)
  var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx";
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
  
  var fechaStr = Utilities.formatDate(new Date(), "GMT-6", "dd-MM-yyyy");
  var fileName = "Respaldo_Cartera_" + fechaStr + ".xlsx";
  var blob = response.getBlob().setName(fileName);

  // 2. Guardar en la carpeta compartida de Google Drive
  var folderId = "0AEchVV7Q5NCVUk9PVA"; // Tu ID de carpeta
  var folder = DriveApp.getFolderById(folderId);
  var newFile = folder.createFile(blob);
  var downloadUrl = newFile.getUrl();

  // 3. Crear el formato del correo (IDÉNTICO A TU IMAGEN)
  var colorDiff = diff > 0 ? "#28a745" : (diff < 0 ? "#dc3545" : "#555");
  var signoDiff = diff > 0 ? "+" : "";

    var idPublico = "1Z6zB6IZGMzQQDIcIEnFz8CTSTbwBAWOX6I2ORM9L6us";
  var urlPublica = "https://docs.google.com/spreadsheets/d/" + idPublico + "/edit";

  var htmlBody = `
  <div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden; background-color: #f6f8fa;">
    <div style="background-color: #0b3d60; color: white; padding: 25px; text-align: center;">
      <h2 style="margin: 0;">📊 Reporte de Cartera</h2>
      <p style="margin: 5px 0 0 0; font-size: 14px;">Sincronización automatizada completada</p>
    </div>
    <div style="padding: 30px; color: #333; background-color: white;">
      <p>Hola Alexis,</p>
      <p>El robot automatizado ha actualizado la cartera en la nube de manera exitosa y el respaldo ha sido guardado en la carpeta compartida. Aquí tienes el resumen del día:</p>
      
      <div style="background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 8px; padding: 15px; margin: 20px 0;">
        <table style="width: 100%; border-collapse: collapse;">
          <tr style="border-bottom: 1px dotted #ccc;">
            <td style="padding: 10px 0; font-weight: bold; color: #555;">Total Filas (Ayer):</td>
            <td style="padding: 10px 0; text-align: right; color: #555;">${filasAyer.toLocaleString()}</td>
          </tr>
          <tr style="border-bottom: 1px dotted #ccc;">
            <td style="padding: 10px 0; font-weight: bold; color: #333;">Total Filas (Hoy):</td>
            <td style="padding: 10px 0; text-align: right; font-weight: bold; color: #333;">${filasHoy.toLocaleString()}</td>
          </tr>
          <tr>
            <td style="padding: 10px 0; font-weight: bold; color: #555;">Diferencia Neta:</td>
            <td style="padding: 10px 0; text-align: right; color: ${colorDiff}; font-weight: bold;">${signoDiff}${diff.toLocaleString()}</td>
          </tr>
        </table>
      </div>

      <div style="text-align: center; margin-top: 30px; margin-bottom: 20px;">
        <a href="${urlPublica}" style="background-color: #2ea44f; color: white; padding: 12px 20px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 14px; margin-right: 10px; display: inline-block;">📊 Ver en Google Sheets</a>
        <a href="${downloadUrl}" style="background-color: #0366d6; color: white; padding: 12px 20px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 14px; display: inline-block;">💾 Ver Respaldo Drive</a>
      </div>
    </div>
    
    <div style="background-color: #f6f8fa; padding: 20px; text-align: center; border-top: 1px solid #e0e0e0;">
      <p style="font-size: 11px; color: #888; margin: 0; line-height: 1.5;">
        Generado por Google Apps Script • Tiempo: ${new Date().toLocaleTimeString()}<br>
        Este es un mensaje automático, por favor no respondas a este correo.
      </p>
    </div>
  </div>
  `;


  // 4. Enviar el correo final
  /*MailApp.sendEmail({
    to: "aux.adm6@gph.mx",
    subject: "✅ Cartera Post Venta Actualizada Exitosamente",
    htmlBody: htmlBody,
    attachments: [blob]
  });*/


  // 4. Enviar el correo final con Timestamp
  var timestampAsunto = Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy HH:mm");
  
  MailApp.sendEmail({
    to: "aux.adm6@gph.mx,aux.adm4@gph.mx,aux.adm2@gph.mx",
    subject: "✅ Cartera Post Venta Actualizada Exitosamente 📅 " + timestampAsunto,
    htmlBody: htmlBody,
    attachments: [blob]
  });
}

// =====================================================================
// MODO AUTOMÁTICO (CON AUTO-REANUDACIÓN DE TIEMPO)
// =====================================================================
function iniciarEjecucionAutomatica() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty("AUTO_INDEX");
  limpiarTriggersTemporales();
  continuarEjecucion();
}

function continuarEjecucion() {
  var startTime = new Date().getTime();
  var props = PropertiesService.getScriptProperties();
  var index = parseInt(props.getProperty("AUTO_INDEX") || "0");
  var token = props.getProperty("AUTO_TOKEN");
  var proyectosStr = props.getProperty("AUTO_PROYECTOS");
  
  if (index === 0) {
    token = loginGAS();
    var proyectos = obtenerProyectosGAS(token);
    props.setProperty("AUTO_TOKEN", token);
    props.setProperty("AUTO_PROYECTOS", JSON.stringify(proyectos));
    proyectosStr = JSON.stringify(proyectos);
  }
  
  var proyectos = JSON.parse(proyectosStr);
  
  for (var i = index; i < proyectos.length; i++) {
    if (new Date().getTime() - startTime > 270000) { 
      props.setProperty("AUTO_INDEX", i.toString());
      ScriptApp.newTrigger("continuarEjecucion").timeBased().after(60000).create();
      console.log("⏱️ Pausado por límite de tiempo. Reanudando en 1 min...");
      return; 
    }
    procesarProyectoGAS(token, proyectos[i].id_proy, (i === 0));
  }
  
  props.deleteProperty("AUTO_INDEX");
  props.deleteProperty("AUTO_TOKEN");
  props.deleteProperty("AUTO_PROYECTOS");
  limpiarTriggersTemporales();
  
  console.log("✅ Barrido completado. Generando Excel y Correo...");
  // 🔥 AL TERMINAR TODO EN AUTOMÁTICO, DISPARAMOS LA CREACIÓN DEL CORREO
  generarRespaldoYEnviarCorreo();
  console.log("✅ ¡Correo y respaldo enviados con éxito!");
}

function limpiarTriggersTemporales() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "continuarEjecucion") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}




// =====================================================================
// 🕒 CONFIGURACIÓN DE HORARIOS Y PRUEBAS (ZONA QUERÉTARO)
// =====================================================================

// ▶️ 1. FUNCIÓN PARA PROBAR AHORITA MISMO
function probarAhoraMismo() {
  limpiarTodosLosTriggersConfigurados();
  
  // Crea una alarma para dentro de exactamente 1 minuto
  ScriptApp.newTrigger("iniciarEjecucionAutomatica")
    .timeBased()
    .after(60 * 1000) // 60,000 milisegundos = 1 minuto
    .create();
    
  SpreadsheetApp.getUi().alert("✅ PRUEBA INICIADA: Cierra esta ventana y no toques tu teclado. En aprox 1 minuto el robot arrancará de fondo solo. Vigila tu correo en unos minutos.");
}

// ⏰ 2. INSTALAR EL DISPARADOR DIARIO (A LAS 7:00 AM EXACTAS)
function instalarDisparador7AM() {
  limpiarTodosLosTriggersConfigurados();

  // Despertamos al script a las 6 AM (Google lo ejecutará entre 6 y 7)
  ScriptApp.newTrigger("prepararEjecucionDeLas7")
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  SpreadsheetApp.getUi().alert("✅ PROGRAMADO: El robot quedó configurado para ejecutarse todos los días a las 7:00 AM exactas (Hora central/Querétaro).");
}

// Esta función es llamada internamente para saltarse la restricción de 1 hora de Google
function prepararEjecucionDeLas7() {
  var hoy = new Date();
  hoy.setHours(7, 0, 0, 0); // Le indicamos a Google que queremos las 07:00:00 exactas
  
  ScriptApp.newTrigger("iniciarEjecucionAutomatica")
    .timeBased()
    .at(hoy)
    .create();
}

function limpiarTodosLosTriggersConfigurados() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var nombre = triggers[i].getHandlerFunction();
    if (nombre === "iniciarEjecucionAutomatica" || nombre === "prepararEjecucionDeLas7") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}