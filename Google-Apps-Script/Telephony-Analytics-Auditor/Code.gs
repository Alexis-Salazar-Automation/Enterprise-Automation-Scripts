// ==========================================
// 🚀 CONFIGURACIÓN GLOBAL
// ==========================================
const ID_ARCHIVO_CARTERA = 'YOUR_SPREADSHEET_ID_HERE'; 
const API_BASE_URL = "https://api.your-company.com/index.php";
const RECIPIENT_EMAILS = ["manager@example.com"];

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('📞 Auditor de Llamadas GPH')
      .addItem('🖥️ Abrir Terminal de Auditoría', 'abrirTerminal')
      .addToUi();
}

function abrirTerminal() {
  var html = HtmlService.createHtmlOutputFromFile('Terminal')
      .setWidth(550)
      .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Terminal de Extracción de Datos');
}

// ==========================================
// 🔑 TOKEN DINÁMICO
// ==========================================
function obtenerCredencialesGPH() {
  var cache = CacheService.getUserCache();
  var credsGuardadas = cache.get('gph_creds_auditor');
  if (credsGuardadas) return JSON.parse(credsGuardadas);

  var correoActivo = Session.getActiveUser().getEmail() || "[correo_validado]";
  var nombreDerivado = correoActivo.split('@')[0].toUpperCase();
  var credencialesDinamicas = { "googleId": "[aqui se obtiene el id]", "email": correoActivo, "name": nombreDerivado };

  try {
    var resp = UrlFetchApp.fetch("https://api.your-company.com/index.php", {
      "method": "post", "contentType": "application/json",
      "payload": JSON.stringify({"data": credencialesDinamicas}), "muteHttpExceptions": true
    });
    var json = JSON.parse(resp.getContentText());
    if (json.status === 1 && json.data.token && json.data.idUser) {
      var creds = { token: json.data.token, idUser: json.data.idUser };
      cache.put('gph_creds_auditor', JSON.stringify(creds), 1500); 
      return creds;
    }
  } catch(e) {} 
  return null; 
}

// ==========================================
// 🛠️ 1. PREPARACIÓN Y EXTRACCIÓN
// ==========================================

function prepararHoja() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var cabeceras =[
    "id_agenda", "fechaA", "fecha_operacion", "observaciones", "asunto", 
    "respuesta", "id_tipoDescuento", "deudaStatus", "forma_pago", "cantidad", 
    "meses_vencidos", "id_cobranza", "id_usuario", "status", "fecha_registro", 
    "convenio", "id_asesor", "usuario_nombre", "medio", "baseneodata", 
    "cliente_nombre", "vivienda", "cliente", "StatusTerreno", "tipo_descuento", 
    "vivh", "folioPago", "monto_pago", "monto_comision", "cuotas_cubiertas", 
    "fechaRecibido", "fpagodesde", "fpagohasta", "exibicion", "idestatusComision", 
    "nestatusComision", "idSucursal", "idcomision", "Escomision", "ejecutivo", 
    "Esevidencia", "evidenciasn", "GCS", 
    "Fecha Corta", "Comentario App", "Duración llamada (s)", "Llamada Efectiva"
  ];

  // Preparar Hoja APP
  var hojaApp = libro.getSheetByName("Reporte de Llamadas") || libro.insertSheet("Reporte de Llamadas");
  hojaApp.clear();
  hojaApp.getRange(1, 1, 1, cabeceras.length).setValues([cabeceras]).setFontWeight("bold").setBackground("#2c3e50").setFontColor("white");

  // Preparar Hoja Call Center Oficial
  var hojaCC = libro.getSheetByName("Reporte Llamadas Call Center") || libro.insertSheet("Reporte Llamadas Call Center");
  hojaCC.clear();
  hojaCC.getRange(1, 1, 1, cabeceras.length).setValues([cabeceras]).setFontWeight("bold").setBackground("#4c1130").setFontColor("white");

  return "✅ Hojas base de App y Call Center listas (47 Columnas).";
}

function mapLlamadasEstandar(datosFiltrados) {
  var regexNoEfectiva = /buz[oó]n|fuera del [aá]rea|no existe|ocupad[oa]|no acepta|conmutador|no conectan|sin [eé]xito|no disponible|no enlaza|no responde|no entra la llamada|verific(?:ar|emos) la marcaci[oó]n|no lo conoce|sin contacto|no contest[oóóa]|cuelga|cuelgan|directo a buz|no se escuchaba|se corta|intentamos comunicarnos|intentamos dejar/i;

  return datosFiltrados.map(function(d) {
    var obs = d.observaciones || "";
    
    var comentarioApp = "";
    var matchCom = obs.match(/Comentario:\s*([\s\S]*?)\[Llamada/i);
    if (matchCom && matchCom[1]) comentarioApp = matchCom[1].trim();

    var duracionS = 0;
    // Identifica "Duración de 609 segundos" (CC) o "Duración: 39s" (App)
    var matchDur = obs.match(/Duraci[oó]n(?: de|:)\s*(\d+)/i);
    if (matchDur && matchDur[1]) duracionS = parseInt(matchDur[1]);

    var fechaCorta = "";
    if (d.fecha_registro) {
      var partesFecha = d.fecha_registro.split(" ")[0].split("-"); 
      fechaCorta = new Date(parseInt(partesFecha[0]), parseInt(partesFecha[1])-1, parseInt(partesFecha[2]));
    }

    var llamadaEfectiva = "Si"; 
    var textoEvaluar = comentarioApp ? comentarioApp : obs;
    var evidenciaStatus = (d.Esevidencia || "").toString().trim().toUpperCase();
    
    if (evidenciaStatus === "NO" || evidenciaStatus === "N") {
      llamadaEfectiva = "No";
    } else if (duracionS <= 20) {
      llamadaEfectiva = "No";
    } else if (regexNoEfectiva.test(textoEvaluar)) {
      llamadaEfectiva = "No";
    }

    return [
      d.id_agenda, d.fechaA, d.fecha_operacion, obs, d.asunto, 
      d.respuesta, d.id_tipoDescuento, d.deudaStatus, d.forma_pago, d.cantidad, 
      d.meses_vencidos, d.id_cobranza, d.id_usuario, d.status, d.fecha_registro, 
      d.convenio, d.id_asesor, d.usuario_nombre, d.medio, d.baseneodata, 
      d.cliente_nombre, d.vivienda, d.cliente, d.StatusTerreno, d.tipo_descuento, 
      d.vivh, d.folioPago, d.monto_pago, d.monto_comision, d.cuotas_cubiertas, 
      d.fechaRecibido, d.fpagodesde, d.fpagohasta, d.exibicion, d.idestatusComision, 
      d.nestatusComision, d.idSucursal, d.idcomision, d.Escomision, d.ejecutivo, 
      d.Esevidencia, d.evidenciasn, d.GCS, 
      fechaCorta, comentarioApp, duracionS, llamadaEfectiva
    ];
  });
}

function procesarDiaEspecial(fechaStr) {
  var credenciales = obtenerCredencialesGPH();
  if (!credenciales) throw new Error("Error de credenciales.");

  var payload = {
    "desde": fechaStr, "hasta": fechaStr, "proyectos":[], "condominios":[],
    "estatus_cob": "", "id_usuario": null, "asesor": credenciales.idUser,
    "lotes": false, "seg": false
  };

  var respuesta = UrlFetchApp.fetch("https://api.your-company.com/index.php", {
    "method": "post", "contentType": "application/json",
    "headers": { "authorization": credenciales.token },
    "payload": JSON.stringify(payload), "muteHttpExceptions": true
  });
  
  var textoRespuesta = respuesta.getContentText();
  if (textoRespuesta.trim().startsWith("<")) throw new Error("Server Crash (HTML devuelto)");

  var datosJson = JSON.parse(textoRespuesta);
  if (!Array.isArray(datosJson)) return 0; 

  // Filtramos solo las que son LLAMADAS TELEFÓNICAS primero
  var todasLlamadas = datosJson.filter(function(g) { return g.asunto === "LLAMADA TELEFÓNICA"; });
  if (todasLlamadas.length === 0) return 0;

  var textoFiltro = "Gestión registrada por medio de app Central Telefónica GPH";
  
  // Separar en 2 grupos: App y Call Center
  var llamadasApp = todasLlamadas.filter(function(g) { return g.observaciones && g.observaciones.startsWith(textoFiltro); });
  var llamadasCC = todasLlamadas.filter(function(g) { return !g.observaciones || !g.observaciones.startsWith(textoFiltro); });

  var filasApp = llamadasApp.length > 0 ? mapLlamadasEstandar(llamadasApp) : [];
  var filasCC = llamadasCC.length > 0 ? mapLlamadasEstandar(llamadasCC) : [];

  var libro = SpreadsheetApp.getActiveSpreadsheet();
  if (filasApp.length > 0) {
    var hojaApp = libro.getSheetByName("Reporte de Llamadas");
    hojaApp.getRange(hojaApp.getLastRow() + 1, 1, filasApp.length, filasApp[0].length).setValues(filasApp);
  }
  if (filasCC.length > 0) {
    var hojaCC = libro.getSheetByName("Reporte Llamadas Call Center");
    hojaCC.getRange(hojaCC.getLastRow() + 1, 1, filasCC.length, filasCC[0].length).setValues(filasCC);
  }

  // La terminal sumará este total y te lo mostrará en pantalla
  return filasApp.length + filasCC.length;
}

/*function finalizarHoja() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reporte de Llamadas");
  hoja.getRange("B:C").setNumberFormat("yyyy-mm-dd hh:mm:ss");
  hoja.getRange("AR:AR").setNumberFormat("dd/mm/yyyy"); 
  hoja.autoResizeColumns(1, 47);
  hoja.setFrozenRows(1);
  return "✅ Formatos aplicados.";
}*/

function finalizarHoja() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  
  var hojaApp = libro.getSheetByName("Reporte de Llamadas");
  if (hojaApp) {
    hojaApp.getRange("B:C").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    hojaApp.getRange("AR:AR").setNumberFormat("dd/mm/yyyy"); 
    hojaApp.autoResizeColumns(1, 47); hojaApp.setFrozenRows(1);
  }

  var hojaCC = libro.getSheetByName("Reporte Llamadas Call Center");
  if (hojaCC) {
    hojaCC.getRange("B:C").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    hojaCC.getRange("AR:AR").setNumberFormat("dd/mm/yyyy"); 
    hojaCC.autoResizeColumns(1, 47); hojaCC.setFrozenRows(1);
  }

  return "✅ Formatos aplicados en ambas hojas.";
}

function generarTablasDinamicas() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Dashboard de la APP
  crearDashboardEstandar(libro, "Reporte de Llamadas", "Dashboard Llamadas", "APP CENTRAL TELEFÓNICA", "#0052cc");
  
  // 2. Dashboard del CALL CENTER
  crearDashboardEstandar(libro, "Reporte Llamadas Call Center", "Dashboard Llamadas Call Center", "CALL CENTER OFICIAL", "#4c1130");

  // 3. Generar la super tabla Comparativa
  generarResumenComparativo(libro);

  // 4. 📧 NUEVO: ENVIAR REPORTE POR CORREO
  var msjCorreo = enviarCorreoReporte();

  return "✅ ¡Dashboards Creados! " + msjCorreo;
}

// Función auxiliar para construir Dashboards idénticos con 4 tablas
/*function crearDashboardEstandar(libro, nombreHojaDatos, nombreHojaDash, titulo, colorFondo) {
  var hojaDatos = libro.getSheetByName(nombreHojaDatos);
  if (!hojaDatos || hojaDatos.getLastRow() <= 1) return; // Si no hay datos, saltar

  var hojaTablas = libro.getSheetByName(nombreHojaDash) || libro.insertSheet(nombreHojaDash);
  hojaTablas.clear();

  var sourceData = hojaDatos.getDataRange();

  // 🧮 CALCULAR TAMAÑO DINÁMICO PARA SEPARACIÓN PERFECTA
  // Leer ejecutivos únicos para saber exactamente de qué tamaño será la tabla de arriba
  var datosEjecutivos = hojaDatos.getRange(2, 18, hojaDatos.getLastRow() - 1, 1).getValues();
  var unicos = {};
  datosEjecutivos.forEach(function(row) {
    var nombre = row[0] ? row[0].toString().trim() : "";
    if (nombre) unicos[nombre] = true;
  });
  var numEjecutivos = Object.keys(unicos).length;
  
  // Matemáticas de espaciado: Deja exactamente 3 filas vacías de separación
  var filaTituloBottom = numEjecutivos + 7;
  var filaTablaBottom = numEjecutivos + 8;

  // Tablas Superiores
  var p1 = hojaTablas.getRange('A2').createPivotTable(sourceData);
  p1.addRowGroup(18).showTotals(true); 
  p1.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA).setDisplayName("Total Llamadas");
  
  var p2 = hojaTablas.getRange('E2').createPivotTable(sourceData);
  p2.addRowGroup(18).showTotals(true); 
  p2.addColumnGroup(44).showTotals(true);
  p2.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA).setDisplayName("Total Llamadas");

  // 🟢 Tablas Inferiores (Usando la Fila Calculada Dinámicamente)
  var p3 = hojaTablas.getRange('A' + filaTablaBottom).createPivotTable(sourceData);
  p3.addRowGroup(47).showTotals(true); 
  p3.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA).setDisplayName("No. Llamadas");

  var p4 = hojaTablas.getRange('E' + filaTablaBottom).createPivotTable(sourceData);
  p4.addRowGroup(18).showTotals(true); 
  p4.addColumnGroup(47).showTotals(true); 
  p4.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA).setDisplayName("Total General");

  // Títulos Superiores
  hojaTablas.getRange("A1:B1").merge().setValue("📊 TOTAL LLAMADAS (" + titulo + ")").setFontWeight("bold").setBackground(colorFondo).setFontColor("white").setHorizontalAlignment("center");
  hojaTablas.getRange("E1:J1").merge().setValue("📅 DETALLE DÍA A DÍA").setFontWeight("bold").setBackground(colorFondo).setFontColor("white").setHorizontalAlignment("center");
  
  // 🟢 Títulos Inferiores Dinámicos
  hojaTablas.getRange("A" + filaTituloBottom + ":B" + filaTituloBottom).merge().setValue("🎯 RESUMEN EFECTIVIDAD").setFontWeight("bold").setBackground(colorFondo).setFontColor("white").setHorizontalAlignment("center");
  hojaTablas.getRange("E" + filaTituloBottom + ":H" + filaTituloBottom).merge().setValue("🕵️‍♂️ DETALLE EFECTIVAS POR EJECUTIVO").setFontWeight("bold").setBackground(colorFondo).setFontColor("white").setHorizontalAlignment("center");
  
  hojaTablas.autoResizeColumns(1, 15);
}*/

function crearDashboardEstandar(libro, nombreHojaDatos, nombreHojaDash, titulo, colorFondo) {
  var hojaDatos = libro.getSheetByName(nombreHojaDatos);
  if (!hojaDatos || hojaDatos.getLastRow() <= 1) return; // Si no hay datos, saltar

  var hojaTablas = libro.getSheetByName(nombreHojaDash) || libro.insertSheet(nombreHojaDash);
  hojaTablas.clear();

  // 🟢 1. CALCULAR MÉTRICAS PARA LAS TARJETAS VISUALES
  var datos = hojaDatos.getRange(2, 47, hojaDatos.getLastRow() - 1, 1).getValues();
  var totalLlamadas = datos.length;
  var efectivas = datos.filter(function(r) { return r[0] === "Si"; }).length;
  var noEfectivas = totalLlamadas - efectivas;
  var tasa = totalLlamadas > 0 ? (efectivas / totalLlamadas) : 0;

  // 🟢 2. DIBUJAR EL HEADER Y LAS TARJETAS (Filas 1 a 4)
  hojaTablas.getRange("A1:H1").merge().setValue("📊 DASHBOARD: " + titulo).setFontWeight("bold").setBackground(colorFondo).setFontColor("white").setHorizontalAlignment("center").setFontSize(16);

  hojaTablas.getRange("A3:B3").merge().setValue(totalLlamadas).setFontSize(26).setFontWeight("bold").setFontColor("#1565c0").setHorizontalAlignment("center");
  hojaTablas.getRange("C3:D3").merge().setValue(efectivas).setFontSize(26).setFontWeight("bold").setFontColor("#1b5e20").setHorizontalAlignment("center").setBackground("#f1f8e9");
  hojaTablas.getRange("E3:F3").merge().setValue(noEfectivas).setFontSize(26).setFontWeight("bold").setFontColor("#b71c1c").setHorizontalAlignment("center").setBackground("#ffebee");
  hojaTablas.getRange("G3:H3").merge().setValue(tasa).setNumberFormat("0.00%").setFontSize(26).setFontWeight("bold").setFontColor("#0d47a1").setHorizontalAlignment("center").setBackground("#e3f2fd");

  hojaTablas.getRange("A4:B4").merge().setValue("TOTAL LLAMADAS").setFontSize(10).setFontWeight("bold").setFontColor("#555").setHorizontalAlignment("center");
  hojaTablas.getRange("C4:D4").merge().setValue("✅ EFECTIVAS").setFontSize(10).setFontWeight("bold").setFontColor("#2e7d32").setHorizontalAlignment("center").setBackground("#e8f5e9");
  hojaTablas.getRange("E4:F4").merge().setValue("❌ NO EFECTIVAS").setFontSize(10).setFontWeight("bold").setFontColor("#c62828").setHorizontalAlignment("center").setBackground("#ffcdd2");
  hojaTablas.getRange("G4:H4").merge().setValue("🎯 TASA EFECTIVIDAD").setFontSize(10).setFontWeight("bold").setFontColor("#0d47a1").setHorizontalAlignment("center").setBackground("#bbdefb");

  var borderStyle = SpreadsheetApp.BorderStyle.SOLID;
  hojaTablas.getRange("A3:B4").setBorder(true, true, true, true, false, false, "#cfd8dc", borderStyle);
  hojaTablas.getRange("C3:D4").setBorder(true, true, true, true, false, false, "#c8e6c9", borderStyle);
  hojaTablas.getRange("E3:F4").setBorder(true, true, true, true, false, false, "#ffcdd2", borderStyle);
  hojaTablas.getRange("G3:H4").setBorder(true, true, true, true, false, false, "#bbdefb", borderStyle);

  // 🟢 3. CÁLCULO DINÁMICO DEL ESPACIADO
  var datosEjecutivos = hojaDatos.getRange(2, 18, hojaDatos.getLastRow() - 1, 1).getValues();
  var unicos = {};
  datosEjecutivos.forEach(function(row) {
    var nombre = row[0] ? row[0].toString().trim() : "";
    if (nombre) unicos[nombre] = true;
  });
  var numEjecutivos = Object.keys(unicos).length;
  
  // Fila de inicio de las tablas superiores: 8
  // Espacio que ocuparán: numEjecutivos + 2 (Por el total y la cabecera)
  // Espacio en blanco deseado: 3 filas
  var filaTituloBottom = 8 + numEjecutivos + 4;
  var filaTablaBottom = filaTituloBottom + 1;

  // 🟢 4. CREAR TABLAS DINÁMICAS
  var sourceData = hojaDatos.getDataRange();

  var p1 = hojaTablas.getRange('A8').createPivotTable(sourceData); 
  p1.addRowGroup(18).showTotals(true); 
  p1.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA).setDisplayName("Total Llamadas");
  
  var p2 = hojaTablas.getRange('E8').createPivotTable(sourceData); 
  p2.addRowGroup(18).showTotals(true); 
  p2.addColumnGroup(44).showTotals(true);
  p2.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA).setDisplayName("Total Llamadas");

  var p3 = hojaTablas.getRange('A' + filaTablaBottom).createPivotTable(sourceData);
  p3.addRowGroup(47).showTotals(true); 
  p3.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA).setDisplayName("No. Llamadas");

  var p4 = hojaTablas.getRange('E' + filaTablaBottom).createPivotTable(sourceData);
  p4.addRowGroup(18).showTotals(true); 
  p4.addColumnGroup(47).showTotals(true); 
  p4.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA).setDisplayName("Total General");

  // Títulos Superiores
  hojaTablas.getRange("A7:B7").merge().setValue("📊 TOTAL LLAMADAS").setFontWeight("bold").setBackground(colorFondo).setFontColor("white").setHorizontalAlignment("center");
  hojaTablas.getRange("E7:J7").merge().setValue("📅 DETALLE DÍA A DÍA").setFontWeight("bold").setBackground(colorFondo).setFontColor("white").setHorizontalAlignment("center");
  
  // Títulos Inferiores (Dinámicos)
  hojaTablas.getRange("A" + filaTituloBottom + ":B" + filaTituloBottom).merge().setValue("🎯 RESUMEN EFECTIVIDAD").setFontWeight("bold").setBackground(colorFondo).setFontColor("white").setHorizontalAlignment("center");
  hojaTablas.getRange("E" + filaTituloBottom + ":H" + filaTituloBottom).merge().setValue("🕵️‍♂️ DETALLE EFECTIVAS POR EJECUTIVO").setFontWeight("bold").setBackground(colorFondo).setFontColor("white").setHorizontalAlignment("center");
  
  hojaTablas.autoResizeColumns(1, 15);
}


function generarResumenComparativo(libro) {
  var hojaResumen = libro.getSheetByName("Resumen Comparativo") || libro.insertSheet("Resumen Comparativo", 0);
  hojaResumen.clear();
  
  // Limpiar filtros (slicers) anteriores si existen
  var slicers = hojaResumen.getSlicers();
  slicers.forEach(function(s) { s.remove(); });

  var hojaApp = libro.getSheetByName("Reporte de Llamadas");
  var hojaCC = libro.getSheetByName("Reporte Llamadas Call Center");

  var resumen = {}; 

  // 🟢 Recolectar datos APP (Columna 18 = Ejecutivo, Columna 47 = Efectiva)
  if (hojaApp && hojaApp.getLastRow() > 1) {
    var datosApp = hojaApp.getRange(2, 1, hojaApp.getLastRow() - 1, 47).getValues(); 
    for (var i = 0; i < datosApp.length; i++) {
      var exec = datosApp[i][17] ? datosApp[i][17].toString().trim() : "SIN ASIGNAR";
      var esEfectiva = datosApp[i][46] === "Si" ? 1 : 0;
      
      if (!resumen[exec]) resumen[exec] = { app: 0, appEfec: 0, cc: 0, ccEfec: 0 };
      resumen[exec].app++;
      resumen[exec].appEfec += esEfectiva;
    }
  }

  // 🟢 Recolectar datos Call Center (Columna 18 = Ejecutivo, Columna 47 = Efectiva)
  if (hojaCC && hojaCC.getLastRow() > 1) {
    var datosCC = hojaCC.getRange(2, 1, hojaCC.getLastRow() - 1, 47).getValues();
    for (var k = 0; k < datosCC.length; k++) {
      var execCC = datosCC[k][17] ? datosCC[k][17].toString().trim() : "SIN ASIGNAR";
      var esEfectivaCC = datosCC[k][46] === "Si" ? 1 : 0;
      
      if (!resumen[execCC]) resumen[execCC] = { app: 0, appEfec: 0, cc: 0, ccEfec: 0 };
      resumen[execCC].cc++;
      resumen[execCC].ccEfec += esEfectivaCC;
    }
  }

  // 🟢 NUEVO: Tabla de 8 Columnas (Uso y Efectividad por Platarforma)
  var tabla = [["👨‍💼 Ejecutivo / Asesor", "📱 Llamadas Central GPH", "🚀 % Uso Central", "✅ % Efec. Central", "📞 Llamadas Call Center", "🚀 % Uso CC", "✅ % Efec. CC", "📈 Total Global"]];
  var totalApp = 0, totalAppEfec = 0, totalCC = 0, totalCCEfec = 0, totalGlobal = 0;

  var ejecutivos = Object.keys(resumen).sort();
  for (var j = 0; j < ejecutivos.length; j++) {
    var e = ejecutivos[j];
    var t = resumen[e].app + resumen[e].cc;
    
    var pctUsoApp = t > 0 ? (resumen[e].app / t) : 0; 
    var pctEfecApp = resumen[e].app > 0 ? (resumen[e].appEfec / resumen[e].app) : 0;
    
    var pctUsoCC = t > 0 ? (resumen[e].cc / t) : 0; 
    var pctEfecCC = resumen[e].cc > 0 ? (resumen[e].ccEfec / resumen[e].cc) : 0;
    
    tabla.push([e, resumen[e].app, pctUsoApp, pctEfecApp, resumen[e].cc, pctUsoCC, pctEfecCC, t]);
    
    totalApp += resumen[e].app;
    totalAppEfec += resumen[e].appEfec;
    totalCC += resumen[e].cc;
    totalCCEfec += resumen[e].ccEfec;
    totalGlobal += t;
  }
  
  var pGlobalUsoApp = totalGlobal > 0 ? (totalApp / totalGlobal) : 0;
  var pGlobalEfecApp = totalApp > 0 ? (totalAppEfec / totalApp) : 0;
  var pGlobalUsoCC = totalGlobal > 0 ? (totalCC / totalGlobal) : 0;
  var pGlobalEfecCC = totalCC > 0 ? (totalCCEfec / totalCC) : 0;
  
  tabla.push(["🏆 TOTAL GENERAL", totalApp, pGlobalUsoApp, pGlobalEfecApp, totalCC, pGlobalUsoCC, pGlobalEfecCC, totalGlobal]);

  // Imprimir Tabla en Hoja (Ahora usa 8 columnas)
  hojaResumen.getRange(1, 1, 1, 8).merge().setValue("⚖️ REPORTE DE MIGRACIÓN: APP VS CALL CENTER OFICIAL").setFontWeight("bold").setBackground("#166534").setFontColor("white").setHorizontalAlignment("center").setFontSize(14);
  hojaResumen.getRange(2, 1, tabla.length, 8).setValues(tabla).setHorizontalAlignment("center");
  hojaResumen.getRange(2, 1, 1, 8).setFontWeight("bold").setBackground("#e0e0e0"); // Cabecera
  hojaResumen.getRange(tabla.length + 1, 1, 1, 8).setFontWeight("bold").setBackground("#dcfce7"); // Fila de totales
  
  // 🟢 Formatos de porcentaje (Columnas C, D, F y G)
  hojaResumen.getRange(3, 3, tabla.length - 1, 2).setNumberFormat("0.0%"); 
  hojaResumen.getRange(3, 6, tabla.length - 1, 2).setNumberFormat("0.0%"); 
  hojaResumen.autoResizeColumns(1, 8);

  // ========================================================
  // 🌟 BASE CONSOLIDADA Y TABLA INTERACTIVA
  // ========================================================
  var hojaBD = libro.getSheetByName("BD_Consolidada") || libro.insertSheet("BD_Consolidada");
  hojaBD.clear();

  var dataConsolidada = [];
  var headers = [];

  if (hojaApp && hojaApp.getLastRow() > 1) {
    headers = hojaApp.getRange(1, 1, 1, 47).getValues()[0];
    headers.push("Origen"); 
    var datosA = hojaApp.getRange(2, 1, hojaApp.getLastRow() - 1, 47).getValues();
    datosA.forEach(function(r) { r.push("📱 APP"); dataConsolidada.push(r); });
  }

  if (hojaCC && hojaCC.getLastRow() > 1) {
    if(headers.length === 0) {
      headers = hojaCC.getRange(1, 1, 1, 47).getValues()[0];
      headers.push("Origen");
    }
    var datosC = hojaCC.getRange(2, 1, hojaCC.getLastRow() - 1, 47).getValues();
    datosC.forEach(function(r) { r.push("📞 CALL CENTER"); dataConsolidada.push(r); });
  }

  if (dataConsolidada.length > 0) {
    hojaBD.getRange(1, 1, 1, headers.length).setValues([headers]);
    hojaBD.getRange(2, 1, dataConsolidada.length, headers.length).setValues(dataConsolidada);

    var filaInicioInteractiva = tabla.length + 5;
    hojaResumen.getRange(filaInicioInteractiva, 1, 1, 5).merge().setValue("📅 TENDENCIA DIARIA DE ADOPCIÓN (INTERACTIVO)").setFontWeight("bold").setBackground("#0f766e").setFontColor("white").setHorizontalAlignment("center").setFontSize(12);

    var sourceRange = hojaBD.getDataRange();
    var pivotTendencia = hojaResumen.getRange(filaInicioInteractiva + 2, 1).createPivotTable(sourceRange);
    
    pivotTendencia.addRowGroup(18).showTotals(false); // 18 = usuario_nombre
    pivotTendencia.addRowGroup(44).showTotals(true);  // 44 = Fecha Corta
    pivotTendencia.addColumnGroup(48).showTotals(true); // 48 = Origen
    pivotTendencia.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA).setDisplayName("Gestiones");

    var slicer = hojaResumen.insertSlicer(sourceRange, filaInicioInteractiva + 2, 9); // Flotando en columna I (9)
    var filtroVacio = SpreadsheetApp.newFilterCriteria().build();
    slicer.setColumnFilterCriteria(18, filtroVacio); 
    slicer.setTitle("🔎 FILTRAR POR EJECUTIVO");
  }
}

/*function enviarCorreoReporte() {
  try {
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.flush(); // Guardar cambios

    var hojaApp = libro.getSheetByName("Reporte de Llamadas");
    var hojaCC = libro.getSheetByName("Reporte Llamadas Call Center");
    var hojaResumen = libro.getSheetByName("Resumen Comparativo");
    
    var totalApp = (hojaApp && hojaApp.getLastRow() > 1) ? hojaApp.getLastRow() - 1 : 0;
    var totalCC = (hojaCC && hojaCC.getLastRow() > 1) ? hojaCC.getLastRow() - 1 : 0;
    var totalGlobal = totalApp + totalCC;
    
    if (totalGlobal === 0) return "No hay datos para enviar por correo.";

    // 🟢 Calcular Efectividad APP
    var efectivasApp = 0;
    if (totalApp > 0) {
      var datosAppEfect = hojaApp.getRange(2, 47, totalApp, 1).getValues(); 
      efectivasApp = datosAppEfect.filter(function(r) { return r[0] === "Si"; }).length;
    }

    // 🟢 Calcular Efectividad CALL CENTER TRADICIONAL
    var efectivasCC = 0;
    if (totalCC > 0) {
      var datosCCEfect = hojaCC.getRange(2, 47, totalCC, 1).getValues(); 
      efectivasCC = datosCCEfect.filter(function(r) { return r[0] === "Si"; }).length;
    }

    // Calcular Tasas % Visuales para el correo
    var txtTasaEfecApp = totalApp > 0 ? ((efectivasApp / totalApp) * 100).toFixed(1) + "%" : "0.0%";
    var txtTasaEfecCC = totalCC > 0 ? ((efectivasCC / totalCC) * 100).toFixed(1) + "%" : "0.0%";

    // 🟢 CONSTRUCCIÓN DE LA TABLA HTML PARA EL CORREO (Ahora 8 columnas completas)
    var htmlTabla = "";
    if (hojaResumen) {
      var dataResumen = hojaResumen.getDataRange().getDisplayValues(); 
      var tablaExtraida = [];
      
      for(var i = 1; i < dataResumen.length; i++) {
        if (dataResumen[i][0] === "") break; 
        tablaExtraida.push(dataResumen[i]);
        if (dataResumen[i][0].includes("TOTAL GENERAL")) break; 
      }

      htmlTabla += "<table style='border-collapse: collapse; width: 100%; font-size: 11px; margin-top: 15px;'>"; // Letra más pequeña para que quepan 8 cols
      for(var r = 0; r < tablaExtraida.length; r++) {
        var isHeader = (r === 0);
        var isTotal = (r === tablaExtraida.length - 1);
        var bg = isHeader ? "#e0e0e0" : (isTotal ? "#dcfce7" : "#ffffff");
        var weight = (isHeader || isTotal) ? "bold" : "normal";
        
        htmlTabla += `<tr style='background-color: ${bg}; font-weight: ${weight}; text-align: center;'>`;
        for(var c = 0; c < 8; c++) { // 🟢 Ahora el ciclo llega hasta 8 columnas
          var align = (c === 0 && !isHeader && !isTotal) ? "left" : "center"; 
          var tag = isHeader ? "th" : "td";
          htmlTabla += `<${tag} style='border: 1px solid #ccc; padding: 6px; text-align: ${align};'>${tablaExtraida[r][c]}</${tag}>`;
        }
        htmlTabla += "</tr>";
      }
      htmlTabla += "</table>";
    }

    // Exportar Excel
    var urlDescarga = "https://docs.google.com/spreadsheets/d/" + libro.getId() + "/export?format=xlsx";
    var tokenAuth = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(urlDescarga, { headers: { 'Authorization': 'Bearer ' + tokenAuth }, muteHttpExceptions: true });
    
    var fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy");
    var blobExcel = response.getBlob().setName("Reporte_Auditoria_Llamadas_" + fechaHoy + ".xlsx");

    // 🟢 Construir el Cuerpo del Correo (Con los súper indicadores de Efectividad)
    var htmlBody = `
      <div style="font-family: Arial, sans-serif; color: #333; max-width: 900px; margin: auto;">
        <h2 style="color: #0f766e;">📊 Reporte Automático: Auditoría de Gestiones GPH</h2>
        <p>Hola equipo,</p>
        <p>Se ha completado exitosamente la extracción y análisis de las llamadas telefónicas. Adjunto a este correo encontrarán el archivo Excel (<strong>.xlsx</strong>) con la base de datos completa y los Dashboards interactivos.</p>
        
        <h3 style="border-bottom: 2px solid #0f766e; padding-bottom: 5px;">📈 Resumen de Adopción Global:</h3>
        <ul>
          <li>📱 <b>Llamadas en Central GPH:</b> ${totalApp}</li>
          <li>📞 <b>Llamadas en Call Center Tradicional:</b> ${totalCC}</li>
          <li>🏆 <b style="color:#0f766e;">Total de Gestiones:</b> ${totalGlobal}</li>
        </ul>

        <h3 style="border-bottom: 2px solid #0052cc; padding-bottom: 5px;">🎯 Efectividad de Central GPH:</h3>
        <ul>
          <li>Llamadas Efectivas (Con contacto real): <b>${efectivasApp}</b></li>
          <li>Llamadas No Efectivas (Buzón, cuelga, sin éxito): <b>${totalApp - efectivasApp}</b></li>
          <li style="margin-top: 5px; list-style-type: none;">🔥 Tasa de Efectividad: <b style="color:#0052cc; font-size: 16px;">${txtTasaEfecApp}</b></li>
        </ul>

        <h3 style="border-bottom: 2px solid #9f1239; padding-bottom: 5px;">🎯 Efectividad de Call Center Tradicional:</h3>
        <ul>
          <li>Llamadas Efectivas (Con contacto real): <b>${efectivasCC}</b></li>
          <li>Llamadas No Efectivas (Buzón, cuelga, sin éxito): <b>${totalCC - efectivasCC}</b></li>
          <li style="margin-top: 5px; list-style-type: none;">🔥 Tasa de Efectividad: <b style="color:#9f1239; font-size: 16px;">${txtTasaEfecCC}</b></li>
        </ul>

        <h3 style="border-bottom: 2px solid #166534; padding-bottom: 5px;">⚖️ Detalle de Adopción y Desempeño por Ejecutivo:</h3>
        ${htmlTabla}

        <br><br>
        <p>Por favor, abran el archivo adjunto para explorar los detalles en la pestaña <i>Resumen Comparativo</i> usando la gráfica de tendencia diaria.</p>
        <br>
        <p style="font-size: 12px; color: #777; border-top: 1px solid #ddd; padding-top: 10px;">
          Este es un mensaje generado automáticamente por el Motor de Extracción GPH 🤖. No responder a este correo.
        </p>
      </div>
    `;

    // Ejecutar el envío
    if (typeof CORREOS_DESTINO !== 'undefined' && CORREOS_DESTINO.length > 0) {
      MailApp.sendEmail({
        to: CORREOS_DESTINO.join(","),
        cc: CORREOS_COPIA.join(","),
        subject: "📊 Reporte de Auditoría de Llamadas GPH - " + fechaHoy,
        htmlBody: htmlBody,
        attachments: [blobExcel]
      });
      return "📧 Correo enviado a " + CORREOS_DESTINO.join(", ");
    } else {
      return "⚠️ No se envió correo porque no hay destinatarios configurados.";
    }

  } catch (error) {
    return "❌ Error al enviar el correo: " + error.message;
  }
}*/

function enviarCorreoReporte() {
  try {
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.flush(); // Guardar cambios antes de exportar

    var hojaApp = libro.getSheetByName("Reporte de Llamadas");
    var hojaCC = libro.getSheetByName("Reporte Llamadas Call Center");
    var hojaResumen = libro.getSheetByName("Resumen Comparativo");
    
    var totalApp = (hojaApp && hojaApp.getLastRow() > 1) ? hojaApp.getLastRow() - 1 : 0;
    var totalCC = (hojaCC && hojaCC.getLastRow() > 1) ? hojaCC.getLastRow() - 1 : 0;
    var totalGlobal = totalApp + totalCC;
    
    if (totalGlobal === 0) return "No hay datos para enviar por correo.";

    // Calcular Efectividad APP
    var efectivasApp = 0;
    if (totalApp > 0) {
      var datosAppEfect = hojaApp.getRange(2, 47, totalApp, 1).getValues(); 
      efectivasApp = datosAppEfect.filter(function(r) { return r[0] === "Si"; }).length;
    }

    // Calcular Efectividad CALL CENTER
    var efectivasCC = 0;
    if (totalCC > 0) {
      var datosCCEfect = hojaCC.getRange(2, 47, totalCC, 1).getValues(); 
      efectivasCC = datosCCEfect.filter(function(r) { return r[0] === "Si"; }).length;
    }

    // Métricas Globales
    var efectivasGlobal = efectivasApp + efectivasCC;
    var tasaGlobal = totalGlobal > 0 ? ((efectivasGlobal / totalGlobal) * 100).toFixed(1) + "%" : "0.0%";
    var txtTasaEfecApp = totalApp > 0 ? ((efectivasApp / totalApp) * 100).toFixed(1) + "%" : "0.0%";
    var txtTasaEfecCC = totalCC > 0 ? ((efectivasCC / totalCC) * 100).toFixed(1) + "%" : "0.0%";

    // 🟢 CONSTRUCCIÓN DE LA TABLA HTML Y BÚSQUEDA DEL EJECUTIVO DESTACADO
    var htmlTabla = "";
    var topEjecutivo = "N/A";
    var topLlamadas = 0;

    if (hojaResumen) {
      var dataResumen = hojaResumen.getDataRange().getDisplayValues(); 
      var tablaExtraida = [];
      
      for(var i = 1; i < dataResumen.length; i++) {
        if (dataResumen[i][0] === "") break; 
        tablaExtraida.push(dataResumen[i]);
        
        // Buscar al top performer (saltando cabecera y totales)
        if (i > 1 && !dataResumen[i][0].includes("TOTAL GENERAL")) {
          var llamadasEjecutivo = parseInt(dataResumen[i][7]) || 0; // Columna 7 es Total Global
          if (llamadasEjecutivo > topLlamadas) {
            topLlamadas = llamadasEjecutivo;
            topEjecutivo = dataResumen[i][0];
          }
        }
        if (dataResumen[i][0].includes("TOTAL GENERAL")) break; 
      }

      htmlTabla += "<table style='border-collapse: collapse; width: 100%; font-size: 12px; font-family: Arial, sans-serif;'>"; 
      for(var r = 0; r < tablaExtraida.length; r++) {
        var isHeader = (r === 0);
        var isTotal = (r === tablaExtraida.length - 1);
        
        // Diseño de cebras para la tabla
        var bg = isHeader ? "#1e293b" : (isTotal ? "#e2e8f0" : (r % 2 === 0 ? "#f8fafc" : "#ffffff"));
        var colorTexto = isHeader ? "#ffffff" : (isTotal ? "#0f172a" : "#334155");
        var weight = (isHeader || isTotal) ? "bold" : "normal";
        var border = isHeader ? "none" : "1px solid #e2e8f0";
        
        htmlTabla += `<tr style='background-color: ${bg}; color: ${colorTexto}; font-weight: ${weight}; text-align: center;'>`;
        for(var c = 0; c < 8; c++) { 
          var align = (c === 0) ? "left" : "center"; 
          var tag = isHeader ? "th" : "td";
          htmlTabla += `<${tag} style='border: ${border}; padding: 10px 8px; text-align: ${align};'>${tablaExtraida[r][c]}</${tag}>`;
        }
        htmlTabla += "</tr>";
      }
      htmlTabla += "</table>";
    }

    // Exportar Excel
    var urlDescarga = "https://docs.google.com/spreadsheets/d/" + libro.getId() + "/export?format=xlsx";
    var tokenAuth = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(urlDescarga, { headers: { 'Authorization': 'Bearer ' + tokenAuth }, muteHttpExceptions: true });
    
    var fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy");
    var blobExcel = response.getBlob().setName("Reporte_Auditoria_Llamadas_" + fechaHoy + ".xlsx");

    // 🟢 NUEVO DISEÑO DEL CORREO (ESTILO CORPORATIVO MODERNO)
    var htmlBody = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 900px; margin: auto; background-color: #f8fafc; padding: 20px; border-radius: 10px;">
        
        <!-- Header -->
        <div style="background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); border-radius: 8px 8px 0 0; padding: 25px; text-align: center;">
          <h1 style="color: #ffffff; margin: 0; font-size: 24px;">📊 Reporte Ejecutivo de Auditoría</h1>
          <p style="color: #94a3b8; margin: 5px 0 0 0; font-size: 14px;">Central Telefónica GPH | Fecha: ${fechaHoy}</p>
        </div>

        <!-- Content -->
        <div style="background-color: #ffffff; padding: 30px; border-radius: 0 0 8px 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
          <p style="color: #334155; font-size: 16px; line-height: 1.5;">Hola equipo,</p>
          <p style="color: #334155; font-size: 15px; line-height: 1.5; margin-bottom: 25px;">Se ha completado el análisis de efectividad. Hemos procesado un total de <strong>${totalGlobal} gestiones</strong>. Adjunto encontrarán la base de datos completa.</p>

          <!-- 4 KPI Cards -->
          <table width="100%" cellpadding="10" cellspacing="0" style="margin-bottom: 30px;">
            <tr>
              <td width="25%" align="center">
                <div style="background-color: #f1f5f9; border-top: 4px solid #3b82f6; padding: 15px; border-radius: 6px;">
                  <p style="margin:0; font-size: 11px; color: #64748b; font-weight: bold; text-transform: uppercase;">Total Gestiones</p>
                  <h2 style="margin: 10px 0 5px 0; font-size: 26px; color: #1e293b;">${totalGlobal}</h2>
                  <p style="margin:0; font-size: 11px; color: #94a3b8;">Llamadas procesadas</p>
                </div>
              </td>
              <td width="25%" align="center">
                <div style="background-color: #f1f5f9; border-top: 4px solid #8b5cf6; padding: 15px; border-radius: 6px;">
                  <p style="margin:0; font-size: 11px; color: #64748b; font-weight: bold; text-transform: uppercase;">Efec. Global</p>
                  <h2 style="margin: 10px 0 5px 0; font-size: 26px; color: #1e293b;">${tasaGlobal}</h2>
                  <p style="margin:0; font-size: 11px; color: #94a3b8;">${efectivasGlobal} con contacto</p>
                </div>
              </td>
              <td width="25%" align="center">
                <div style="background-color: #f0fdf4; border-top: 4px solid #10b981; padding: 15px; border-radius: 6px;">
                  <p style="margin:0; font-size: 11px; color: #64748b; font-weight: bold; text-transform: uppercase;">Efec. Central GPH</p>
                  <h2 style="margin: 10px 0 5px 0; font-size: 26px; color: #166534;">${txtTasaEfecApp}</h2>
                  <p style="margin:0; font-size: 11px; color: #94a3b8;">${efectivasApp} de ${totalApp}</p>
                </div>
              </td>
              <td width="25%" align="center">
                <div style="background-color: #fef2f2; border-top: 4px solid #ef4444; padding: 15px; border-radius: 6px;">
                  <p style="margin:0; font-size: 11px; color: #64748b; font-weight: bold; text-transform: uppercase;">Efec. Call Center</p>
                  <h2 style="margin: 10px 0 5px 0; font-size: 26px; color: #991b1b;">${txtTasaEfecCC}</h2>
                  <p style="margin:0; font-size: 11px; color: #94a3b8;">${efectivasCC} de ${totalCC}</p>
                </div>
              </td>
            </tr>
          </table>

          <!-- Highlight Section -->
          <div style="background-color: #fffbeb; border-left: 4px solid #f59e0b; padding: 15px 20px; margin-bottom: 30px; border-radius: 4px;">
            <h4 style="margin: 0 0 5px 0; color: #b45309; font-size: 16px;">🌟 Destacado del periodo</h4>
            <p style="margin: 0; color: #78350f; font-size: 14px;">El ejecutivo con mayor número de gestiones fue <strong>${topEjecutivo}</strong> con <strong>${topLlamadas} llamadas</strong> en total.</p>
          </div>

          <!-- Table -->
          <h3 style="color: #0f172a; margin-bottom: 15px; font-size: 18px; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px;">⚖️ Detalle de Adopción y Desempeño por Ejecutivo</h3>
          <div style="overflow-x: auto;">
            ${htmlTabla}
          </div>

          <p style="font-size: 12px; color: #94a3b8; border-top: 1px solid #e2e8f0; padding-top: 15px; margin-top: 30px; text-align: center;">
            🤖 Motor de Extracción y Auditoría GPH<br>Este es un reporte generado automáticamente.
          </p>
        </div>
      </div>
    `;

    // Ejecutar el envío
    if (typeof CORREOS_DESTINO !== 'undefined' && CORREOS_DESTINO.length > 0) {
      MailApp.sendEmail({
        to: CORREOS_DESTINO.join(","),
        cc: CORREOS_COPIA.join(","),
        subject: "📊 Reporte Ejecutivo de Auditoría GPH - " + fechaHoy,
        htmlBody: htmlBody,
        attachments: [blobExcel]
      });
      return "📧 Correo enviado a " + CORREOS_DESTINO.join(", ");
    } else {
      return "⚠️ No se envió correo porque no hay destinatarios configurados.";
    }

  } catch (error) {
    return "❌ Error al enviar el correo: " + error.message;
  }
}