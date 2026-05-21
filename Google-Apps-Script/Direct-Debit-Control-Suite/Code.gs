// ID de la hoja de cálculo "Cartera"
const CARTERA_ID = 'YOUR_MASTER_DATABASE_ID';

// CORREOS PARA EL REPORTE
const CORREOS_REPORTE = [
  'supervisor1@example.com',
  'admin_operaciones@example.com'
];

// Nombres de las hojas
const HOJA_DOMICILIACIONES = 'Domiciliaciones';
const HOJA_MAESTRA_LOCAL = '_ListasMaestras';
const HOJA_CLIENTES = '_Listas_Clientes';
const HOJA_ASESORES = '_Listas_Asesores';
const HOJA_PLANTILLA_ACUSE = 'PlantillaAcuse';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const puedeEditar = SpreadsheetApp.getActiveSpreadsheet().getRange('A1').canEdit();
  
  const menu = ui.createMenu('⚙️ Domiciliaciones');
  if (puedeEditar) {
    menu.addItem('➕ Registrar Nueva Domiciliación', 'mostrarFormularioDomiciliacion');
    menu.addSeparator();
    menu.addItem('📄 Generar Acuse de Domiciliaciones', 'mostrarSelectorAnalista');
    menu.addSeparator();
    menu.addItem('⭐ Proceso de Actualización Integral', 'procesoActualizacionIntegralUI'); 
    menu.addItem('🔄 Sincronizar con Cartera', 'sincronizarDatosCompletosUI');
  }
  menu.addToUi();
}

// Wrappers para el menú
function procesoActualizacionIntegralUI() { procesoActualizacionIntegral(false); }
function sincronizarDatosCompletosUI() { sincronizarDatosCompletos(false); }

// --- FUNCIONES DEL FORMULARIO ---

function mostrarFormularioDomiciliacion() {
  const html = HtmlService.createHtmlOutputFromFile('FormularioDomiciliacion').setWidth(800).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Registrar Nueva Domiciliación');
}

function obtenerDatosParaFormulario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const asesores = ss.getSheetByName(HOJA_ASESORES) ? ss.getSheetByName(HOJA_ASESORES).getDataRange().getValues().flat().filter(Boolean) : [];
  const puedeVerTarjeta = ss.getRange('A1').canEdit(); 
  return { asesores, puedeVerTarjeta };
}

function guardarDomiciliacion(datosFormulario) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaDestino = ss.getSheetByName(HOJA_DOMICILIACIONES);
    if (!hojaDestino) throw new Error(`La hoja "${HOJA_DOMICILIACIONES}" no fue encontrada.`);
    
    const { commonDetails, propertiesData } = datosFormulario;
    const emailUsuario = Session.getActiveUser().getEmail();

    const fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const registradoPor = `${emailUsuario} ${fechaActual}`;

    const colAValues = hojaDestino.getRange("A1:A").getValues().flat().reverse();
    const lastID = colAValues.find(id => String(id).startsWith('DOM-'));
    const nextNum = lastID ? (parseInt(lastID.split('-')[2], 10) || 0) + 1 : 1;
    const year = new Date().getFullYear();
    const ID_UNICO = `DOM-${year}-${String(nextNum).padStart(4, '0')}`;

    const comentarioFinal = commonDetails.comentariosGPH === 'otro' ? commonDetails.comentarioPersonalizado : commonDetails.comentariosGPH;
    const esCancelacion = commonDetails.comentariosGPH === 'Cancelación';

    let montoFinal, diaFinal, mesFinal, anioFinal;

    if (esCancelacion) {
      montoFinal = "-";
      diaFinal = "-";
      mesFinal = "-";
      anioFinal = "-";
    } else {
      montoFinal = (commonDetails.monto && !isNaN(parseFloat(commonDetails.monto))) ? parseFloat(commonDetails.monto) : "";
      diaFinal = commonDetails.dia;
      mesFinal = commonDetails.mes;
      anioFinal = commonDetails.anio;
    }

    const filasAGuardar = propertiesData.map(prop => [
        ID_UNICO, 
        commonDetails.fechaRecepcion, 
        commonDetails.fechaEnvioAdmin,
        prop.propietario, 
        prop.proyecto, 
        prop.condominio, 
        prop.lote, 
        prop.loteCartera,
        prop.referencia, 
        prop.asesor, 
        montoFinal, 
        commonDetails.tarjeta,
        diaFinal,   
        mesFinal,   
        anioFinal,  
        commonDetails.tipoDocumento.join(', '), 
        comentarioFinal, 
        "", 
        registradoPor, 
        "", "", "", "" 
    ]);

    if (filasAGuardar.length > 0) {
      hojaDestino.getRange(_encontrarProximaFilaVacia(hojaDestino), 1, filasAGuardar.length, filasAGuardar[0].length).setValues(filasAGuardar);
    }
    return `¡Éxito! Se registró la domiciliación con Folio: ${ID_UNICO}.`;
  } catch (e) {
    Logger.log(e);
    return `Error: ${e.message}`;
  } finally {
    lock.releaseLock();
  }
}

// --- FUNCIONES DE ACUSE ---

function mostrarSelectorAnalista() {
    const html = HtmlService.createHtmlOutputFromFile('SelectorAnalista').setWidth(400).setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'Seleccionar Analista para Acuse');
}

function obtenerAnalistasConPendientes() {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_DOMICILIACIONES);
    if (!hoja) return [];
    const data = hoja.getDataRange().getValues();
    const headers = data.shift();
    const idxRecordatorio = headers.indexOf("RECORDATORIO");
    const idxRegistrado = headers.indexOf("REGISTRADO POR");

    if (idxRegistrado === -1 || idxRecordatorio === -1) throw new Error('No se encontraron las columnas "REGISTRADO POR" o "RECORDATORIO".');

    const analistasPendientes = new Set();
    data.forEach(row => {
        if (row[idxRegistrado] && row[idxRecordatorio] === "") {
            analistasPendientes.add(row[idxRegistrado]);
        }
    });
    return Array.from(analistasPendientes).sort();
}

function generarAcusePorLote(selectedAnalistas) {
  const ui = SpreadsheetApp.getUi();
  if (!selectedAnalistas || selectedAnalistas.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast("No se seleccionó ningún analista.", "Fallo", 5);
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDomi = ss.getSheetByName(HOJA_DOMICILIACIONES);
  const datos = hojaDomi.getDataRange().getValues();
  const headers = datos.shift();
  const colIndex = headers.reduce((acc, h, i) => { acc[h.toUpperCase()] = i; return acc; }, {});
  
  const datosParaPDF = [];
  const filasAActualizar = new Set();

  datos.forEach((row, index) => {
    if (selectedAnalistas.includes(row[colIndex["REGISTRADO POR"]]) && row[colIndex["RECORDATORIO"]] === "") {
        datosParaPDF.push({
            folio: row[colIndex.FOLIO],
            desarrollo: row[colIndex.DESARROLLO],
            condominio: row[colIndex.CONDOMINIO],
            lote: row[colIndex.LOTE],
            fechaEnvio: row[colIndex["FECHA DE ENVIO ADMIN"]],
            fechaRecepcion: row[colIndex["FECHA DE RECEPCIÓN"]]
        });
        filasAActualizar.add(index + 2);
    }
  });

  if (datosParaPDF.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast("No se encontraron domiciliaciones pendientes para los analistas seleccionados.", "Info", 5);
    return;
  }
  
  try {
    const nombreQuienGenera = getNombreUsuario();
    const datosAcuse = { 
        nombreAnalista: nombreQuienGenera, 
        domiciliaciones: datosParaPDF 
    };
    _generarYEnviarPDFAcuse(datosAcuse, Session.getActiveUser().getEmail());
    
    const fechaMarca = "Generado el " + new Date().toLocaleString();
    const colRecordatorio = colIndex["RECORDATORIO"] + 1;
    filasAActualizar.forEach(numFila => hojaDomi.getRange(numFila, colRecordatorio).setValue(fechaMarca));
    
    SpreadsheetApp.getActiveSpreadsheet().toast(`¡Éxito! Se generó un acuse con ${datosParaPDF.length} registros.`, "Proceso Completado", 10);
  } catch (e) {
    ui.alert("Error", `Ocurrió un error: ${e.message}`, ui.ButtonSet.OK);
  }
}

function _generarYEnviarPDFAcuse(datos, destinatarioEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const plantilla = ss.getSheetByName(HOJA_PLANTILLA_ACUSE);
  if (!plantilla) throw new Error('No se encontró la hoja "PlantillaAcuse"');
  
  plantilla.showSheet();
  const tempSheet = plantilla.copyTo(ss).setName('AcuseTemp_' + Date.now());
  
  try {
    const hoy = new Date();
    
    tempSheet.createTextFinder('{{fecha_acuse}}').replaceAllWith(formatearFechaEnEspanol(hoy));
    tempSheet.createTextFinder('{{semana_del_anio}}').replaceAllWith("Semana " + Utilities.formatDate(hoy, Session.getScriptTimeZone(), "w"));
    tempSheet.createTextFinder('{{nombre_analista}}').replaceAllWith(getNombreUsuario());
    
    const tablaInicioFila = 14;
    
    const datosTabla = datos.domiciliaciones.map((d, i) => [
        i + 1,
        d.folio,
        d.desarrollo,
        d.condominio,
        d.lote,
        d.fechaEnvio instanceof Date ? Utilities.formatDate(d.fechaEnvio, Session.getScriptTimeZone(), 'dd/MM/yyyy') : d.fechaEnvio,
        d.fechaRecepcion instanceof Date ? Utilities.formatDate(d.fechaRecepcion, Session.getScriptTimeZone(), 'dd/MM/yyyy') : d.fechaRecepcion,
        "-"
    ]);
    
    if (datosTabla.length > 0) {
        const rangoDeTrabajo = tempSheet.getRange("B14:I100");
        rangoDeTrabajo.breakApart().clearContent().clearFormat();
        const rangoTabla = tempSheet.getRange(tablaInicioFila, 2, datosTabla.length, datosTabla[0].length);
        rangoTabla.setValues(datosTabla)
                .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID)
                .setHorizontalAlignment("center")
                .setVerticalAlignment("middle")
                .setFontSize(10)
                .setWrap(true);
        for (let i = 0; i < datosTabla.length; i++) {
            if ((i + 1) % 2 === 0) {
                tempSheet.getRange(tablaInicioFila + i, 2, 1, datosTabla[0].length).setBackground("#f3f3f3");
            }
        }
        const filaInicioFirmas = tablaInicioFila + datosTabla.length + 3;
        const firmantes = ["NOMBRE DE QUIEN RECIBE:", "PUESTO:", "ÁREA:", "FECHA Y FIRMA:"];
        firmantes.forEach((label, i) => {
            const filaActual = filaInicioFirmas + (i * 2);
            tempSheet.getRange(filaActual, 2, 1, 3).merge().setValue(label).setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");
            tempSheet.getRange(filaActual, 5, 1, 4).merge().setBorder(null, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
        });
        [20, 120, 120, 150, 40, 100, 100, 50].forEach((width, i) => {
            tempSheet.setColumnWidth(i + 2, width);
        });
    }
    
    SpreadsheetApp.flush();
    const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?exportFormat=pdf&gid=${tempSheet.getSheetId()}&size=letter&portrait=true&fitw=true&gridlines=false`;
    const pdfBlob = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName(`Acuse_Domiciliaciones_${hoy.toISOString().slice(0, 10)}.pdf`);
    
    const nombreAnalista = getNombreUsuario();
    MailApp.sendEmail({
      to: destinatarioEmail,
      subject: `Acuse de Domiciliaciones - ${nombreAnalista}`,
      htmlBody: `Se adjunta el acuse PDF con ${datos.domiciliaciones.length} registros generados por <b>${nombreAnalista}</b>.`,
      attachments: [pdfBlob]
    });
  } finally {
    plantilla.hideSheet();
    ss.deleteSheet(tempSheet);
  }
}

function formatearFechaEnEspanol(fecha) {
  const meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
  const dia = fecha.getDate();
  const mes = meses[fecha.getMonth()];
  const anio = fecha.getFullYear();
  return `${dia} de ${mes} de ${anio}`;
}

// --- FUNCIONES DE SINCRONIZACIÓN Y AUXILIARES ---

function procesoActualizacionIntegral(modoSilencioso = false) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaMaestra = ss.getSheetByName(HOJA_MAESTRA_LOCAL);
    if (!hojaMaestra) throw new Error(`No se encontró "${HOJA_MAESTRA_LOCAL}". Sincroniza primero.`);
    
    const datosMaestros = hojaMaestra.getDataRange().getValues();
    const headersMaestros = datosMaestros[0];
    const idxReferenciaMaestra = 4, idxMorosidadMaestra = 5, idxAsesorMaestra = 8, idxEstatusMaestra = headersMaestros.indexOf('ESTATUS');
    if (idxEstatusMaestra === -1) throw new Error(`No se encontró la columna "ESTATUS".`);

    const carteraMap = new Map();
    for (let i = 1; i < datosMaestros.length; i++) {
        const fila = datosMaestros[i];
        if (fila[idxReferenciaMaestra]) carteraMap.set(fila[idxReferenciaMaestra], {
            estatus: fila[idxEstatusMaestra], asesor: fila[idxAsesorMaestra], morosidad: fila[idxMorosidadMaestra]
        });
    }
    
    const resumen = _actualizarDatosDeCarteraEnHoja(HOJA_DOMICILIACIONES, carteraMap);
    
    if (!modoSilencioso) {
        SpreadsheetApp.getUi().alert('Proceso Completado', resumen, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return resumen;
  } catch (e) {
    if (!modoSilencioso) {
        SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    throw new Error(e.message);
  }
} 

function _actualizarDatosDeCarteraEnHoja(nombreHoja, carteraMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) return `Hoja "${nombreHoja}" no encontrada.`;
  
  const lastRow = hoja.getRange("A:A").getValues().filter(String).length;
  if (lastRow < 2) return `No hay datos para procesar en "${nombreHoja}".`;
  
  const datosHoja = hoja.getRange(1, 1, lastRow, hoja.getLastColumn()).getValues();
  const headers = datosHoja[0];
  const idxReferencia = headers.indexOf('REFERENCIA'), idxEstatus = headers.indexOf('ESTATUS (CARTERA)'),
        idxAsesor = headers.indexOf('ASESOR (CARTERA)'), idxMorosidad = headers.indexOf('MOROSIDAD (CARTERA)');

  if ([idxReferencia, idxEstatus, idxAsesor, idxMorosidad].includes(-1)) {
    return `Faltan columnas de destino en "${nombreHoja}". Revisa: REFERENCIA, ESTATUS (CARTERA), ASESOR (CARTERA), MOROSIDAD (CARTERA).`;
  }
  
  let actualizados = 0;
  const resultados = datosHoja.slice(1).map(fila => {
    const datosDeCartera = carteraMap.get(fila[idxReferencia]);
    if (datosDeCartera) {
      actualizados++;
      return [datosDeCartera.estatus, datosDeCartera.asesor, datosDeCartera.morosidad];
    }
    return [fila[idxEstatus], fila[idxAsesor], fila[idxMorosidad]];
  });
  
  if (resultados.length > 0) {
    hoja.getRange(2, idxEstatus + 1, resultados.length, 1).setValues(resultados.map(r => [r[0]]));
    hoja.getRange(2, idxAsesor + 1, resultados.length, 1).setValues(resultados.map(r => [r[1]]));
    hoja.getRange(2, idxMorosidad + 1, resultados.length, 1).setValues(resultados.map(r => [r[2]]));
  }
  return `En "${nombreHoja}", se actualizaron ${actualizados} de ${lastRow - 1} registros.`;
}

function sincronizarDatosCompletos(modoSilencioso = false) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const obtenerOCrearHoja = (nombreHoja) => ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
        
        const hojaMaestraLocal = obtenerOCrearHoja(HOJA_MAESTRA_LOCAL);
        const hojaClientesLocal = obtenerOCrearHoja(HOJA_CLIENTES);
        const hojaAsesoresLocal = obtenerOCrearHoja(HOJA_ASESORES);
        
        const carteraRemota = SpreadsheetApp.openById(CARTERA_ID);
        const hojaMaestraRemota = carteraRemota.getSheetByName('_TablaMaestra');
        if (!hojaMaestraRemota) throw new Error("No se encontró la hoja '_TablaMaestra' en Cartera.");
        
        const datosMaestros = hojaMaestraRemota.getDataRange().getValues();
        hojaMaestraLocal.clear();
        hojaMaestraLocal.getRange(1, 1, datosMaestros.length, datosMaestros[0].length).setValues(datosMaestros);
        
        const clientesUnicos = [...new Set(datosMaestros.slice(1).map(row => row[3]).filter(Boolean))].sort().map(c => [c]);
        hojaClientesLocal.clear();
        if (clientesUnicos.length > 0) hojaClientesLocal.getRange(1, 1, clientesUnicos.length, 1).setValues(clientesUnicos);

        const asesoresUnicos = [...new Set(
            datosMaestros.slice(1).map(row => {
                const nombreAsesor = row[8]; 
                if (typeof nombreAsesor === 'string' && nombreAsesor.trim() !== '') {
                    return nombreAsesor.toUpperCase().trim().replace(/\s+/g, ' ');
                }
                return null;
            }).filter(Boolean) 
        )].sort().map(a => [a]);
        
        hojaAsesoresLocal.clear();
        if (asesoresUnicos.length > 0) hojaAsesoresLocal.getRange(1, 1, asesoresUnicos.length, 1).setValues(asesoresUnicos);

        if (!modoSilencioso) {
             SpreadsheetApp.getActiveSpreadsheet().toast('Sincronización con Cartera completada.');
        }
        return "Sincronización Exitosa";
    } catch (e) {
        if (!modoSilencioso) {
            SpreadsheetApp.getUi().alert('Error en Sincronización', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
        }
        throw e;
    }
}

function buscarClientes(textoBusqueda) {
  if (!textoBusqueda || textoBusqueda.length < 3) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaClientesLocal = ss.getSheetByName(HOJA_CLIENTES);
  const todosLosClientes = (hojaClientesLocal && hojaClientesLocal.getLastRow() > 0)
    ? hojaClientesLocal.getDataRange().getValues().flat()
    : [...new Set(SpreadsheetApp.openById(CARTERA_ID).getSheetByName('_TablaMaestra').getRange("D2:D").getValues().flat().filter(String))].sort();
  const textoEnMinusculas = textoBusqueda.toLowerCase();
  return todosLosClientes.filter(cliente => cliente.toLowerCase().includes(textoEnMinusculas)).slice(0, 10);
}

function obtenerPropiedadesPorVariosClientes(nombresClientes) {
  if (!nombresClientes || nombresClientes.length === 0) return [];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_MAESTRA_LOCAL);
  const data = sheet.getDataRange().getValues();
  const propiedades = [];
  for (let i = 1; i < data.length; i++) {
    const fila = data[i];
    if (nombresClientes.includes(fila[3])) {
        propiedades.push({
            proyecto: fila[0], condominio: fila[1], loteCartera: fila[2],
            lote: (fila[2].split('-').pop() || fila[2]),
            propietario: fila[3], referencia: fila[4],
            asesor: (fila[8] || "").toString().trim().replace(/\s+/g, ' ')
        });
    }
  }
  return propiedades;
}

function getNombreUsuario() {
  try {
    const response = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v1/userinfo?alt=json', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
    return JSON.parse(response.getContentText()).name || Session.getActiveUser().getEmail();
  } catch (e) {
    return Session.getActiveUser().getEmail();
  }
}

function _encontrarProximaFilaVacia(sheet) {
  if (!sheet) return 1;
  return sheet.getRange("A:A").getValues().filter(String).length + 1;
}

// --------------------------------------------------------
// --- REPORTE DIARIO, KPIs Y AUTOMATIZACIÓN ---
// --------------------------------------------------------

function ejecutarReporteDiario() {
  const reporte = {
    sincronizacion: "",
    actualizacion: "",
    erroresValidacion: [],
    morosidad: [],
    stats: { totalRegistros: 0, montoTotal: 0, saludCartera: 100 } // KPIs
  };
  
  try { reporte.sincronizacion = sincronizarDatosCompletos(true); } catch (e) { reporte.sincronizacion = "Error: " + e.message; }
  try { reporte.actualizacion = procesoActualizacionIntegral(true); } catch (e) { reporte.actualizacion = "Error: " + e.message; }
  
  try {
    const resultados = _realizarValidacionesDeNegocio();
    reporte.erroresValidacion = resultados.errores;
    reporte.morosidad = resultados.morosos;
    reporte.stats = resultados.stats;
  } catch (e) {
    reporte.erroresValidacion.push({folio: "SISTEMA", mensaje: "Error validando: " + e.message});
  }
  
  _enviarCorreoReporteDiario(reporte);
}

function _realizarValidacionesDeNegocio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_DOMICILIACIONES);
  const data = hoja.getDataRange().getValues();
  const headers = data.shift();
  
  const colMap = {};
  headers.forEach((h, i) => colMap[h.toUpperCase().trim()] = i);
  
  const requiredCols = ["COMENTARIOS GPH", "ESTATUS (CARTERA)", "MOROSIDAD (CARTERA)", "FOLIO", "LOTE CARTERA", "MONTO A DOMICILIAR", "REGISTRADO POR"];
  const missing = requiredCols.filter(c => colMap[c] === undefined);
  if (missing.length > 0) throw new Error("Faltan columnas: " + missing.join(", "));
  
  // LIMPIEZA VISUAL: Solo Columnas A (1) hasta T (20)
  const lastRow = hoja.getLastRow();
  if (lastRow > 1) {
    hoja.getRange(2, 1, lastRow - 1, 20).setBackground(null);
  }

  const errores = [];
  const morosos = [];
  let montoTotal = 0;
  
  // PASO 1: Determinar el estado VIGENTE de cada lote basado en FECHA y ÍNDICE DE FILA
  // Si hay repetidos, gana el de fecha mayor. Si misma fecha, gana el último en la lista.
  const estadoVigentePorLote = {};

  data.forEach((row, i) => {
    const loteCartera = String(row[colMap["LOTE CARTERA"]]).trim();
    if (!loteCartera) return;

    const fechaRegistro = _extraerFechaDeRegistro(row[colMap["REGISTRADO POR"]]);
    const estatusLocal = _normalizarParaComparar(row[colMap["COMENTARIOS GPH"]]);
    const rowIndex = i; // Usamos el índice para desempatar fechas iguales (último gana)

    if (!estadoVigentePorLote[loteCartera]) {
      estadoVigentePorLote[loteCartera] = { fecha: fechaRegistro, estatus: estatusLocal, rowIndex: rowIndex };
    } else {
      const actual = estadoVigentePorLote[loteCartera];
      // Si la fecha nueva es mayor, o es igual pero está más abajo en el sheet -> actualizamos
      if (fechaRegistro > actual.fecha || (fechaRegistro.getTime() === actual.fecha.getTime() && rowIndex > actual.rowIndex)) {
        estadoVigentePorLote[loteCartera] = { fecha: fechaRegistro, estatus: estatusLocal, rowIndex: rowIndex };
      }
    }
  });

  // PASO 2: Validar
  data.forEach((row, i) => {
    const filaReal = i + 2;
    const folio = row[colMap["FOLIO"]] || "Sin Folio";
    const loteCartera = String(row[colMap["LOTE CARTERA"]]).trim();
    
    // Solo validamos si este registro es el VIGENTE para ese lote
    const esVigente = estadoVigentePorLote[loteCartera] && estadoVigentePorLote[loteCartera].rowIndex === i;
    
    const comentariosGPH_Norm = _normalizarParaComparar(row[colMap["COMENTARIOS GPH"]]);
    const estatusCartera_Norm = _normalizarParaComparar(row[colMap["ESTATUS (CARTERA)"]]);
    
    // KPI Monto: Sumamos solo si es vigente y es domiciliación
    if (esVigente && comentariosGPH_Norm === "DOMICILIACION") {
        const monto = parseFloat(row[colMap["MONTO A DOMICILIAR"]]);
        if (!isNaN(monto)) montoTotal += monto;
    }

    let tieneError = false;

    // Solo validamos estatus si es el registro vigente
    if (esVigente) {
        if (comentariosGPH_Norm === "DOMICILIACION") {
          if (!estatusCartera_Norm.includes("DOMICILIACION") && !estatusCartera_Norm.includes("FCD")) {
            errores.push({
              folio: folio,
              lote: loteCartera,
              tipo: "Estatus Incorrecto",
              detalle: `GPH (Vigente): "${row[colMap["COMENTARIOS GPH"]]}" vs Cartera: "${row[colMap["ESTATUS (CARTERA)"]]}"`
            });
            tieneError = true;
          }
        }
        
        if (comentariosGPH_Norm === "CANCELACION") {
          if (estatusCartera_Norm.includes("DOMICILIACION") || estatusCartera_Norm.includes("FCD")) {
             // Ya sabemos que es el vigente y es cancelación, por tanto NO hay una reactivación posterior.
             // Si cartera sigue activa, es error.
              errores.push({
                folio: folio,
                lote: loteCartera,
                tipo: "Estatus Incorrecto (Cancelado)",
                detalle: `GPH (Vigente): Cancelación vs Cartera: Activo.`
              });
              tieneError = true;
          }
        }
        
        // Morosidad solo si es Domiciliación Vigente
        if (comentariosGPH_Norm === "DOMICILIACION") {
            const morosidad = parseFloat(row[colMap["MOROSIDAD (CARTERA)"]]) || 0;
            if (morosidad >= 2) {
                morosos.push({
                    folio: folio,
                    lote: loteCartera,
                    morosidad: morosidad
                });
                tieneError = true;
            }
        }
    }

    // PINTAR: Solo columnas 1 a 20 (A-T)
    if (tieneError) {
        hoja.getRange(filaReal, 1, 1, 20).setBackground('#ffe6e6');
    }
  });
  
  const totalRegistros = Object.keys(estadoVigentePorLote).length; // Contamos lotes únicos vigentes
  const totalIncidentes = errores.length + morosos.length;
  const salud = totalRegistros > 0 ? ((totalRegistros - totalIncidentes) / totalRegistros) * 100 : 100;
  
  return { 
      errores, 
      morosos, 
      stats: { 
          totalRegistros, 
          montoTotal, 
          saludCartera: Math.max(0, Math.round(salud)) 
      } 
  };
}

function _extraerFechaDeRegistro(texto) {
  // Formato esperado: "email@dominio.com 29/01/2026"
  if (!texto) return new Date(0); // Fecha muy antigua si no hay dato
  
  // Buscar patrón dd/mm/yyyy
  const match = String(texto).match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (match) {
    // Mes en JS es 0-11
    return new Date(match[3], match[2] - 1, match[1]);
  }
  return new Date(0);
}

function _normalizarParaComparar(texto) {
  if (!texto) return "";
  return String(texto).toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}

function _enviarCorreoReporteDiario(reporte) {
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const formatearDinero = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN' });
  const montoStr = formatearDinero.format(reporte.stats.montoTotal);
  
  const colorSalud = reporte.stats.saludCartera >= 90 ? '#28a745' : (reporte.stats.saludCartera >= 70 ? '#ffc107' : '#dc3545');

  const styleTable = 'width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; margin-top: 5px;';
  const styleTh = 'background-color: #f8f9fa; border-bottom: 2px solid #dee2e6; padding: 10px; text-align: left; color: #495057;';
  const styleTd = 'border-bottom: 1px solid #dee2e6; padding: 10px; color: #555;';
  const styleCard = 'border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px; margin-bottom: 25px; background-color: #ffffff; box-shadow: 0 2px 4px rgba(0,0,0,0.05);';
  const styleH4 = 'margin-top: 0; color: #0056b3; border-bottom: 1px solid #eee; padding-bottom: 10px; margin-bottom: 15px; font-size: 16px;';
  const styleKpiBox = 'display: inline-block; width: 30%; text-align: center; border-right: 1px solid #eee;';
  
  let htmlBody = `
    <div style="font-family: 'Segoe UI', Arial, sans-serif; color: #333; max-width: 850px; margin: auto; background-color: #f4f6f9; padding: 30px; border-radius: 10px;">
      
      <div style="background-color: #004494; color: white; padding: 20px; border-radius: 8px 8px 0 0; text-align: center;">
        <h2 style="margin: 0; font-size: 24px;">📊 Reporte Diario de Control</h2>
        <p style="margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;">${fecha}</p>
      </div>

      <div style="background-color: white; padding: 20px; border-radius: 0 0 8px 8px; margin-bottom: 25px; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
        <table style="width: 100%;">
          <tr>
            <td style="${styleKpiBox}">
              <span style="display: block; font-size: 12px; color: #666; text-transform: uppercase; letter-spacing: 1px;">Lotes Activos</span>
              <strong style="font-size: 28px; color: #333;">${reporte.stats.totalRegistros}</strong>
            </td>
            <td style="${styleKpiBox}">
              <span style="display: block; font-size: 12px; color: #666; text-transform: uppercase; letter-spacing: 1px;">Salud Cartera</span>
              <strong style="font-size: 28px; color: ${colorSalud};">${reporte.stats.saludCartera}%</strong>
            </td>
            <td style="text-align: center; width: 30%;">
              <span style="display: block; font-size: 12px; color: #666; text-transform: uppercase; letter-spacing: 1px;">Monto Activo Mensual</span>
              <strong style="font-size: 28px; color: #0056b3;">${montoStr}</strong>
            </td>
          </tr>
        </table>
      </div>

      <div style="${styleCard}">
        <h4 style="${styleH4}">📡 Estado del Sistema</h4>
        <p style="margin: 5px 0;">🔄 <b>Sincronización:</b> ${reporte.sincronizacion}</p>
        <p style="margin: 5px 0;">⭐ <b>Actualización Integral:</b> ${reporte.actualizacion}</p>
      </div>
  `;
  
  if (reporte.erroresValidacion.length > 0) {
    htmlBody += `<div style="${styleCard} border-left: 5px solid #dc3545;">
      <h4 style="${styleH4} color: #dc3545;">⚠️ Alertas de Estatus (${reporte.erroresValidacion.length})</h4>
      <p style="font-size: 13px; color: #666;">Se han marcado en <b>rojo</b> en el Excel las siguientes discrepancias:</p>
      <table style="${styleTable}">
        <thead><tr><th style="${styleTh}">Folio</th><th style="${styleTh}">Lote Cartera</th><th style="${styleTh}">Tipo Error</th><th style="${styleTh}">Detalle</th></tr></thead>
        <tbody>`;
    reporte.erroresValidacion.forEach(e => {
      htmlBody += `<tr><td style="${styleTd} font-weight: bold;">${e.folio}</td><td style="${styleTd}">${e.lote}</td><td style="${styleTd} color: #dc3545;">${e.tipo}</td><td style="${styleTd}">${e.detalle}</td></tr>`;
    });
    htmlBody += `</tbody></table></div>`;
  } else {
    htmlBody += `<div style="${styleCard} border-left: 5px solid #28a745;"><h4 style="${styleH4} color: #28a745;">✅ Estatus de Cartera</h4><p style="color: #28a745; margin: 0;">Todos los estatus coinciden correctamente.</p></div>`;
  }

  if (reporte.morosidad.length > 0) {
    htmlBody += `<div style="${styleCard} border-left: 5px solid #ffc107;">
      <h4 style="${styleH4} color: #d39e00;">📉 Alertas de Morosidad (${reporte.morosidad.length})</h4>
      <p style="font-size: 13px; color: #666;">Domiciliaciones activas con adeudo crítico (>= 2 meses):</p>
      <table style="${styleTable}">
        <thead><tr><th style="${styleTh}">Folio</th><th style="${styleTh}">Lote Cartera</th><th style="${styleTh}">Adeudo</th></tr></thead>
        <tbody>`;
    reporte.morosidad.forEach(m => {
      htmlBody += `<tr><td style="${styleTd} font-weight: bold;">${m.folio}</td><td style="${styleTd}">${m.lote}</td><td style="${styleTd} color: #d39e00; font-weight: bold; font-size: 14px;">${m.morosidad}</td></tr>`;
    });
    htmlBody += `</tbody></table></div>`;
  } else {
    htmlBody += `<div style="${styleCard} border-left: 5px solid #28a745;"><h4 style="${styleH4} color: #28a745;">✅ Control de Cobranza</h4><p style="color: #28a745; margin: 0;">Ninguna domiciliación activa presenta morosidad crítica.</p></div>`;
  }
  
  htmlBody += `
    <div style="text-align: center; margin-top: 30px; font-size: 11px; color: #999;">
      <p>Reporte generado automáticamente por el Sistema de Domiciliaciones GPH.</p>
    </div>
    </div>
  `;
  
  const destinatarios = CORREOS_REPORTE.join(',');
  if(destinatarios.trim() !== "") {
    MailApp.sendEmail({
        to: destinatarios,
        subject: `📊 Reporte Diario de Domiciliaciones - ${fecha}`,
        htmlBody: htmlBody
    });
  }
}