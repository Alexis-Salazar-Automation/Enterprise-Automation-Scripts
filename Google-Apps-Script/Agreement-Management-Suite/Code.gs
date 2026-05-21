// ID de tu hoja de cálculo "Cartera".
const CARTERA_ID = 'ID_PROTEGIDO_POR_SEGURIDAD';

// Nombres de las hojas locales.
const HOJA_CLIENTES = '_Listas_Clientes';
const HOJA_ASESORES = '_Listas_Asesores';
const HOJA_MAESTRA_LOCAL = '_ListasMaestras';
const HOJA_CORREOS_LOCAL = '_Listas_Correos';
const HOJA_CONCENTRADO = 'Concentrado';
const HOJA_CONVENIOS_ACTIVOS = 'Convenios Activos';
const HOJA_NEGOCIACIONES = 'Negociacion-PP';
const HOJA_PLANTILLA_ACUSE = 'PlantillaAcuse';

// --- CONFIGURACIÓN DE CORREO PARA ACUSES ---
const DESTINATARIO_FIJO_ACUSE = "Operations Manager";
const CORREOS_SUPERVISION_CC = []; // Añade correos aquí si lo necesitas. Ej: ["supervisor@email.com"]
const CORREOS_ADICIONALES_VENCIMIENTO_CC = ["ejecutivo1@empresa.com,ejecutivo2@empresa.com,ejecutivo3@empresa.com"];
const CORREOS_ADICIONALES_NEGOCIACION_CC = ["ejecutivo1@empresa.com,ejecutivo2@empresa.com,ejecutivo3@empresa.com"];
const CORREO_REPORTE_DIARIO = ["ejecutivo1@empresa.com,ejecutivo2@empresa.com,ejecutivo3@empresa.com"]; 

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- MÉTODO CORREGIDO PARA VERIFICAR PERMISOS SIN AUTORIZACIÓN ---
  // Verificamos si el usuario actual tiene permiso para editar la celda A1.
  // Esto devuelve 'true' para Editores/Propietarios y 'false' para Lectores.
  // Este método no requiere permisos especiales y funciona en un onOpen.
  const puedeEditar = ss.getRange('A1').canEdit();

  // --- MENÚ COMPLETO PARA EDITORES Y PROPIETARIOS ---
  if (puedeEditar) {
    // Menú de Trámites
    ui.createMenu('⚙️ Trámites')
      .addItem('➕ Registrar Nuevo Trámites', 'mostrarFormulario')
      .addSeparator()
      .addItem('⭐ Proceso de Actualización Integral', 'procesoActualizacionIntegral')
      .addItem('🔄 Actualizar Estatus en Concentrado', 'actualizarEstatusGeneralConcentrado')
      .addSeparator()
      .addItem('🔄 Actualizar Convenios Vencidos', 'actualizarConvenios')
      .addItem('🔄 Actualizar Negociaciones Vencidas', 'actualizacionNegociacionesVencidas')
      .addSeparator()
      .addItem('✉️ Enviar Recordatorios de Vencimiento', 'enviarRecordatoriosDeVencimiento')
      .addItem('🔔 Enviar Recordatorios de Negociación', 'recordatorioNegociacionesPP')
      .addSeparator()
      .addItem('🔄 Sincronizar Datos de Cartera', 'sincronizarDatosCompletos')
      .addSeparator()
      .addItem('📄 Generar Acuse de Convenios Pendientes', 'mostrarSelectorAnalista')
      .addToUi();

    // Menú de Reportes y Búsquedas
    ui.createMenu('📊 Reportes y Búsquedas')
      .addItem('📈 Generar Reporte Mensual', 'mostrarDialogoReporte')
      .addSeparator()
      .addItem('🔎 Buscador Inteligente de Trámites', 'mostrarDialogoBusqueda')
      .addItem('📜 Ver Historial de Lote', 'mostrarDialogoHistorial')
      .addToUi();

    ui.createMenu('🔍 Filtros')
    .addItem('⚡ Asistente de Filtrado Dinámico', 'mostrarDialogoFiltro')
    .addItem('⏩ Filtro Rápido (en la hoja actual)', 'mostrarDialogoFiltroRapido')
    .addItem('❌ Quitar Filtros de la Hoja Actual', 'quitarFiltrosDeHoja')
    .addToUi();
  } 
  // --- MENÚ LIMITADO PARA USUARIOS DE SOLO LECTURA ---
  else {
    ui.createMenu('📊 Búsquedas')
      .addItem('🔎 Buscador Inteligente de Trámites', 'mostrarDialogoBusqueda')
      .addToUi();
  }
}

function sincronizarDatosCompletos() {
    // Se eliminan las alertas de UI que causaban el error en el trigger.
    // El reporte se encargará de notificar el éxito o fracaso.
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const obtenerOCrearHoja = (nombreHoja) => {
          let hoja = ss.getSheetByName(nombreHoja);
          if (!hoja) hoja = ss.insertSheet(nombreHoja);
          return hoja;
        };
        const hojaMaestraLocal = obtenerOCrearHoja(HOJA_MAESTRA_LOCAL);
        const hojaClientesLocal = obtenerOCrearHoja(HOJA_CLIENTES);
        const hojaAsesoresLocal = obtenerOCrearHoja(HOJA_ASESORES);
        const hojaCorreosLocal = obtenerOCrearHoja(HOJA_CORREOS_LOCAL);
        const carteraRemota = SpreadsheetApp.openById(CARTERA_ID);
        const hojaMaestraRemota = carteraRemota.getSheetByName('_TablaMaestra');
        if (!hojaMaestraRemota) throw new Error("No se encontró la hoja '_TablaMaestra' en el archivo de Cartera.");
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
                    return nombreAsesor.trim().replace(/\s+/g, ' ');
                }
                return null;
            }).filter(Boolean)
        )].sort().map(a => [a]);
        
        hojaAsesoresLocal.clear();
        if (asesoresUnicos.length > 0) hojaAsesoresLocal.getRange(1, 1, asesoresUnicos.length, 1).setValues(asesoresUnicos);
        
        const hojaCorreosRemota = carteraRemota.getSheetByName('Correos');
        if (hojaCorreosRemota) {
            const datosCorreos = hojaCorreosRemota.getDataRange().getValues();
            hojaCorreosLocal.clear();
            if (datosCorreos.length > 0) hojaCorreosLocal.getRange(1, 1, datosCorreos.length, datosCorreos[0].length).setValues(datosCorreos);
        }
        // No devuelve nada, el éxito se asume si no hay error.
    } catch (e) {
        // Si hay un error, lo "lanza" para que la función principal lo capture.
        throw new Error(e.message);
    }
}

function mostrarFormulario() {
  const html = HtmlService.createHtmlOutputFromFile('FormularioConvenio').setWidth(750).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Registrar Nuevo Trámite');
}

function mostrarSelectorAnalista() {
    const html = HtmlService.createHtmlOutputFromFile('SelectorAnalista').setWidth(400).setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'Seleccionar Analista');
}

function obtenerAnalistasConPendientes() {
    const sheetConcentrado = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONCENTRADO);
    const data = sheetConcentrado.getDataRange().getValues();
    const headers = data.shift();
    const folioIndex = headers.indexOf("Folio");
    const acuseIndex = headers.indexOf("Acuse Generado");
    const registradoPorIndex = headers.indexOf("Registrado por");

    if (registradoPorIndex === -1) {
        throw new Error('No se encontró la columna "Registrado por" en la hoja "Concentrado".');
    }

    const analistasPendientes = new Set();
    data.forEach(row => {
        const folio = row[folioIndex];
        const acuseStatus = row[acuseIndex];
        const registradoPor = row[registradoPorIndex];
        if (String(folio).startsWith('CONV-') && acuseStatus === "" && registradoPor) {
            analistasPendientes.add(registradoPor);
        }
    });

    return Array.from(analistasPendientes).sort();
}

function obtenerDatosParaFormulario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const asesores = ss.getSheetByName(HOJA_ASESORES).getDataRange().getValues().flat();
  return { asesores };
}

function buscarClientes(textoBusqueda) {
  if (!textoBusqueda || textoBusqueda.length < 3) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaClientesLocal = ss.getSheetByName(HOJA_CLIENTES);
  let todosLosClientes = [];
  if (hojaClientesLocal && hojaClientesLocal.getLastRow() > 0) {
    todosLosClientes = hojaClientesLocal.getRange(1, 1, hojaClientesLocal.getLastRow(), 1).getValues().flat().filter(String);
  } else {
    try {
      const carteraRemota = SpreadsheetApp.openById(CARTERA_ID);
      const hojaMaestraRemota = carteraRemota.getSheetByName('_TablaMaestra');
      if (!hojaMaestraRemota) throw new Error("No se pudo encontrar la hoja '_TablaMaestra' en el archivo remoto.");
      todosLosClientes = [...new Set(hojaMaestraRemota.getRange("D2:D").getValues().flat().filter(String))].sort();
    } catch (e) {
      Logger.log("Fallo crítico al leer la hoja maestra remota: " + e.message);
      throw new Error("No se pudo acceder a la fuente de datos principal.");
    }
  }
  const textoEnMinusculas = textoBusqueda.toLowerCase();
  return todosLosClientes.filter(cliente => cliente.toLowerCase().includes(textoEnMinusculas)).slice(0, 10);
}

function obtenerPropiedadesPorVariosClientes(nombresClientes) {
  if (!nombresClientes || nombresClientes.length === 0) return [];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_MAESTRA_LOCAL);
  const data = sheet.getDataRange().getValues();
  const propiedades = [];
  const timeZone = Session.getScriptTimeZone();
  for (let i = 1; i < data.length; i++) {
    const fila = data[i];
    if (nombresClientes.includes(fila[3])) {
      
      let nombreAsesor = fila[8] || '';
      if (typeof nombreAsesor === 'string') {
        nombreAsesor = nombreAsesor.trim().replace(/\s+/g, ' ');
      }

      propiedades.push({
        proyecto: fila[0], condominio: fila[1], lote: fila[2], propietario: fila[3], referencia: fila[4],
        morosidad: fila[5],
        ultimoPago: fila[6] instanceof Date ? Utilities.formatDate(fila[6], timeZone, "dd-MM-yyyy") : (fila[6] || 'N/A'),
        ultimaGestion: fila[7] instanceof Date ? Utilities.formatDate(fila[7], timeZone, "dd-MM-yyyy") : (fila[7] || 'N/A'),
        asesor: nombreAsesor
      });
    }
  }
  return propiedades;
}

function _normalizarNombre(nombre) {
  if (typeof nombre !== 'string' || !nombre) {
    return "";
  }
  return nombre
    .toLowerCase()
    .normalize("NFD") // Descompone caracteres en base + acento
    .replace(/[\u0300-\u036f]/g, "") // Elimina los acentos
    .replace(/ñ/g, 'n') // Reemplaza la ñ específicamente
    .trim()
    .replace(/\s+/g, ' '); // Reemplaza múltiples espacios por uno solo
}

function obtenerCorreosPorAsesor(nombreAsesor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CORREOS_LOCAL);
  if (!sheet) return {};

  const asesorBuscadoNormalizado = _normalizarNombre(nombreAsesor);
  if (!asesorBuscadoNormalizado) return {}; // Si el nombre de entrada está vacío, no buscar.

  const data = sheet.getDataRange().getValues();
  
  // Recorremos la lista de correos desde la segunda fila (asumiendo que la primera es encabezado)
  for (let i = 1; i < data.length; i++) {
    const nombreEnLista = data[i][0]; // Nombre del asesor en la hoja de Correos
    const nombreEnListaNormalizado = _normalizarNombre(nombreEnLista);

    // --- La Magia de la Búsqueda Fuerte ---
    // Comparamos si el nombre más largo CONTIENE al nombre más corto.
    // Esto resuelve casos como "Alexis Salazar Flores" vs "Alexis Salazar".
    if (nombreEnListaNormalizado && asesorBuscadoNormalizado && 
        (nombreEnListaNormalizado.includes(asesorBuscadoNormalizado) || asesorBuscadoNormalizado.includes(nombreEnListaNormalizado))) {
      
      // ¡Coincidencia encontrada! Devolvemos los correos y terminamos.
      return {
        correoEjecutivo: data[i][1],
        correoSupervisor: data[i][3],
        correoCoordinador: data[i][5]
      };
    }
  }
}

function _encontrarProximaFilaVacia(sheet) {
  if (!sheet) return 1;
  return sheet.getRange("A:A").getValues().filter(String).length + 1;
}

function getNombreUsuario() {
  try {
    const url = 'https://www.googleapis.com/oauth2/v1/userinfo?alt=json';
    const options = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } };
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    return data.name || Session.getActiveUser().getEmail();
  } catch (e) {
    Logger.log("No se pudo obtener el nombre del usuario vía UrlFetch: " + e.message);
    return Session.getActiveUser().getEmail();
  }
}

function guardarTramite(datosFormulario) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const { commonDetails, selectedProperties } = datosFormulario;
        const tipoRegistro = commonDetails.tipoRegistro;
        const usuarioActualEmail = Session.getActiveUser().getEmail();
        const nombreUsuario = getNombreUsuario();
        const sheetConcentrado = ss.getSheetByName(HOJA_CONCENTRADO);
        if (!sheetConcentrado) throw new Error(`La hoja "${HOJA_CONCENTRADO}" no fue encontrada.`);
        
        let ID_PREFIX, lastRelevantID, nextNum = 1;
        const colAValues = sheetConcentrado.getRange("A1:A").getValues().flat().reverse();
        if (tipoRegistro === 'convenio') {
            ID_PREFIX = 'CONV';
            lastRelevantID = colAValues.find(id => id && String(id).startsWith(ID_PREFIX + '-'));
        } else {
            ID_PREFIX = (commonDetails.comentario === 'Promesa de pago') ? 'PP' : 'NEG';
            lastRelevantID = colAValues.find(id => id && (String(id).startsWith('NEG-') || String(id).startsWith('PP-')));
        }
        if (lastRelevantID) {
            const num = parseInt(lastRelevantID.split('-')[2], 10);
            if (!isNaN(num)) nextNum = num + 1;
        }
        const year = new Date().getFullYear();
        const ID_UNICO = `${ID_PREFIX}-${year}-${String(nextNum).padStart(4, '0')}`;
        
        const hojaDestinoSecundariaNombre = tipoRegistro === 'convenio' ? HOJA_CONVENIOS_ACTIVOS : HOJA_NEGOCIACIONES;
        
        let filasParaConcentrado = [];
        let filasParaHojaSecundaria = [];

        if (tipoRegistro === 'convenio') {
            filasParaConcentrado = selectedProperties.map((prop, index) => {
                const esPrimeraFila = (index === 0);
                const tipo = tipoRegistro === 'convenio' ? 'NUEVO CONVENIO' : 'NUEVA NEGOCIACIÓN/PP';
                registrarEvento(prop.lote, tipo, `Se registró el trámite para el propietario ${prop.propietario}.`, ID_UNICO);
                return [
                    ID_UNICO, commonDetails.entregaConvenio, commonDetails.fechaRecibo, prop.propietario, prop.proyecto,
                    prop.referencia, prop.condominio, (prop.lote.split('-').pop() || prop.lote), prop.lote,
                    // --- INICIO DE LA CORRECCIÓN ---
                    prop.asesor, // Se usa el asesor específico de la propiedad (prop.asesor)
                    // --- FIN DE LA CORRECCIÓN ---
                    commonDetails.fechaInicio, commonDetails.fechaFin, 
                    esPrimeraFila ? commonDetails.parciales : "-", esPrimeraFila ? commonDetails.adeudo : "-", esPrimeraFila ? commonDetails.monto : "-",
                    prop.morosidad, prop.ultimaGestion, prop.ultimoPago, "", "", commonDetails.comentario,
                    commonDetails.correoEjecutivo, commonDetails.correoSupervisor, commonDetails.correoCoordinador,
                    "", nombreUsuario, "", "", "", "", ""
                ];
            });
            filasParaHojaSecundaria = filasParaConcentrado;

        } else {
            filasParaHojaSecundaria = selectedProperties.map((prop, index) => {
                const esPrimeraFila = (index === 0);
                return [
                    ID_UNICO, commonDetails.entregaConvenio, commonDetails.fechaRecibo, prop.propietario, prop.proyecto, 
                    prop.referencia, prop.condominio, (prop.lote.split('-').pop() || prop.lote), prop.lote,
                    // --- INICIO DE LA CORRECCIÓN ---
                    prop.asesor, // Se usa el asesor específico de la propiedad (prop.asesor)
                    // --- FIN DE LA CORRECCIÓN ---
                    commonDetails.fechaInicio, commonDetails.fechaFin, esPrimeraFila ? commonDetails.monto : "-", 
                    prop.morosidad, prop.ultimaGestion, prop.ultimoPago, "", commonDetails.comentario, 
                    usuarioActualEmail, nombreUsuario
                ];
            });
            
            filasParaConcentrado = selectedProperties.map((prop, index) => {
                const esPrimeraFila = (index === 0);
                return [
                    ID_UNICO, commonDetails.entregaConvenio, commonDetails.fechaRecibo, prop.propietario, prop.proyecto,
                    prop.referencia, prop.condominio, (prop.lote.split('-').pop() || prop.lote), prop.lote,
                    // --- INICIO DE LA CORRECCIÓN ---
                    prop.asesor, // Se usa el asesor específico de la propiedad (prop.asesor)
                    // --- FIN DE LA CORRECCIÓN ---
                    commonDetails.fechaInicio, commonDetails.fechaFin, "-", "-", 
                    esPrimeraFila ? commonDetails.monto : "-", prop.morosidad, prop.ultimaGestion, prop.ultimoPago,
                    "", "", commonDetails.comentario, usuarioActualEmail, "", "", "", nombreUsuario, "", "", "", "", ""
                ];
            });
        }

        if (filasParaConcentrado.length > 0) {
            const proximaFilaConcentrado = _encontrarProximaFilaVacia(sheetConcentrado);
            sheetConcentrado.getRange(proximaFilaConcentrado, 1, filasParaConcentrado.length, filasParaConcentrado[0].length).setValues(filasParaConcentrado);
            
            let hojaSecundaria = ss.getSheetByName(hojaDestinoSecundariaNombre);
            if (!hojaSecundaria) {
                hojaSecundaria = ss.insertSheet(hojaDestinoSecundariaNombre);
                const encabezados = (tipoRegistro === 'convenio') 
                    ? [[ 'Folio', 'Entrega del convenio', 'Fecha Recibo', 'Nombre Propietario', 'Desarrollo', 'Referencia', 'Condominio', 'Lote', 'Lote Cartera', 'Ejecutivo Asignado', 'Fecha Inicio', 'Fecha Fin', 'Parcialidades', 'Monto', 'Monto Adeudo Total', 'Meses Morosidad', 'Ultima Gestion', 'Ultimo pago Registrado', 'Vencimiento', 'Estatus convenio', 'Comentarios', 'Correo ejecutivo', 'Correo supervisor', 'Correo coordinador', 'Acuse Generado', 'Registrado por', 'Recordatorio Enviado', 'Estatus (Cartera)', 'Asesor (Cartera)', 'Morosidad (Cartera)', 'Cierre Convenio' ]]
                    : [[ 'Folio', 'Entrega del documento', 'Fecha Recibo', 'Nombre Propietario', 'Desarrollo', 'Referencia', 'Condominio', 'Lote', 'Lote Cartera', 'Ejecutivo Asignado', 'Fecha Inicio', 'Fecha Fin', 'Monto Adeudo Total', 'Meses Morosidad', 'Ultima Gestion', 'Ultimo pago Registrado', 'Vencimiento', 'Comentarios', 'Correo Analista', 'Registrado por' ]];
                hojaSecundaria.getRange(1, 1, 1, encabezados[0].length).setValues(encabezados).setFontWeight('bold');
            }
            const proximaFilaSecundaria = _encontrarProximaFilaVacia(hojaSecundaria);
            hojaSecundaria.getRange(proximaFilaSecundaria, 1, filasParaHojaSecundaria.length, filasParaHojaSecundaria[0].length).setValues(filasParaHojaSecundaria);
        }
        
        return `¡Éxito! Se ha registrado el trámite con el Folio: ${ID_UNICO}.`;
    } catch (e) {
        Logger.log(e);
        return `Error: ${e.message}`;
    } finally {
        lock.releaseLock();
    }
}


function generarAcusePorLote(selectedAnalistas) {
    const ui = SpreadsheetApp.getUi();
    
    if (!selectedAnalistas || !Array.isArray(selectedAnalistas) || selectedAnalistas.length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast("Error: No se seleccionó ningún analista.", "Fallo", 5);
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetConcentrado = ss.getSheetByName(HOJA_CONCENTRADO);
    const sheetActivos = ss.getSheetByName(HOJA_CONVENIOS_ACTIVOS);
    const emailUsuarioActual = Session.getActiveUser().getEmail();

    const dataConcentrado = sheetConcentrado.getDataRange().getValues();
    const headers = dataConcentrado[0];
    
    const colIndex = headers.reduce((acc, header, i) => { acc[header] = i; return acc; }, {});
    
    // <-- CAMBIO CLAVE AQUÍ: Añadimos "Fecha Recibo" a las columnas requeridas.
    const requiredColumns = ["Folio", "Acuse Generado", "Desarrollo", "Condominio", "Lote Cartera", "Entrega del convenio", "Fecha Recibo", "Registrado por"];
    for (const col of requiredColumns) {
        if (colIndex[col] === undefined) {
             ui.alert("Error de Configuración", `No se encontró la columna requerida "${col}" en la hoja "${HOJA_CONCENTRADO}".`, ui.ButtonSet.OK);
             return;
        }
    }

    const conveniosParaPDF = [];
    const foliosAActualizar = new Set(); 

    dataConcentrado.slice(1).forEach((row) => {
        const folio = row[colIndex["Folio"]];
        const acuseStatus = row[colIndex["Acuse Generado"]];
        const registradoPor = row[colIndex["Registrado por"]];
        
        if (String(folio).startsWith('CONV-') && acuseStatus === "" && selectedAnalistas.includes(registradoPor)) {
            // <-- CAMBIO CLAVE AQUÍ: Añadimos el nuevo dato al objeto.
            conveniosParaPDF.push({
                folio: folio,
                desarrollo: row[colIndex["Desarrollo"]],
                condominio: row[colIndex["Condominio"]],
                lote: row[colIndex["Lote Cartera"]],
                fechaEntrega: row[colIndex["Entrega del convenio"]],
                fechaRecibo: row[colIndex["Fecha Recibo"]] // <--- Nueva línea
            });
            foliosAActualizar.add(folio);
        }
    });

    if (conveniosParaPDF.length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast("No se encontraron convenios pendientes para los analistas seleccionados.", "Información", 5);
      return;
    }

    try {
        const nombreAnalistaParaPDF = selectedAnalistas.join(' y ');
        const datosParaAcuse = {
            folioUnico: "Lote-" + new Date().toISOString().slice(0, 10),
            nombreAnalista: nombreAnalistaParaPDF,
            convenios: conveniosParaPDF
        };
        
        // El resto de la función que envía el correo y actualiza la hoja no cambia.
        generarYEnviarAcusePDF(datosParaAcuse, emailUsuarioActual);
        // ... (el resto del código sigue igual)
        const ahora = new Date();
        const fechaMarca = "Generado el " + ahora.toLocaleString();
        const actualizarHoja = (sheet) => {
          if (!sheet || sheet.getLastRow() < 2) return;
          const folioColNum = colIndex["Folio"] + 1;
          const acuseColNum = colIndex["Acuse Generado"] + 1;
          const foliosEnHoja = sheet.getRange(2, folioColNum, sheet.getLastRow() - 1, 1).getValues();
          const acusesEnHoja = sheet.getRange(2, acuseColNum, sheet.getLastRow() - 1, 1).getValues();
          for (let i = 0; i < foliosEnHoja.length; i++) {
            if (foliosAActualizar.has(foliosEnHoja[i][0])) {
              acusesEnHoja[i][0] = fechaMarca;
            }
          }
          sheet.getRange(2, acuseColNum, acusesEnHoja.length, 1).setValues(acusesEnHoja);
        };
        actualizarHoja(sheetConcentrado);
        actualizarHoja(sheetActivos);
        SpreadsheetApp.getActiveSpreadsheet().toast(`¡Éxito! Se generó un acuse con ${conveniosParaPDF.length} convenios.`, "Proceso Completado", 10);
    } catch (e) {
        Logger.log(e);
        ui.alert("Error", "Ocurrió un error al generar el PDF o enviar el correo: " + e.message, ui.ButtonSet.OK);
    }
}

function formatearFechaEnEspanol(fecha) {
  const meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
  const dia = fecha.getDate();
  const mes = meses[fecha.getMonth()];
  const anio = fecha.getFullYear();
  return `${dia} de ${mes} de ${anio}`;
}

function generarYEnviarAcusePDF(datos, destinatarioEmail) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const plantillaSheet = ss.getSheetByName(HOJA_PLANTILLA_ACUSE);
    if (!plantillaSheet) throw new Error(`No se encontró la hoja plantilla "${HOJA_PLANTILLA_ACUSE}"`);

    plantillaSheet.showSheet();
    
    let tempSheet;
    
    try {
        tempSheet = plantillaSheet.copyTo(ss).setName("Acuse Temporal " + new Date().getTime());
        const hoy = new Date();
        const fechaFormateada = formatearFechaEnEspanol(hoy);
        const semanaDelAnio = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "w");

        tempSheet.createTextFinder('{{fecha_acuse}}').replaceAllWith(fechaFormateada);
        tempSheet.createTextFinder('{{semana_del_anio}}').replaceAllWith("Semana " + semanaDelAnio);
        tempSheet.createTextFinder('{{nombre_analista}}').replaceAllWith(datos.nombreAnalista);
        
        // --- INICIO DE LA CORRECCIÓN ---
        // Se corrige el nombre de la variable a DESTINATARIO_FIJO_ACUSE (con 'A' mayúscula)
        tempSheet.createTextFinder('{{destinatario}}').replaceAllWith(DESTINATARIO_FIJO_ACUSE);
        // --- FIN DE LA CORRECCIÓN ---
        
        const tablaInicioFila = 14; 
        const timeZone = Session.getScriptTimeZone();

        const datosTabla = datos.convenios.map((convenio, index) => {
            const folioCorto = convenio.folio.split('-').pop() || "0000";
            const loteNumero = parseInt(convenio.lote.split('-').pop(), 10) || convenio.lote;
            
            const fechaEntregaStr = convenio.fechaEntrega ? Utilities.formatDate(new Date(convenio.fechaEntrega), timeZone, "dd/MM/yyyy") : "";
            const fechaRecepcionStr = convenio.fechaRecibo ? Utilities.formatDate(new Date(convenio.fechaRecibo), timeZone, "dd/MM/yyyy") : "";

            return [
                index + 1, folioCorto, convenio.desarrollo, convenio.condominio,
                loteNumero, fechaEntregaStr, fechaRecepcionStr, "-"
            ];
        });

        if (datosTabla.length > 0) {
            
            const rangoDeTrabajo = tempSheet.getRange("B14:J200");
            rangoDeTrabajo.breakApart();
            rangoDeTrabajo.clearContent().clearFormat();

            const rangoTabla = tempSheet.getRange(tablaInicioFila, 2, datosTabla.length, datosTabla[0].length);
            rangoTabla.setValues(datosTabla);
            rangoTabla.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID)
                    .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true)
                    .setFontWeight("normal").setFontSize(10);        
            const colorFilaAlterna = "#f3f3f3";
            for (let i = 0; i < datosTabla.length; i++) {
              if ((i + 1) % 2 === 0) {
                tempSheet.getRange(tablaInicioFila + i, 2, 1, datosTabla[0].length).setBackground(colorFilaAlterna);
              }
            }
            
            const filaInicioFirmas = tablaInicioFila + datosTabla.length + 3;

            const labelStartCol = 2;
            const labelNumCols = 3;
            const lineStartCol = 5;
            const lineNumCols = 4;

            tempSheet.getRange(filaInicioFirmas, labelStartCol, 1, labelNumCols).merge().setValue("NOMBRE DE QUIEN RECIBE:").setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");
            tempSheet.getRange(filaInicioFirmas, lineStartCol, 1, lineNumCols).merge().setBorder(null, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
            tempSheet.getRange(filaInicioFirmas + 2, labelStartCol, 1, labelNumCols).merge().setValue("PUESTO:").setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");
            tempSheet.getRange(filaInicioFirmas + 2, lineStartCol, 1, lineNumCols).merge().setBorder(null, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
            tempSheet.getRange(filaInicioFirmas + 4, labelStartCol, 1, labelNumCols).merge().setValue("ÁREA:").setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");
            tempSheet.getRange(filaInicioFirmas + 4, lineStartCol, 1, lineNumCols).merge().setBorder(null, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);
            tempSheet.getRange(filaInicioFirmas + 6, labelStartCol, 1, labelNumCols).merge().setValue("FECHA Y FIRMA:").setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");
            tempSheet.getRange(filaInicioFirmas + 6, lineStartCol, 1, lineNumCols).merge().setBorder(null, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID);

            tempSheet.setColumnWidth(2, 20); tempSheet.setColumnWidth(3, 30); tempSheet.setColumnWidth(4, 130);
            tempSheet.setColumnWidth(5, 170); tempSheet.setColumnWidth(6, 30); tempSheet.setColumnWidth(7, 125);
            tempSheet.setColumnWidth(8, 125); tempSheet.setColumnWidth(9, 60);
        }
        
        SpreadsheetApp.flush();
        Utilities.sleep(1500);

        const ssId = ss.getId(), sheetId = tempSheet.getSheetId();
        const url_ext = 'export?exportFormat=pdf&format=pdf&size=letter&portrait=true&fitw=true&sheetnames=false&printtitle=false&gridlines=false&gid=' + sheetId;
        const url = `https://docs.google.com/spreadsheets/d/${ssId}/${url_ext}`;
        
        const options = {
          headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
          muteHttpExceptions: true
        };

        let response;
        const maxRetries = 3;
        for (let i = 0; i < maxRetries; i++) {
          response = UrlFetchApp.fetch(url, options);
          if (response.getResponseCode() === 200) { break; }
          if (i < maxRetries - 1) { Utilities.sleep(2000); }
        }

        if (response.getResponseCode() !== 200) {
          throw new Error(`El servidor de Google falló al intentar generar el PDF (código ${response.getResponseCode()}). Inténtalo de nuevo en un momento.`);
        }

        const nombreArchivo = `Acuse_Convenios_${datos.nombreAnalista.replace(/ /g, "_")}_${hoy.toISOString().slice(0, 10)}.pdf`;
        const pdfBlob = response.getBlob().setName(nombreArchivo);

        const asunto = `Acuse de Entrega de Convenios - ${datos.nombreAnalista} - ${Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd/MM/yyyy")}`;
        const listaFoliosHtml = '<ul>' + datos.convenios.map(c => `<li>${c.folio}</li>`).join('') + '</ul>';
        const cuerpoHtml = `<p>Hola, se adjunta el acuse en formato PDF con el resumen de los <b>${datos.convenios.length} convenios</b> registrados por <b>${datos.nombreAnalista}</b>. Los folios incluidos son:</p>${listaFoliosHtml}<p><i>Este es un correo generado automáticamente.</i></p>`;
        MailApp.sendEmail({to: destinatarioEmail, cc: CORREOS_SUPERVISION_CC.join(','), subject: asunto, htmlBody: cuerpoHtml, attachments: [pdfBlob]});

    } finally {
        plantillaSheet.hideSheet();
        if (tempSheet) {
          ss.deleteSheet(tempSheet);
        }
    }
}

function actualizarConvenios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActivos = ss.getSheetByName('Convenios Activos');
  if (!hojaActivos) return { resumen: 'Error: No se encontró la hoja "Convenios Activos".', incompletos: [], completados: [] };

  const datosActivos = hojaActivos.getDataRange().getValues();
  if (datosActivos.length < 2) return { resumen: 'No se encontraron convenios para mover.', incompletos: [], completados: [] };
  
  const encabezados = datosActivos[0];
  const idx = {
    vencimiento: encabezados.indexOf('Vencimiento'),
    folio: encabezados.indexOf('Folio'),
    propietario: encabezados.indexOf('Nombre Propietario'),
    lote: encabezados.indexOf('Lote Cartera'),
    cierre: encabezados.indexOf('Cierre Convenio'),
    morosidadCartera: encabezados.indexOf('Morosidad (Cartera)'),
    fechaCierre: encabezados.indexOf('Fecha Cierre') // Busca la columna de fecha
  };

  if (Object.values(idx).some(i => i === -1)) {
    const columnaFaltante = Object.keys(idx).find(key => idx[key] === -1);
    return { resumen: `Error: No se encontró la columna "${columnaFaltante}" en "Convenios Activos".`, incompletos: [], completados: [] };
  }

  const filasAMover = [];
  const incompletosParaReporte = [];
  const completadosParaReporte = [];
  const indicesFilasAEliminar = [];
  const fechaDeHoy = new Date(); // Definimos la fecha aquí, fuera del bucle

  for (let i = 1; i < datosActivos.length; i++) {
    const fila = datosActivos[i];
    let debeMoverse = false;
    let comentarioCierre = "";

    const diasParaVencer = fila[idx.vencimiento];
    const morosidad = fila[idx.morosidadCartera];

    if (typeof diasParaVencer === 'number' && diasParaVencer < 0) {
      debeMoverse = true;
      comentarioCierre = "Incompleto";
      incompletosParaReporte.push([ fila[idx.folio], fila[idx.propietario], fila[idx.lote] ]);
    } 
    else if (morosidad === 0 || morosidad === '-' || morosidad === 1) {
      debeMoverse = true;
      comentarioCierre = "Completado";
      completadosParaReporte.push([ fila[idx.folio], fila[idx.propietario], fila[idx.lote] ]);
    }

    if (debeMoverse) {
      let filaModificada = [...fila];
      filaModificada[idx.cierre] = comentarioCierre;
      
      // --- ESTA ES LA LÍNEA QUE FALTABA O ESTABA INCORRECTA ---
      // Nos aseguramos de que el índice de 'fechaCierre' exista y asignamos la fecha.
      if(idx.fechaCierre !== -1) {
        filaModificada[idx.fechaCierre] = fechaDeHoy;
      }
      // --- FIN DE LA CORRECCIÓN ---

      const detalle = `Convenio movido a Vencidos. Causa: ${comentarioCierre}.`;
      registrarEvento(fila[idx.lote], 'CIERRE DE CONVENIO', detalle, fila[idx.folio]);
      
      filasAMover.push(filaModificada);
      indicesFilasAEliminar.push(i + 1);
    }
  }

  if (filasAMover.length === 0) {
    return { resumen: 'No se encontraron convenios vencidos o completados para mover.', incompletos: [], completados: [] };
  }

  let hojaVencidos = ss.getSheetByName('Convenios Vencidos');
  if (!hojaVencidos) {
    hojaVencidos = ss.insertSheet('Convenios Vencidos');
    hojaVencidos.getRange(1, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight('bold');
  }
  hojaVencidos.getRange(_encontrarProximaFilaVacia(hojaVencidos), 1, filasAMover.length, filasAMover[0].length).setValues(filasAMover);
  
  indicesFilasAEliminar.reverse().forEach(rowIndex => {
    hojaActivos.deleteRow(rowIndex);
  });
  
  return {
    resumen: `${filasAMover.length} convenios fueron movidos a "Convenios Vencidos".`,
    incompletos: incompletosParaReporte,
    completados: completadosParaReporte
  };
}

function actualizarFechaCierreConcentrado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaVencidos = ss.getSheetByName('Convenios Vencidos');
  const hojaNegociaciones = ss.getSheetByName('Negociacion-PP'); // <-- NUEVO: Leemos la hoja de negociaciones
  const hojaConcentrado = ss.getSheetByName('Concentrado');

  if (!hojaConcentrado) return;
  
  const mapaFechasCierre = new Map();

  // Función interna para poblar el mapa desde cualquier hoja
  const poblarMapa = (hoja, nombreColumnaCierre) => {
    if (!hoja) return;
    const datos = hoja.getDataRange().getValues();
    const encabezados = datos.shift();
    const idxFolio = encabezados.indexOf('Folio');
    const idxFechaCierre = encabezados.indexOf(nombreColumnaCierre);
    if (idxFolio === -1 || idxFechaCierre === -1) return;
    
    datos.forEach(fila => {
      const folio = fila[idxFolio];
      const fechaCierre = fila[idxFechaCierre];
      if (folio && fechaCierre) {
        mapaFechasCierre.set(folio, fechaCierre);
      }
    });
  };

  // 1. Llenamos el mapa con datos de AMBAS hojas
  poblarMapa(hojaVencidos, 'Fecha Cierre');
  poblarMapa(hojaNegociaciones, 'Fecha Cierre'); // <-- NUEVO

  if (mapaFechasCierre.size === 0) return;

  // 2. Recorremos el Concentrado para actualizar
  const datosConcentrado = hojaConcentrado.getDataRange().getValues();
  const encabezadosConcentrado = datosConcentrado[0];
  const idxFolioConcentrado = encabezadosConcentrado.indexOf('Folio');
  const idxFechaCierreConcentrado = encabezadosConcentrado.indexOf('Fecha Cierre');
  
  if (idxFolioConcentrado === -1 || idxFechaCierreConcentrado === -1) return;

  const nuevosValoresFecha = datosConcentrado.map((fila, index) => {
    if (index === 0) return [encabezadosConcentrado[idxFechaCierreConcentrado]];
    
    const folioActual = fila[idxFolioConcentrado];
    if (mapaFechasCierre.has(folioActual)) {
      return [mapaFechasCierre.get(folioActual)];
    } else {
      return [fila[idxFechaCierreConcentrado]];
    }
  });
  
  // 3. Escribimos toda la columna de una sola vez
  hojaConcentrado.getRange(1, idxFechaCierreConcentrado + 1, nuevosValoresFecha.length, 1).setValues(nuevosValoresFecha);
  Logger.log('Fechas de cierre sincronizadas en el Concentrado.');
}

function actualizacionNegociacionesVencidas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaNegociaciones = ss.getSheetByName('Negociacion-PP');
  if (!hojaNegociaciones) return { resumen: 'Error: No se encontró la hoja "Negociacion-PP".', incompletos: [], completados: [] };

  const lastRow = hojaNegociaciones.getRange("A:A").getValues().filter(String).length;
  if (lastRow < 2) return { resumen: 'No se encontraron negociaciones para procesar.', incompletos: [], completados: [] };
  
  const datos = hojaNegociaciones.getRange(1, 1, lastRow, hojaNegociaciones.getLastColumn()).getValues();
  const encabezados = datos[0];
  const idx = {
    vencimiento: encabezados.indexOf('Vencimiento'),
    morosidadCartera: encabezados.indexOf('Morosidad (Cartera)'),
    cierre: encabezados.indexOf('Cierre NN-PP'),
    fechaCierre: encabezados.indexOf('Fecha Cierre'),
    folio: encabezados.indexOf('Folio'),
    propietario: encabezados.indexOf('Nombre Propietario'),
    lote: encabezados.indexOf('Lote Cartera')
  };

  if (Object.values(idx).some(i => i === -1)) {
    const columnaFaltante = Object.keys(idx).find(key => idx[key] === -1);
    return { resumen: `Error en "Negociacion-PP": No se encontró la columna requerida "${columnaFaltante}".`, incompletos: [], completados: [] };
  }

  const incompletosParaReporte = [];
  const completadosParaReporte = [];
  let filasActualizadas = 0;
  const fechaDeHoy = new Date();
  
  let columnaCierreActualizada = datos.map(fila => [fila[idx.cierre]]);
  let columnaFechaCierreActualizada = datos.map(fila => [fila[idx.fechaCierre]]);

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (fila[idx.cierre] === '') {
      let nuevoEstatus = '';
      const diasParaVencer = fila[idx.vencimiento];
      const morosidad = fila[idx.morosidadCartera];

      if (typeof diasParaVencer === 'number' && diasParaVencer < 0) {
        nuevoEstatus = 'Incompleto';
        incompletosParaReporte.push([fila[idx.folio], fila[idx.propietario], fila[idx.lote]]);
      } 
      // --- INICIO DE LA CORRECCIÓN CLAVE ---
      // Se amplía la condición para que sea idéntica a la de convenios
      else if (morosidad === 0 || morosidad === '-' || morosidad === 1) {
      // --- FIN DE LA CORRECCIÓN CLAVE ---
        nuevoEstatus = 'Completado';
        completadosParaReporte.push([fila[idx.folio], fila[idx.propietario], fila[idx.lote]]);
      }
      
      if (nuevoEstatus !== '') {
        columnaCierreActualizada[i][0] = nuevoEstatus;
        columnaFechaCierreActualizada[i][0] = fechaDeHoy;
        filasActualizadas++;
        
        // También registramos el evento para la "Hoja de Vida"
        const detalle = `El estatus de la negociación cambió a: ${nuevoEstatus}.`;
        registrarEvento(fila[idx.lote], 'CIERRE DE NEGOCIACIÓN', detalle, fila[idx.folio]);
      }
    }
  }

  if (filasActualizadas > 0) {
    hojaNegociaciones.getRange(1, idx.cierre + 1, columnaCierreActualizada.length, 1).setValues(columnaCierreActualizada);
    hojaNegociaciones.getRange(1, idx.fechaCierre + 1, columnaFechaCierreActualizada.length, 1).setValues(columnaFechaCierreActualizada);
  }

  return {
    resumen: `Se actualizaron ${filasActualizadas} estatus en la hoja "Negociacion-PP".`,
    incompletos: incompletosParaReporte,
    completados: completadosParaReporte
  };
}

function enviarRecordatoriosDeVencimiento() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActivos = ss.getSheetByName('Convenios Activos');
  if (!hojaActivos) return { resumen: 'Error: No se encontró la hoja "Convenios Activos".', datos: [] };

  const datos = hojaActivos.getDataRange().getValues();
  const encabezados = datos.shift();
  
  const indices = { 
    vencimiento: encabezados.indexOf('Vencimiento'), 
    ejecutivo: encabezados.indexOf('Correo ejecutivo'), 
    supervisor: encabezados.indexOf('Correo supervisor'), 
    coordinador: encabezados.indexOf('Correo coordinador'), 
    lote: encabezados.indexOf('Lote Cartera'), 
    folio: encabezados.indexOf('Folio'), 
    propietario: encabezados.indexOf('Nombre Propietario'), 
    recordatorio: encabezados.indexOf('Recordatorio Enviado'),
    fechaFin: encabezados.indexOf('Fecha Fin'),
    ejecutivoAsignado: encabezados.indexOf('Ejecutivo Asignado') // <-- NUEVA LÍNEA
  };

  if (Object.values(indices).includes(-1)) {
      const columnaFaltante = Object.keys(indices).find(key => indices[key] === -1);
      return { resumen: `Error: No se encontró la columna requerida "${columnaFaltante}".`, datos: [] };
  }

  const datosParaReporte = [];
  const fechaActual = new Date();
  
  datos.forEach((fila, index) => {
    const diasParaVencer = fila[indices.vencimiento];
    if (typeof diasParaVencer === 'number' && diasParaVencer >= 0 && diasParaVencer <= 7 && !fila[indices.recordatorio]) {
      const correoEjecutivo = fila[indices.ejecutivo];
      if (!correoEjecutivo) return;
      
      let listaCC = [fila[indices.supervisor], fila[indices.coordinador]].concat(CORREOS_ADICIONALES_VENCIMIENTO_CC);
      listaCC = listaCC.filter(correo => correo && correo.toString().trim() !== '').join(',');

      // --- SE OBTIENE EL NOMBRE DEL EJECUTIVO ---
      const nombreEjecutivo = fila[indices.ejecutivoAsignado] || "ejecutivo(a)"; // <-- NUEVA LÍNEA

      const loteCartera = fila[indices.lote];
      const folioConvenio = fila[indices.folio];
      const nombrePropietario = fila[indices.propietario];
      const fechaFin = fila[indices.fechaFin];
      const fechaFinFormateada = (fechaFin instanceof Date) ? Utilities.formatDate(fechaFin, Session.getScriptTimeZone(), "dd/MM/yyyy") : (fechaFin || '');

      const asunto = `Recordatorio de Vencimiento de Convenio: ${loteCartera}`;
      
      // --- SE MODIFICA EL CUERPO DEL CORREO ---
      const cuerpoHtml = `
        <p>Estimado(a) ${nombreEjecutivo},</p> <!-- <-- LÍNEA MODIFICADA -->
        <p>Este es un recordatorio para informarle que el siguiente convenio está próximo a vencer:</p>
        <ul>
          <li><b>Folio:</b> ${folioConvenio}</li>
          <li><b>Propietario:</b> ${nombrePropietario}</li>
          <li><b>Lote Cartera:</b> ${loteCartera}</li>
          <li><b>Fecha de Fin de Convenio:</b> ${fechaFinFormateada}</li>
          <li><b>Días para vencer:</b> ${Math.round(diasParaVencer)}</li>
        </ul>
        <p>Le solicitamos amablemente dar el seguimiento correspondiente para asegurar la renovación o el cierre adecuado del mismo.</p>
        <p>Saludos cordiales.</p>
        <br>
        <p><i><br>Sistema de Notificaciones GPH.</i></p>
      `;
      
      MailApp.sendEmail({ to: correoEjecutivo, cc: listaCC, subject: asunto, htmlBody: cuerpoHtml });
      // <-- AÑADIR REGISTRO DE EVENTO AQUÍ
      registrarEvento(loteCartera, 'RECORDATORIO ENVIADO', `Se envió recordatorio de vencimiento a ${correoEjecutivo}.`, folioConvenio);
      hojaActivos.getRange(index + 2, indices.recordatorio + 1).setValue(fechaActual);
      
      datosParaReporte.push([ folioConvenio, nombrePropietario, loteCartera, fechaFinFormateada, Math.round(diasParaVencer) ]);
    }
  });

  SpreadsheetApp.flush(); 

  return {
    resumen: `Se enviaron ${datosParaReporte.length > 0 ? datosParaReporte.length : '0'} nuevos recordatorios de vencimiento.`,
    datos: datosParaReporte
  };
}

function procesoActualizacionIntegral() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaMaestra = ss.getSheetByName(HOJA_MAESTRA_LOCAL);
    if (!hojaMaestra) throw new Error(`No se encontró la hoja maestra "${HOJA_MAESTRA_LOCAL}".`);
    
    // --- PASO 1: Construir el mapa de cartera UNA SOLA VEZ ---
    const datosMaestros = hojaMaestra.getDataRange().getValues();
    const headersMaestros = datosMaestros[0];
    const idxReferenciaMaestra = 4; // Asumiendo que 'Referencia' es la columna E
    const idxMorosidadMaestra = 5;
    const idxAsesorMaestra = 8;
    const idxEstatusMaestra = headersMaestros.indexOf('ESTATUS');
    if (idxEstatusMaestra === -1) throw new Error(`No se encontró la columna "ESTATUS" en "${HOJA_MAESTRA_LOCAL}".`);

    const carteraMap = new Map();
    for (let i = 1; i < datosMaestros.length; i++) {
      const fila = datosMaestros[i];
      const referencia = fila[idxReferenciaMaestra];
      if (referencia) {
        carteraMap.set(referencia, {
          estatus: fila[idxEstatusMaestra],
          asesor: fila[idxAsesorMaestra],
          morosidad: fila[idxMorosidadMaestra]
        });
      }
    }
    
    // --- PASO 2: Ejecutar la actualización en cada hoja ---
    const resumenActivos = _actualizarDatosDeCarteraEnHoja('Convenios Activos', carteraMap);
    const resumenNegociaciones = _actualizarDatosDeCarteraEnHoja('Negociacion-PP', carteraMap);
    
    // --- PASO 3: Devolver un resumen combinado ---
    // Se eliminó la alerta. El resumen se envía directamente a la función del reporte diario.
    return `${resumenActivos} \n${resumenNegociaciones}`;

  } catch (e) {
    // Se eliminó la alerta de error. El error se "lanza" para que la función del reporte diario lo capture y lo muestre en el correo.
    throw new Error(e.message);
  }
} 

function _actualizarDatosDeCarteraEnHoja(nombreHoja, carteraMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) return `No se procesó "${nombreHoja}" porque no fue encontrada.`;
  
  // --- ESTA ES LA LÍNEA CORREGIDA ---
  // Contamos las filas que realmente tienen un valor en la columna A (Folio), ignorando las vacías.
  const lastRow = hoja.getRange("A:A").getValues().filter(String).length;

  if (lastRow < 2) { // Si solo hay encabezado (o está vacía), no hay nada que procesar.
    return `No se encontraron datos para procesar en "${nombreHoja}".`;
  }
  
  const datosHoja = hoja.getRange(1, 1, lastRow, hoja.getLastColumn()).getValues();
  const headers = datosHoja[0];

  const idxReferencia = headers.indexOf('Referencia');
  const idxEstatusDestino = headers.indexOf('Estatus (Cartera)');
  const idxAsesorDestino = headers.indexOf('Asesor (Cartera)');
  const idxMorosidadDestino = headers.indexOf('Morosidad (Cartera)');

  if ([idxReferencia, idxEstatusDestino, idxAsesorDestino, idxMorosidadDestino].includes(-1)) {
    return `No se procesó "${nombreHoja}" porque le falta una de las columnas requeridas (Referencia, Estatus, Asesor o Morosidad de Cartera).`;
  }

  const estatusResultados = [];
  const asesorResultados = [];
  const morosidadResultados = [];
  let actualizados = 0;

  // El bucle ahora solo recorrerá las filas con datos reales.
  for (let i = 1; i < datosHoja.length; i++) {
    const referencia = datosHoja[i][idxReferencia];
    const datosDeCartera = carteraMap.get(referencia);

    if (datosDeCartera) {
      estatusResultados.push([datosDeCartera.estatus]);
      asesorResultados.push([datosDeCartera.asesor]);
      morosidadResultados.push([datosDeCartera.morosidad]);
      actualizados++;
    } else {
      estatusResultados.push(['No encontrado']);
      asesorResultados.push(['No encontrado']);
      morosidadResultados.push(['No encontrado']);
    }
  }

  if (estatusResultados.length > 0) {
    hoja.getRange(2, idxEstatusDestino + 1, estatusResultados.length, 1).setValues(estatusResultados);
    hoja.getRange(2, idxAsesorDestino + 1, asesorResultados.length, 1).setValues(asesorResultados);
    hoja.getRange(2, idxMorosidadDestino + 1, morosidadResultados.length, 1).setValues(morosidadResultados);
  }
  
  return `En "${nombreHoja}", se actualizaron los datos de ${actualizados} de ${lastRow - 1} registros.`;
}

function ejecucionDiariaProgramada() {
  let reporteHtml = "<h2>Resumen de Tareas Automáticas Ejecutadas</h2><p>A continuación se detalla el resultado de las tareas programadas para el día de hoy:</p>";
  
  // --- TAREAS EN SEGUNDO PLANO (SE EJECUTAN PERO NO SE REPORTAN SI TIENEN ÉXITO) ---
  try {
    sincronizarDatosCompletos();
  } catch (e) {
    reporteHtml += `<p style='color:red;'>❌ <b>¡FALLO CRÍTICO EN SEGUNDO PLANO!</b> Error en Sincronización: ${e.message}</p>`;
  }
  
  try {
    procesoActualizacionIntegral();
  } catch (e) {
    reporteHtml += `<p style='color:red;'>❌ <b>¡FALLO CRÍTICO EN SEGUNDO PLANO!</b> Error en Actualización Integral: ${e.message}</p>`;
  }

  // Se ejecuta la actualización de convenios ANTES de actualizar el concentrado
  let resultadoVencidos;
  try {
      resultadoVencidos = actualizarConvenios();
  } catch(e) {
      reporteHtml += `<p style='color:red;'>❌ Falló la Tarea de Movimiento de Convenios: ${e.message}</p>`;
  }
  
  // Se ejecuta la actualización de negociaciones ANTES de actualizar el concentrado
  let resultadoNegociaciones;
  try {
      resultadoNegociaciones = actualizacionNegociacionesVencidas();
  } catch (e) {
      reporteHtml += `<p style='color:red;'>❌ Falló la Tarea de Actualización de Negociaciones: ${e.message}</p>`;
  }

  // --- NUEVO PASO SILENCIOSO AÑADIDO ---
  // Se ejecuta después de mover los convenios para sincronizar la fecha de cierre.
  try {
    actualizarFechaCierreConcentrado();
  } catch (e) {
    reporteHtml += `<p style='color:red;'>❌ <b>¡FALLO CRÍTICO EN SEGUNDO PLANO!</b> Error al actualizar Fecha Cierre en Concentrado: ${e.message}</p>`;
  }
  
  // Esta tarea de fondo se ejecuta después de que los estatus hayan cambiado
  try {
    actualizarEstatusGeneralConcentrado();
  } catch (e) {
    reporteHtml += `<p style='color:red;'>❌ <b>¡FALLO CRÍTICO EN SEGUNDO PLANO!</b> Error al actualizar Estatus en Concentrado: ${e.message}</p>`;
  }


  // --- INICIO DEL REPORTE VISIBLE PARA EL USUARIO ---

  // TAREA 1 (Visible): Movimiento de Convenios Vencidos
  if (resultadoVencidos) {
    reporteHtml += "<h3>1. Movimiento de Convenios Vencidos</h3>";
    reporteHtml += `<p style='color:green;'>✅ ${resultadoVencidos.resumen}</p>`;
    const tieneIncompletosConv = resultadoVencidos.incompletos && resultadoVencidos.incompletos.length > 0;
    const tieneCompletadosConv = resultadoVencidos.completados && resultadoVencidos.completados.length > 0;
    if (tieneIncompletosConv || tieneCompletadosConv) {
      reporteHtml += "<h4>Cambiar en cartera los estatus de los siguientes casos:</h4>";
      if (tieneIncompletosConv) {
        reporteHtml += "<h5>Convenios Incompletos (Vencidos):</h5>";
        reporteHtml += crearTablaHTML(['Folio', 'Propietario', 'Lote Cartera'], resultadoVencidos.incompletos);
      }
      if (tieneCompletadosConv) {
        reporteHtml += "<h5>Convenios Completados (Liquidados):</h5>";
        reporteHtml += crearTablaHTML(['Folio', 'Propietario', 'Lote Cartera'], resultadoVencidos.completados);
      }
    }
  }

  // TAREA 2 (Visible): Actualización de Estatus en Negociaciones-PP
  if (resultadoNegociaciones) {
    reporteHtml += "<h3>2. Actualización de Estatus en Negociaciones-PP</h3>";
    reporteHtml += `<p style='color:green;'>✅ ${resultadoNegociaciones.resumen}</p>`;
    const tieneIncompletosNeg = resultadoNegociaciones.incompletos && resultadoNegociaciones.incompletos.length > 0;
    const tieneCompletadosNeg = resultadoNegociaciones.completados && resultadoNegociaciones.completados.length > 0;
    if (tieneIncompletosNeg || tieneCompletadosNeg) {
      reporteHtml += "<h4>Cambiar el estatus de las siguientes negociaciones-PP:</h4>";
      if (tieneIncompletosNeg) {
        reporteHtml += "<h5>Negociaciones Incompletas (Vencidas):</h5>";
        reporteHtml += crearTablaHTML(['Folio', 'Propietario', 'Lote Cartera'], resultadoNegociaciones.incompletos);
      }
      if (tieneCompletadosNeg) {
        reporteHtml += "<h5>Negociaciones Completadas (Liquidadas):</h5>";
        reporteHtml += crearTablaHTML(['Folio', 'Propietario', 'Lote Cartera'], resultadoNegociaciones.completados);
      }
    }
  }

  // TAREA 3 (Visible): Envío de Recordatorios de Vencimiento de Convenios
  reporteHtml += "<h3>3. Envío de Recordatorios de Vencimiento de Convenios</h3>";
  try {
    const resultadoRecordatorios = enviarRecordatoriosDeVencimiento();
    reporteHtml += `<p style='color:green;'>✅ ${resultadoRecordatorios.resumen}</p>`;
    if (resultadoRecordatorios.datos && resultadoRecordatorios.datos.length > 0) {
      reporteHtml += "<h4>Detalle de Recordatorios de Convenios Enviados:</h4>";
      const encabezadosRecordatorios = ['Folio', 'Propietario', 'Lote Cartera', 'Fecha Fin', 'Días para Vencer'];
      reporteHtml += crearTablaHTML(encabezadosRecordatorios, resultadoRecordatorios.datos);
    }
  } catch (e) {
    reporteHtml += `<p style='color:red;'>❌ Falló. Error: ${e.message}</p>`;
  }
  
  // TAREA 4 (Visible): Envío de Recordatorios de Negociaciones-PP
  reporteHtml += "<h3>4. Envío de Recordatorios de Negociaciones-PP</h3>";
  try {
    const resultadoRecordatoriosPP = recordatorioNegociacionesPP();
    reporteHtml += `<p style='color:green;'>✅ ${resultadoRecordatoriosPP.resumen}</p>`;
    if (resultadoRecordatoriosPP.datos && resultadoRecordatoriosPP.datos.length > 0) {
      reporteHtml += "<h4>Detalle de Recordatorios de Negociaciones Enviados:</h4>";
      const encabezados = ['Folio', 'Propietario', 'Lote Cartera', 'Fecha Fin', 'Días para Vencer'];
      reporteHtml += crearTablaHTML(encabezados, resultadoRecordatoriosPP.datos);
    }
  } catch (e) {
    reporteHtml += `<p style='color:red;'>❌ Falló. Error: ${e.message}</p>`;
  }

  // TAREA 5 (Visible): Verificación de Estatus en Cartera
  reporteHtml += "<h3>5. Convenios sin estatus correcto en Cartera</h3>";
  try {
    const resultadoEstatus = verificarEstatusConvenioEnCartera();
    reporteHtml += `<p style='color:green;'>✅ ${resultadoEstatus.resumen}</p>`;
    if (resultadoEstatus.datos && resultadoEstatus.datos.length > 0) {
      reporteHtml += "<h4>Desglose de casos que no tienen 'Convenio' como estatus:</h4>";
      const encabezados = ['Folio', 'Propietario', 'Lote Cartera', 'Estatus Encontrado'];
      reporteHtml += crearTablaHTML(encabezados, resultadoEstatus.datos);
    }
  } catch (e) {
    reporteHtml += `<p style='color:red;'>❌ Falló la verificación de estatus. Error: ${e.message}</p>`;
  }

   // TAREA 6 (Visible): Verificación de asesor asignado en cartera
  reporteHtml += "<h3>6. Convenios con Cambio de Asesor en Cartera</h3>";
  try {
    const resultadoCambioAsesor = verificarCambioDeAsesor();
    reporteHtml += `<p style='color:green;'>✅ ${resultadoCambioAsesor.resumen}</p>`;
    if (resultadoCambioAsesor.datos && resultadoCambioAsesor.datos.length > 0) {
      reporteHtml += "<h4>Desglose de convenios que cambiaron de ejecutivo:</h4>";
      const encabezados = ['Folio', 'Lote Cartera', 'Propietario', 'Asignación Original', 'Nueva Asignación (Cartera)'];
      reporteHtml += crearTablaHTML(encabezados, resultadoCambioAsesor.datos);
    }
  } catch (e) {
    reporteHtml += `<p style='color:red;'>❌ Falló la verificación de cambio de asesor. Error: ${e.message}</p>`;
  }

  // TAREA 7 (Visible): Verificación de Estatus de Vencidos en Cartera
  reporteHtml += "<h3>7. Vencidos que aún figuran con Estatus Convenio en Cartera</h3>";
  try {
    const resultadoVencidosEnCartera = verificarEstatusDeVencidosEnCartera();
    reporteHtml += `<p style='color:green;'>✅ ${resultadoVencidosEnCartera.resumen}</p>`;
    if (resultadoVencidosEnCartera.datos && resultadoVencidosEnCartera.datos.length > 0) {
      reporteHtml += "<h4>Se recomienda retirar el estatus 'Convenio' a los siguientes casos en cartera:</h4>";
      const encabezados = ['Folio', 'Propietario', 'Lote Cartera'];
      reporteHtml += crearTablaHTML(encabezados, resultadoVencidosEnCartera.datos);
    }
  } catch (e) {
    reporteHtml += `<p style='color:red;'>❌ Falló la verificación de estatus de vencidos. Error: ${e.message}</p>`;
  }

  // --- Envío del Correo Final ---
  reporteHtml += "<br><p><i>Este es un reporte generado automáticamente.</i></p>";
  const asunto = `Reporte Detallado de Gestión de Convenios - ${new Date().toLocaleDateString()}`;

  MailApp.sendEmail({
    to: CORREO_REPORTE_DIARIO.join(','),
    subject: asunto,
    htmlBody: reporteHtml
  });
}

function crearTablaHTML(encabezados, datos) {
  if (!datos || datos.length === 0) {
    return "";
  }
  
  let tabla = '<table style="border-collapse: collapse; width: 100%; border: 1px solid #ddd; font-family: sans-serif; font-size: 12px;">';
  
  // Encabezados
  tabla += '<thead><tr style="background-color: #f2f2f2;">';
  encabezados.forEach(header => {
    tabla += `<th style="border: 1px solid #ddd; padding: 8px; text-align: left;">${header}</th>`;
  });
  tabla += '</tr></thead>';
  
  // Cuerpo de la tabla
  tabla += '<tbody>';
  datos.forEach(fila => {
    tabla += '<tr>';
    fila.forEach(celda => {
      tabla += `<td style="border: 1px solid #ddd; padding: 8px;">${celda || ''}</td>`;
    });
    tabla += '</tr>';
  });
  tabla += '</tbody></table>';
  
  return tabla;
}

function recordatorioNegociacionesPP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaNegociaciones = ss.getSheetByName('Negociacion-PP');
  if (!hojaNegociaciones) {
    return { resumen: 'Error: No se encontró la hoja "Negociacion-PP".', datos: [] };
  }

  const datos = hojaNegociaciones.getDataRange().getValues();
  if (datos.length < 2) {
      return { resumen: 'No hay datos para procesar en "Negociacion-PP".', datos: [] };
  }
  const encabezados = datos.shift();
  
  const indices = { 
    vencimiento: encabezados.indexOf('Vencimiento'), 
    correoAnalista: encabezados.indexOf('Correo Analista'),
    lote: encabezados.indexOf('Lote Cartera'), 
    folio: encabezados.indexOf('Folio'), 
    propietario: encabezados.indexOf('Nombre Propietario'), 
    recordatorio: encabezados.indexOf('Recordatorio'),
    fechaFin: encabezados.indexOf('Fecha Fin')
  };

  if (Object.values(indices).includes(-1)) {
      const columnaFaltante = Object.keys(indices).find(key => indices[key] === -1);
      return { resumen: `Error en "Negociacion-PP": No se encontró la columna requerida "${columnaFaltante}".`, datos: [] };
  }

  const datosParaReporte = [];
  const fechaActual = new Date();
  
  datos.forEach((fila, index) => {
    const diasParaVencer = fila[indices.vencimiento];
    
    if (typeof diasParaVencer === 'number' && diasParaVencer >= 0 && diasParaVencer <= 3 && !fila[indices.recordatorio]) {
      const correoDestinatario = fila[indices.correoAnalista];
      if (!correoDestinatario) return;
      
      const listaCC = CORREOS_ADICIONALES_NEGOCIACION_CC.filter(correo => correo && correo.toString().trim() !== '').join(',');
      const loteCartera = fila[indices.lote];
      const folio = fila[indices.folio];
      const nombrePropietario = fila[indices.propietario];
      const fechaFin = fila[indices.fechaFin];
      const fechaFinFormateada = (fila[indices.fechaFin] instanceof Date) ? Utilities.formatDate(fila[indices.fechaFin], Session.getScriptTimeZone(), "dd/MM/yyyy") : (fila[indices.fechaFin] || '');

      const asunto = `Recordatorio de Vencimiento de Negociación/PP: ${loteCartera}`;
      
      // --- CUERPO DEL CORREO RESTAURADO ---
      const cuerpoHtml = `
        <p>Hola,</p>
        <p>Este es un recordatorio para informarle que la siguiente <b>negociación o promesa de pago</b> está próxima a vencer:</p>
        <ul>
          <li><b>Folio:</b> ${folio}</li>
          <li><b>Propietario:</b> ${nombrePropietario}</li>
          <li><b>Lote Cartera:</b> ${loteCartera}</li>
          <li><b>Fecha de Fin:</b> ${fechaFinFormateada}</li>
          <li><b>Días para vencer:</b> ${Math.round(diasParaVencer)}</li>
        </ul>
        <p>Le solicitamos amablemente dar el seguimiento correspondiente.</p>
        <p>Saludos cordiales.</p>
        <br>
        <p><i>--<br>Sistema de Notificaciones GPH.</i></p>
      `;
      
      MailApp.sendEmail({ to: correoDestinatario, cc: listaCC, subject: asunto, htmlBody: cuerpoHtml });
      
      registrarEvento(loteCartera, 'RECORDATORIO ENVIADO', `Se envió recordatorio de vencimiento a ${correoDestinatario}.`, folio);
      
      hojaNegociaciones.getRange(index + 2, indices.recordatorio + 1).setValue(fechaActual);
      
      datosParaReporte.push([ folio, nombrePropietario, loteCartera, fechaFinFormateada, Math.round(diasParaVencer) ]);
    }
  });

  SpreadsheetApp.flush(); 

  return {
    resumen: `Se enviaron ${datosParaReporte.length} nuevos recordatorios de Negociaciones/PP.`,
    datos: datosParaReporte
  };
}


// Función para mostrar el diálogo de fechas
function mostrarDialogoReporte() {
  const html = HtmlService.createHtmlOutputFromFile('SelectorFechasReporte').setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Seleccionar Fechas');
}

function generarReporteMensual(fechas) {
  if (!fechas.inicio || !fechas.fin) { throw new Error("Debes seleccionar una fecha de inicio y una de fin."); }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHojaReporte = 'Reporte Mensual';
  let hojaReporte = ss.getSheetByName(nombreHojaReporte);
  if (!hojaReporte) { hojaReporte = ss.insertSheet(nombreHojaReporte); }
  
  const charts = hojaReporte.getCharts();
  charts.forEach(chart => hojaReporte.removeChart(chart));
  hojaReporte.clear();
  hojaReporte.showSheet();
  ss.setActiveSheet(hojaReporte);

  const hojaConcentrado = ss.getSheetByName('Concentrado');
  if (!hojaConcentrado) throw new Error('No se encuentra la hoja "Concentrado".');

  const datos = hojaConcentrado.getDataRange().getValues();
  const encabezados = datos.shift();
  const idx = {
    folio: encabezados.indexOf('Folio'),
    fechaCierre: encabezados.indexOf('Fecha Cierre'),
    estatus: encabezados.indexOf('Estatus General'),
    ejecutivo: encabezados.indexOf('Ejecutivo Asignado'),
    monto: encabezados.indexOf('Monto Adeudo Total')
  };
  
  if (Object.values(idx).some(i => i === -1)) { throw new Error('Faltan columnas clave en Concentrado.'); }

  const fechaInicio = new Date(fechas.inicio);
  const fechaFin = new Date(fechas.fin);
  fechaFin.setHours(23, 59, 59);

  const resultadosPorMes = {};
  const rendimientoPorEjecutivo = {};
  const nombresMeses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

  datos.forEach(fila => {
    const folio = fila[idx.folio], fechaCierre = fila[idx.fechaCierre], estatus = fila[idx.estatus], ejecutivo = fila[idx.ejecutivo] || "Sin Asignar", monto = parseFloat(fila[idx.monto]) || 0;
    if (String(folio).startsWith('CONV-') && fechaCierre instanceof Date && fechaCierre >= fechaInicio && fechaCierre <= fechaFin) {
      const mesKey = `${fechaCierre.getFullYear()}-${fechaCierre.getMonth()}`;
      const mesNombre = `${nombresMeses[fechaCierre.getMonth()]} ${fechaCierre.getFullYear()}`;
      if (!resultadosPorMes[mesKey]) {
        resultadosPorMes[mesKey] = { nombre: mesNombre, exitosos: 0, incumplidos: 0, orden: fechaCierre.getFullYear() * 100 + fechaCierre.getMonth() };
      }
      if (!rendimientoPorEjecutivo[ejecutivo]) {
        rendimientoPorEjecutivo[ejecutivo] = { exitosos: 0, incumplidos: 0, montoExitoso: 0, montoIncumplido: 0 };
      }
      if (estatus === 'Completado') {
        resultadosPorMes[mesKey].exitosos++;
        rendimientoPorEjecutivo[ejecutivo].exitosos++;
        rendimientoPorEjecutivo[ejecutivo].montoExitoso += monto;
      } else if (estatus === 'Incompleto') {
        resultadosPorMes[mesKey].incumplidos++;
        rendimientoPorEjecutivo[ejecutivo].incumplidos++;
        rendimientoPorEjecutivo[ejecutivo].montoIncumplido += monto;
      }
    }
  });
  
  hojaReporte.getRange("B2:H2").merge().setValue("Reporte Mensual de Desempeño de Convenios").setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  hojaReporte.getRange("B3:H3").merge().setValue(`Período: ${fechas.inicio} al ${fechas.fin}`).setFontStyle('italic').setHorizontalAlignment('center');

  let filaActual = 5;
  const datosAgrupados = Object.values(resultadosPorMes).sort((a, b) => a.orden - b.orden);
  
  // --- TABLA 1: RESUMEN POR MES (CON TOTALES) ---
  const datosTablaMes = [['Mes', 'Exitosos', 'Incumplidos', 'Total Cierres', '% Exitosos', '% Incumplidos', 'Total %']];
  datosAgrupados.forEach(mes => {
    const total = mes.exitosos + mes.incumplidos;
    datosTablaMes.push([ mes.nombre, mes.exitosos, mes.incumplidos, total, (total > 0 ? mes.exitosos / total : 0), (total > 0 ? mes.incumplidos / total : 0), 1 ]);
  });
  
  let filaInicioTablaMes = 0;
  if (datosTablaMes.length > 1) {
    const totalExitosos = datosAgrupados.reduce((sum, item) => sum + item.exitosos, 0);
    const totalIncumplidos = datosAgrupados.reduce((sum, item) => sum + item.incumplidos, 0);
    const granTotal = totalExitosos + totalIncumplidos;
    datosTablaMes.push(['TOTALES', totalExitosos, totalIncumplidos, granTotal, (granTotal > 0 ? totalExitosos / granTotal : 0), (granTotal > 0 ? totalIncumplidos / granTotal : 0), 1]);

    hojaReporte.getRange(filaActual, 2).setValue("Resumen por Mes").setFontSize(14).setFontWeight('bold');
    filaActual++;
    filaInicioTablaMes = filaActual;
    const rangoTablaMes = hojaReporte.getRange(filaActual, 2, datosTablaMes.length, 7);
    rangoTablaMes.setValues(datosTablaMes);
    rangoTablaMes.setBorder(true, true, true, true, true, true, '#B7B7B7', SpreadsheetApp.BorderStyle.SOLID);
    hojaReporte.getRange(filaActual, 2, 1, 7).setBackground('#4A86E8').setFontColor('white').setFontWeight('bold');
    const filaTotalesMes = filaActual + datosTablaMes.length - 1;
    hojaReporte.getRange(filaTotalesMes, 2, 1, 7).setBackground('#CFE2F3').setFontWeight('bold');
    hojaReporte.getRange(filaActual + 1, 3, datosTablaMes.length - 1, 5).setHorizontalAlignment('center');
    hojaReporte.getRange(filaActual + 1, 6, datosTablaMes.length - 1, 3).setNumberFormat('0.00%');
    hojaReporte.autoResizeColumns(2, 7);
    filaActual += datosTablaMes.length + 2;
  }

  // --- TABLA 2: RENDIMIENTO POR EJECUTIVO (CON TOTALES) ---
  const datosEjecutivos = [['Ejecutivo', 'Exitosos', 'Incumplidos', 'Tasa Éxito', 'Monto Recuperado', 'Monto en Riesgo']];
  const ejecutivosArray = Object.keys(rendimientoPorEjecutivo).map(nombre => ({ nombre, ...rendimientoPorEjecutivo[nombre] }));
  ejecutivosArray.sort((a, b) => {
      const tasaA = (a.exitosos + a.incumplidos > 0) ? a.exitosos / (a.exitosos + a.incumplidos) : 0;
      const tasaB = (b.exitosos + b.incumplidos > 0) ? b.exitosos / (b.exitosos + b.incumplidos) : 0;
      if (tasaB !== tasaA) return tasaB - tasaA;
      return (b.exitosos + b.incumplidos) - (a.exitosos + a.incumplidos);
  });
  ejecutivosArray.forEach(e => {
      const total = e.exitosos + e.incumplidos;
      datosEjecutivos.push([e.nombre, e.exitosos, e.incumplidos, (total > 0 ? e.exitosos / total : 0), e.montoExitoso, e.montoIncumplido]);
  });
  
  let filaInicioTablaEjec = 0;
  if (datosEjecutivos.length > 1) {
      const totalExitososEjec = ejecutivosArray.reduce((sum, e) => sum + e.exitosos, 0);
      const totalIncumplidosEjec = ejecutivosArray.reduce((sum, e) => sum + e.incumplidos, 0);
      const granTotalEjec = totalExitososEjec + totalIncumplidosEjec;
      const efectividadTotalEjec = (granTotalEjec > 0) ? totalExitososEjec / granTotalEjec : 0;
      const totalMontoRecuperado = ejecutivosArray.reduce((sum, e) => sum + e.montoExitoso, 0);
      const totalMontoEnRiesgo = ejecutivosArray.reduce((sum, e) => sum + e.montoIncumplido, 0);
      datosEjecutivos.push(['TOTALES', totalExitososEjec, totalIncumplidosEjec, efectividadTotalEjec, totalMontoRecuperado, totalMontoEnRiesgo]);
      
      hojaReporte.getRange(filaActual, 2).setValue("Rendimiento por Ejecutivo").setFontSize(14).setFontWeight('bold');
      filaActual++;
      filaInicioTablaEjec = filaActual;
      const rangoTablaEjecutivos = hojaReporte.getRange(filaActual, 2, datosEjecutivos.length, 6);
      rangoTablaEjecutivos.setValues(datosEjecutivos);
      rangoTablaEjecutivos.setBorder(true, true, true, true, true, true, '#B7B7B7', SpreadsheetApp.BorderStyle.SOLID);
      hojaReporte.getRange(filaActual, 2, 1, 6).setBackground('#0B5394').setFontColor('white').setFontWeight('bold');
      const filaTotalesEjec = filaActual + datosEjecutivos.length - 1;
      hojaReporte.getRange(filaTotalesEjec, 2, 1, 6).setBackground('#D9EAD3').setFontWeight('bold');
      hojaReporte.getRange(filaActual + 1, 3, datosEjecutivos.length - 1, 2).setHorizontalAlignment('center');
      hojaReporte.getRange(filaActual + 1, 5, datosEjecutivos.length - 1, 1).setNumberFormat('0.00%');
      hojaReporte.getRange(filaActual + 1, 6, datosEjecutivos.length - 1, 2).setNumberFormat('$#,##0.00');
      hojaReporte.autoResizeColumns(2, 6);
      filaActual += datosEjecutivos.length + 2;
  }
  
  const totalGeneralExitosos = datosAgrupados.reduce((sum, e) => sum + e.exitosos, 0);
  const totalGeneralIncumplidos = datosAgrupados.reduce((sum, e) => sum + e.incumplidos, 0);
  
  if (totalGeneralExitosos > 0 || totalGeneralIncumplidos > 0) {
      if (filaInicioTablaMes > 0) {
          const rangoGraficoBarrasMes = hojaReporte.getRange(filaInicioTablaMes + 1, 2, datosTablaMes.length - 2, 3); // -2 para excluir encabezado y totales
          const barChartMes = hojaReporte.newChart().asColumnChart()
              .addRange(rangoGraficoBarrasMes).setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
              .setNumHeaders(0).setOption('title', 'Volumen de Cierres por Mes')
              .setOption('colors', ['#6AA84F', '#CC0000']).setOption('isStacked', 'true')
              .setOption('legend', { position: 'top' }).setOption('series', { 0: { dataLabel: 'value' }, 1: { dataLabel: 'value' } })
              .setPosition(filaActual, 2, 0, 0).build();
          hojaReporte.insertChart(barChartMes);
      }

      const rangoResumenPie = hojaReporte.getRange(2, 10, 3, 2); // J2:K4
      rangoResumenPie.setValues([['Tipo', 'Total'], ['Exitosos', totalGeneralExitosos], ['Incumplidos', totalGeneralIncumplidos]]);
      const pieChart = hojaReporte.newChart().asPieChart()
          .addRange(rangoResumenPie).setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
          .setNumHeaders(1).setOption('title', 'Distribución Total del Período (Volumen)')
          .setOption('colors', ['#6AA84F', '#CC0000']).setOption('pieHole', 0.4)
          .setOption('legend', { position: 'labeled' }).setOption('pieSliceText', 'value')
          .setPosition(filaActual, 6, 0, 0).build();
      hojaReporte.insertChart(pieChart);
      hojaReporte.hideColumns(10, 2); // Oculta J y K

      if (filaInicioTablaEjec > 0) {
          const rangoGraficoBarrasEjec = hojaReporte.getRange(filaInicioTablaEjec + 1, 2, datosEjecutivos.length - 2, 3); // Excluir encabezado y totales
          const barChartEjecutivo = hojaReporte.newChart().asBarChart()
              .addRange(rangoGraficoBarrasEjec).setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
              .setNumHeaders(0).setOption('title', 'Volumen de Cierres por Ejecutivo')
              .setOption('colors', ['#6AA84F', '#CC0000']).setOption('isStacked', 'true')
              .setOption('legend', { position: 'top' })
              .setPosition(filaActual, 16, 0, 0).build();
          hojaReporte.insertChart(barChartEjecutivo);
      }
  } else {
    hojaReporte.getRange(5, 2).setValue("No se encontraron convenios concluidos en el período para generar un reporte.").setFontStyle('italic');
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast('¡Reporte generado con éxito!', 'Proceso Finalizado', 5);
}

function mostrarResultadosEnDialogo(tablaHTML) {
  const htmlTemplate = HtmlService.createTemplateFromFile('ResultadosBusqueda');
  htmlTemplate.tablaHTML = tablaHTML;
  const html = htmlTemplate.evaluate().setWidth(1000).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Resultados de la Búsqueda');
}

function mostrarDialogoBusqueda() {
  const html = HtmlService.createHtmlOutputFromFile('BuscadorAvanzado').setWidth(600).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Buscador Avanzado de Trámites');
}

function obtenerDatosParaBuscador() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaConcentrado = ss.getSheetByName('Concentrado');
  if (!hojaConcentrado) throw new Error('No se encontró la hoja "Concentrado".');

  const datos = hojaConcentrado.getDataRange().getValues();
  const encabezados = datos.shift();
  const idxAnalista = encabezados.indexOf('Registrado por');
  const idxDesarrollo = encabezados.indexOf('Desarrollo');
  const idxEstatus = encabezados.indexOf('Estatus General'); // <-- NUEVA LÍNEA

  const analistasUnicos = [...new Set(datos.map(fila => fila[idxAnalista]).filter(Boolean))].sort();
  const desarrollosUnicos = [...new Set(datos.map(fila => fila[idxDesarrollo]).filter(Boolean))].sort();
  const estatusUnicos = [...new Set(datos.map(fila => fila[idxEstatus]).filter(Boolean))].sort(); // <-- NUEVA LÍNEA

  return {
    analistas: analistasUnicos,
    desarrollos: desarrollosUnicos,
    estatus: estatusUnicos, // <-- NUEVA LÍNEA
    columnas: encabezados
  };
}

function ejecutarBusquedaAvanzada(busqueda) {
  const resultados = _realizarBusqueda(busqueda);
  const encabezados = resultados.shift();
  
  let tablaHtml = '<table class="table table-striped table-bordered table-sm">';
  tablaHtml += '<thead class="thead-dark"><tr>' + encabezados.map(h => `<th>${h}</th>`).join('') + '</tr></thead>';
  tablaHtml += '<tbody>' + resultados.map(fila => '<tr>' + fila.map(val => `<td>${val instanceof Date ? Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy") : (val || "")}</td>`).join('') + '</tr>').join('') + '</tbody></table>';
  
  return tablaHtml;
}

function exportarBusquedaAExcel(busqueda) {
  const resultados = _realizarBusqueda(busqueda);
  
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const nombreArchivo = `Reporte_Busqueda_${busqueda.metodo}_${fecha}`;
  const nuevoSpreadsheet = SpreadsheetApp.create(nombreArchivo);
  
  const hojaResultados = nuevoSpreadsheet.getActiveSheet().setName("Resultados");
  
  hojaResultados.getRange(1, 1, resultados.length, resultados[0].length).setValues(resultados);
  hojaResultados.getRange(1, 1, 1, resultados[0].length).setFontWeight('bold').setBackground('#eeeeee');
  hojaResultados.autoResizeColumns(1, resultados[0].length);
  
  SpreadsheetApp.getUi().alert('¡Éxito!', `Se ha creado el archivo "${nombreArchivo}" en la raíz de tu Google Drive.`, SpreadsheetApp.getUi().ButtonSet.OK);
  
  // Devuelve un mensaje de éxito para que el HTML sepa que puede cerrarse.
  return "Exportación exitosa.";
}

function _realizarBusqueda(busqueda) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaConcentrado = ss.getSheetByName('Concentrado');
  const datos = hojaConcentrado.getDataRange().getValues();
  const encabezadosOriginales = datos.shift();

  const columnaMap = {
    'referencia': 'Referencia',
    'analista': 'Registrado por',
    'desarrollo': 'Desarrollo'
  };
  const columnaBusquedaIndex = encabezadosOriginales.indexOf(columnaMap[busqueda.metodo]);
  const columnaEstatusIndex = encabezadosOriginales.indexOf('Estatus General'); // <-- NUEVA LÍNEA

  if (columnaBusquedaIndex === -1) {
    throw new Error(`La columna de búsqueda "${columnaMap[busqueda.metodo]}" no fue encontrada.`);
  }

  const valoresBusquedaSet = new Set(busqueda.valores.map(v => String(v).toLowerCase()));
  
  // 1. PRIMER FILTRO (OBLIGATORIO)
  let resultadosFiltrados = datos.filter(fila => {
    const valorCelda = (fila[columnaBusquedaIndex] || "").toString().toLowerCase();
    return valoresBusquedaSet.has(valorCelda);
  });

  // 2. SEGUNDO FILTRO (OPCIONAL, POR ESTATUS)
  if (busqueda.estatus && busqueda.estatus.length > 0) {
    const estatusBusquedaSet = new Set(busqueda.estatus);
    resultadosFiltrados = resultadosFiltrados.filter(fila => {
      const valorEstatus = fila[columnaEstatusIndex];
      return estatusBusquedaSet.has(valorEstatus);
    });
  }

  if (resultadosFiltrados.length === 0) {
    throw new Error(`No se encontraron resultados para la combinación de filtros proporcionada.`);
  }

  const indicesColumnasSeleccionadas = busqueda.columnas.map(nombreCol => encabezadosOriginales.indexOf(nombreCol));
  
  const resultadosFinales = [busqueda.columnas]; // Encabezados primero
  resultadosFiltrados.forEach(fila => {
    const filaFormateada = indicesColumnasSeleccionadas.map(index => index > -1 ? fila[index] : "");
    resultadosFinales.push(filaFormateada);
  });
  
  return resultadosFinales;
}

function actualizarEstatusGeneralConcentrado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const HOJA_CONCENTRADO = 'Concentrado';
  const HOJA_ACTIVOS = 'Convenios Activos';
  const HOJA_VENCIDOS = 'Convenios Vencidos';
  const HOJA_NEGOCIACIONES = 'Negociacion-PP';

  const COL_FOLIO = 'Folio';
  const COL_CIERRE_CONVENIO = 'Cierre Convenio';
  const COL_CIERRE_NEGOCIACION = 'Cierre NN-PP';
  const COL_ESTATUS_GENERAL = 'Estatus General';

  try {
    const hojaConcentrado = ss.getSheetByName(HOJA_CONCENTRADO);
    const hojaActivos = ss.getSheetByName(HOJA_ACTIVOS);
    const hojaVencidos = ss.getSheetByName(HOJA_VENCIDOS);
    const hojaNegociaciones = ss.getSheetByName(HOJA_NEGOCIACIONES);

    if (!hojaConcentrado || !hojaActivos || !hojaVencidos || !hojaNegociaciones) {
      throw new Error('No se encontró una de las hojas requeridas. Verifica los nombres.');
    }

    const crearMapaDeEstatus = (hoja, colEstatusNombre) => {
      const mapa = new Map();
      const datos = hoja.getDataRange().getValues();
      const encabezados = datos.shift();
      const idxFolio = encabezados.indexOf(COL_FOLIO);
      const idxEstatus = encabezados.indexOf(colEstatusNombre);
      if (idxFolio === -1 || idxEstatus === -1) return mapa;
      
      datos.forEach(fila => {
        const folio = fila[idxFolio];
        if (folio) {
          mapa.set(folio, fila[idxEstatus] || '');
        }
      });
      return mapa;
    };

    const mapaActivos = crearMapaDeEstatus(hojaActivos, COL_CIERRE_CONVENIO);
    const mapaVencidos = crearMapaDeEstatus(hojaVencidos, COL_CIERRE_CONVENIO);
    const mapaNegociaciones = crearMapaDeEstatus(hojaNegociaciones, COL_CIERRE_NEGOCIACION);
    
    const datosConcentrado = hojaConcentrado.getDataRange().getValues();
    const encabezadosConcentrado = datosConcentrado[0];
    const idxFolioConcentrado = encabezadosConcentrado.indexOf(COL_FOLIO);
    const idxEstatusGeneral = encabezadosConcentrado.indexOf(COL_ESTATUS_GENERAL);
    
    if (idxEstatusGeneral === -1) {
      throw new Error(`No se encontró la columna "${COL_ESTATUS_GENERAL}" en la hoja "${HOJA_CONCENTRADO}". Por favor, añádela.`);
    }

    const nuevosEstatuses = [];
    for (let i = 1; i < datosConcentrado.length; i++) {
      const folio = datosConcentrado[i][idxFolioConcentrado];
      let estatusFinal = '';

      if (mapaActivos.has(folio)) {
        const estatusActivo = mapaActivos.get(folio);
        estatusFinal = (estatusActivo === '') ? 'Abierto' : estatusActivo;
      } else if (mapaVencidos.has(folio)) {
        estatusFinal = mapaVencidos.get(folio);
      } else if (mapaNegociaciones.has(folio)) {
        // --- INICIO DE LA MODIFICACIÓN ---
        const estatusNegociacion = mapaNegociaciones.get(folio);
        estatusFinal = (estatusNegociacion === '') ? 'Abierto' : estatusNegociacion;
        // --- FIN DE LA MODIFICACIÓN ---
      }
      
      nuevosEstatuses.push([estatusFinal]);
    }
    
    if (nuevosEstatuses.length > 0) {
      hojaConcentrado.getRange(2, idxEstatusGeneral + 1, nuevosEstatuses.length, 1).setValues(nuevosEstatuses);
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Estatus en "Concentrado" actualizados con éxito.');
    return 'Proceso completado exitosamente.';

  } catch (e) {
    throw new Error(e.message);
  }
}

function mostrarDialogoFiltro() {
  const html = HtmlService.createHtmlOutputFromFile('FiltroDinamico').setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Asistente de Filtrado');
}

function obtenerColumnasDeHoja(nombreHoja) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (!hoja) throw new Error(`La hoja "${nombreHoja}" no fue encontrada.`);
  return hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].filter(Boolean);
}

function obtenerValoresUnicosDeColumna(nombreHoja, nombreColumna) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (!hoja) throw new Error(`La hoja "${nombreHoja}" no fue encontrada.`);
  
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos.shift();
  const indiceColumna = encabezados.indexOf(nombreColumna);
  if (indiceColumna === -1) throw new Error(`La columna "${nombreColumna}" no fue encontrada.`);
  
  const valoresUnicos = [...new Set(datos.map(fila => fila[indiceColumna]).filter(Boolean))];
  return valoresUnicos.sort();
}

function aplicarFiltroDinamico(filtro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = ss.getSheetByName(filtro.hoja);
  if (!hojaOrigen) throw new Error(`La hoja "${filtro.hoja}" no fue encontrada.`);

  const datos = hojaOrigen.getDataRange().getValues();
  const encabezados = datos.shift();
  const indiceColumna = encabezados.indexOf(filtro.columna);
  if (indiceColumna === -1) throw new Error(`La columna "${filtro.columna}" no fue encontrada.`);

  const valoresSet = new Set(filtro.valores);
  const resultados = datos.filter(fila => valoresSet.has(fila[indiceColumna]));

  if (resultados.length === 0) {
    SpreadsheetApp.getUi().alert('Sin Resultados', 'No se encontraron filas que coincidan con los filtros seleccionados.', SpreadsheetApp.getUi().ButtonSet.OK);
    return "No se encontraron resultados."; // Devuelve un mensaje para que el HTML sepa que terminó.
  }

  // --- INICIO DE LA NUEVA LÓGICA ---

  // 1. Añadir los encabezados de vuelta al principio de los resultados
  resultados.unshift(encabezados);

  // 2. Construir la tabla HTML (similar al buscador avanzado)
  let tablaHtml = '<table class="table table-striped table-bordered table-sm">';
  tablaHtml += '<thead class="thead-dark"><tr>' + resultados[0].map(h => `<th>${h}</th>`).join('') + '</tr></thead>';
  tablaHtml += '<tbody>';
  resultados.slice(1).forEach(fila => {
    tablaHtml += '<tr>';
    fila.forEach(val => {
      const valorFormateado = val instanceof Date ? Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy") : (val || "");
      tablaHtml += `<td>${valorFormateado}</td>`;
    });
    tablaHtml += '</tr>';
  });
  tablaHtml += '</tbody></table>';
  
  // 3. Llamar a la función que muestra el diálogo con la tabla generada
  mostrarResultadosEnDialogo(tablaHtml);

  // 4. Devolver un mensaje de éxito para que el HTML sepa que puede cerrarse.
  return "Filtro aplicado exitosamente.";
  
  // --- FIN DE LA NUEVA LÓGICA ---
}

function mostrarDialogoFiltroRapido() {
  const html = HtmlService.createHtmlOutputFromFile('FiltroRapido').setWidth(500).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Filtro Rápido en Hoja');
}

function obtenerColumnasDeHojaActiva() {
  const hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Solo permite filtrar en las hojas principales
  const hojasPermitidas = ['Concentrado', 'Convenios Activos', 'Negociacion-PP', 'Convenios Vencidos'];
  if (!hojasPermitidas.includes(hojaActiva.getName())) {
    throw new Error('Esta función solo se puede usar en las hojas: Concentrado, Convenios Activos, Negociacion-PP o Convenios Vencidos.');
  }
  return hojaActiva.getRange(1, 1, 1, hojaActiva.getLastColumn()).getValues()[0].filter(Boolean);
}

function obtenerValoresUnicosDeColumnaActiva(nombreColumna) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos.shift();
  const indiceColumna = encabezados.indexOf(nombreColumna);
  if (indiceColumna === -1) throw new Error(`La columna "${nombreColumna}" no fue encontrada.`);
  
  const valoresUnicos = [...new Set(datos.map(fila => fila[indiceColumna]).filter(Boolean))];
  return valoresUnicos.sort();
}

function aplicarFiltroEnHoja(filtro) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rango = hoja.getDataRange();
  
  // Si ya existe un filtro, lo quitamos para empezar de cero.
  if (rango.getFilter()) {
    rango.getFilter().remove();
  }

  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const indiceColumna = encabezados.indexOf(filtro.columna);
  if (indiceColumna === -1) throw new Error(`La columna "${filtro.columna}" no fue encontrada.`);
  
  // Creamos un nuevo filtro
  const filtroSheet = rango.createFilter();
  
  // Construimos el criterio para el filtro
  const criterio = SpreadsheetApp.newFilterCriteria()
    .setHiddenValues(obtenerValoresUnicosDeColumnaActiva(filtro.columna).filter(valor => !filtro.valores.includes(valor)))
    .build();

  filtroSheet.setColumnFilterCriteria(indiceColumna + 1, criterio);
  SpreadsheetApp.getActiveSpreadsheet().toast('Filtro rápido aplicado.');
}

function quitarFiltrosDeHoja() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const filtro = hoja.getFilter();
  if (filtro) {
    filtro.remove();
    SpreadsheetApp.getActiveSpreadsheet().toast('Filtros eliminados.');
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('No hay filtros activos en esta hoja.');
  }
}

// AÑADE ESTA NUEVA FUNCIÓN A TU ARCHIVO .GS DE CONVENIOS

/**
 * Registra un evento en la hoja de vida de un lote.
 * @param {string} loteCartera El identificador único del lote.
 * @param {string} tipoEvento Una categoría para el evento (ej. "NUEVO CONVENIO").
 * @param {string} detalle Una descripción legible de lo que ocurrió.
 * @param {string} [folio=''] El folio del trámite asociado, si aplica.
 */
function registrarEvento(loteCartera, tipoEvento, detalle, folio = '') {
  try {
    const hojaLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log_Eventos');
    if (!hojaLog) return; // Si la hoja no existe, no hace nada.

    const fechaHora = new Date();
    const usuario = Session.getActiveUser().getEmail();

    hojaLog.appendRow([fechaHora, loteCartera, tipoEvento, detalle, usuario, folio]);
  } catch (e) {
    Logger.log(`Error al registrar evento para el lote ${loteCartera}: ${e.message}`);
  }
}

// AÑADE ESTAS DOS NUEVAS FUNCIONES AL FINAL DE TU .GS

/**
 * Muestra un prompt para que el usuario ingrese un Lote Cartera.
 */
function mostrarDialogoHistorial() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Ver Historial de Lote',
    'Por favor, ingrese el "Lote Cartera" completo que desea consultar (ej. CMLJAL-FLAH-033):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    const loteCartera = response.getResponseText().trim().toUpperCase();
    if (loteCartera) {
      buscarYMostrarHistorial(loteCartera);
    }
  }
}

/**
 * Busca todos los eventos de un lote y los muestra en un diálogo.
 * @param {string} loteCartera El lote a buscar.
 */
function buscarYMostrarHistorial(loteCartera) {
  const hojaLog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log_Eventos');
  if (!hojaLog) {
    SpreadsheetApp.getUi().alert('Error', 'No se encontró la hoja "Log_Eventos".', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const datos = hojaLog.getDataRange().getValues();
  const encabezados = datos.shift();
  const idxLote = encabezados.indexOf('Lote Cartera');

  const historial = datos.filter(fila => fila[idxLote] === loteCartera);

  if (historial.length === 0) {
    SpreadsheetApp.getUi().alert('Sin Resultados', `No se encontraron eventos para el lote "${loteCartera}".`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Construimos una tabla HTML para mostrar los resultados
  let tablaHtml = `<p>Mostrando ${historial.length} evento(s) para el lote <b>${loteCartera}</b>.</p>`;
  tablaHtml += '<table class="table table-striped table-sm" style="font-size: 0.9em;">';
  tablaHtml += '<thead class="thead-dark"><tr><th>Fecha y Hora</th><th>Tipo de Evento</th><th>Detalle</th><th>Usuario</th><th>Folio</th></tr></thead>';
  tablaHtml += '<tbody>';
  historial.reverse().forEach(fila => { // .reverse() para mostrar lo más nuevo primero
    const fecha = fila[0] instanceof Date ? Utilities.formatDate(fila[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : fila[0];
    tablaHtml += `<tr><td>${fecha}</td><td>${fila[2]}</td><td>${fila[3]}</td><td>${fila[4]}</td><td>${fila[5]}</td></tr>`;
  });
  tablaHtml += '</tbody></table>';

  const htmlOutput = HtmlService.createHtmlOutput(tablaHtml)
      .setWidth(800)
      .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Hoja de Vida del Lote: ${loteCartera}`);
}

function verificarEstatusConvenioEnCartera() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActivos = ss.getSheetByName('Convenios Activos');
  
  if (!hojaActivos) {
    return { resumen: 'No se pudo verificar la hoja "Convenios Activos" porque no fue encontrada.', datos: [] };
  }
  
  const datos = hojaActivos.getDataRange().getValues();
  const encabezados = datos.shift();
  
  const idx = {
    estatus: encabezados.indexOf('Estatus (Cartera)'),
    folio: encabezados.indexOf('Folio'),
    propietario: encabezados.indexOf('Nombre Propietario'),
    lote: encabezados.indexOf('Lote Cartera')
  };
  
  if (idx.estatus === -1) {
    return { resumen: 'No se pudo verificar porque no se encontró la columna "Estatus (Cartera)".', datos: [] };
  }
  
  const casosInconsistentes = [];
  datos.forEach(fila => {
    // Limpiamos el valor para una comparación robusta (ignorando mayúsculas/minúsculas y espacios)
    const estatusActual = (fila[idx.estatus] || "").toString().trim().toUpperCase();
    
    // Si el estatus no está vacío y NO es "CONVENIO", lo reportamos.
    if (estatusActual !== '' && estatusActual !== 'CONVENIO') {
      casosInconsistentes.push([
        fila[idx.folio],
        fila[idx.propietario],
        fila[idx.lote],
        fila[idx.estatus] // Mostramos el estatus incorrecto que se encontró
      ]);
    }
  });
  
  const cantidad = casosInconsistentes.length;
  return {
    resumen: `Se encontraron ${cantidad} convenios con un estatus de cartera inconsistente.`,
    datos: casosInconsistentes
  };
}

function verificarCambioDeAsesor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActivos = ss.getSheetByName('Convenios Activos');
  
  if (!hojaActivos) {
    return { resumen: 'No se pudo verificar la hoja "Convenios Activos" (no encontrada).', datos: [] };
  }
  
  const datos = hojaActivos.getDataRange().getValues();
  const encabezados = datos.shift();
  
  const idx = {
    ejecutivoLocal: encabezados.indexOf('Ejecutivo Asignado'),
    asesorCartera: encabezados.indexOf('Asesor (Cartera)'),
    folio: encabezados.indexOf('Folio'),
    propietario: encabezados.indexOf('Nombre Propietario'),
    lote: encabezados.indexOf('Lote Cartera')
  };
  
  if (idx.ejecutivoLocal === -1 || idx.asesorCartera === -1) {
    return { resumen: 'No se pudo verificar porque falta la columna "Ejecutivo Asignado" o "Asesor (Cartera)".', datos: [] };
  }
  
  const casosConCambio = [];
  datos.forEach(fila => {
    // Normalizamos ambos nombres para una comparación justa (sin espacios extra, mayúsculas, etc.)
    const ejecutivoLocal = (fila[idx.ejecutivoLocal] || "").toString().trim().toUpperCase().replace(/\s+/g, ' ');
    const asesorCartera = (fila[idx.asesorCartera] || "").toString().trim().toUpperCase().replace(/\s+/g, ' ');
    
    // Si ambos campos tienen un valor, no son "No encontrado" y son diferentes, lo reportamos.
    if (ejecutivoLocal && asesorCartera && asesorCartera !== 'NO ENCONTRADO' && ejecutivoLocal !== asesorCartera) {
      casosConCambio.push([
        fila[idx.folio],
        fila[idx.lote],
        fila[idx.propietario],
        fila[idx.ejecutivoLocal], // Asignación Original
        fila[idx.asesorCartera]   // Nueva Asignación en Cartera
      ]);
    }
  });
  
  const cantidad = casosConCambio.length;
  return {
    resumen: `Se encontraron ${cantidad} convenios con cambio de asesor en cartera.`,
    datos: casosConCambio
  };
}

function verificarEstatusDeVencidosEnCartera() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaVencidos = ss.getSheetByName('Convenios Vencidos');
  const hojaMaestra = ss.getSheetByName(HOJA_MAESTRA_LOCAL);
  
  if (!hojaVencidos || !hojaMaestra) {
    return { resumen: 'No se pudo verificar porque no se encontró la hoja "Convenios Vencidos" o "_ListasMaestras".', datos: [] };
  }
  
  // 1. Crear un mapa de Referencia -> Estatus desde la Cartera Maestra
  const datosMaestros = hojaMaestra.getDataRange().getValues();
  
  // --- INICIO DE LA CORRECCIÓN CLAVE ---
  // En lugar de buscar los encabezados, usamos los índices de columna fijos que ya conocemos.
  const idxReferenciaMaestra = 4; // Columna E
  const idxEstatusMaestra = 10;   // Columna K
  // --- FIN DE LA CORRECCIÓN CLAVE ---

  const carteraMap = new Map();
  datosMaestros.forEach(fila => {
    const referencia = fila[idxReferenciaMaestra];
    const estatus = fila[idxEstatusMaestra];
    if (referencia) {
      carteraMap.set(referencia, estatus);
    }
  });

  // 2. Recorrer la hoja de Vencidos y comparar (esta parte no cambia)
  const datosVencidos = hojaVencidos.getDataRange().getValues();
  const encabezadosVencidos = datosVencidos.shift();
  const idx = {
    referencia: encabezadosVencidos.indexOf('Referencia'),
    folio: encabezadosVencidos.indexOf('Folio'),
    propietario: encabezadosVencidos.indexOf('Nombre Propietario'),
    lote: encabezadosVencidos.indexOf('Lote Cartera')
  };

  if (idx.referencia === -1) {
    return { resumen: 'No se pudo verificar porque no se encontró la columna "Referencia" en "Convenios Vencidos".', datos: [] };
  }
  
  const casosAcorregir = [];
  datosVencidos.forEach(fila => {
    const referencia = fila[idx.referencia];
    if (referencia && carteraMap.has(referencia)) {
      const estatusCartera = (carteraMap.get(referencia) || "").toString().trim().toUpperCase();
      
      if (estatusCartera === 'CONVENIO') {
        casosAcorregir.push([
          fila[idx.folio],
          fila[idx.propietario],
          fila[idx.lote]
        ]);
      }
    }
  });
  
  const cantidad = casosAcorregir.length;
  return {
    resumen: `Se encontraron ${cantidad} convenios vencidos que aún tienen estatus "Convenio" en cartera.`,
    datos: casosAcorregir
  };
}
