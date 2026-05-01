/**
 * SISTEMA AUTOMATIZADO DE MONITOREO DE TONERS
 * Configuración de IDs y Parámetros
 */
const CONFIG = {
  // IDs de Drive
  FOLDER_ID_REPORTES: '11ToHdJxp48WxCKBlji_8KYVwRxvdHyzL',
  FOLDER_ID_PROCESADOS: '1ck9w2MahQXESI2DICb2irKWpwA1hUi96',
 // ID_DE_TU_BITACORA_PRINCIPAL:'1XrBrjHzU_5T-NfHhYuAC-APuOf21r-QiXPCJicmddWw',
  ID_SHEET_AUDITORIA:'1W42h397de9coivILhD4X0F1tITtSH9sNt8-Uq0ZXqz0',

  // Alertas
  EMAIL_ALERTAS: 'carlanuza0020@gmail.com', // Correo donde llegarán las alertas críticas
  UMBRAL_CRITICO_EMAIL: 0.10, // 10% para enviar correo

  // Reglas de Negocio
  UMBRAL_CENTRAL: 0.3,
  UMBRAL_OTROS: 0.6,
  HOJAS_REGIONES: ["CENTRAL", "MANAGUA", "OCCIDENTE", "SUR", "NORTE", "CENTRO", "RAAN", "RASS"],

  // Estilos
  COLORES: {
    ROJO: '#ff0000',
    VERDE: '#00b050',
    BLANCO: '#ffffff'
  }
};
function guardarAdjuntosCSV() {
  var carpetaId = "11ToHdJxp48WxCKBlji_8KYVwRxvdHyzL";
  var carpeta = DriveApp.getFolderById(carpetaId);

  var hilos = GmailApp.search('from:hp-sds-latam@insightportal.net has:attachment');

  for (var i = 0; i < hilos.length; i++) {
    var mensajes = hilos[i].getMessages();

    for (var j = 0; j < mensajes.length; j++) {
      var mensaje = mensajes[j];

      // 🔒 Evitar procesar correos ya usados
      if (mensaje.isStarred()) continue;

      var fecha = mensaje.getDate();
      var fechaFormato = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");

      var adjuntos = mensaje.getAttachments();

      for (var k = 0; k < adjuntos.length; k++) {
        var archivo = adjuntos[k];

        if (archivo.getName().endsWith(".csv")) {

          // 🆕 Nuevo nombre con fecha
          var nuevoNombre = "reporte_" + fechaFormato + ".csv";

          // 🔍 Verificar si ya existe
          var existentes = carpeta.getFilesByName(nuevoNombre);

          if (!existentes.hasNext()) {
            carpeta.createFile(archivo).setName(nuevoNombre);
          }
        }
      }

      // ⭐ Marcar como procesado
      mensaje.star();
    }
  }
}
/**
 * FUNCIÓN PRINCIPAL - Programa esto en los activadores
 */
function TonerTrack () {

  guardarAdjuntosCSV();
  const ss = SpreadsheetApp.openById('1XrBrjHzU_5T-NfHhYuAC-APuOf21r-QiXPCJicmddWw'); 
  const folderEntrada = DriveApp.getFolderById(CONFIG.FOLDER_ID_REPORTES);
  const latestFile = getUltimoArchivo(folderEntrada);

  if (!latestFile) {
    console.log("No se encontraron archivos nuevos para procesar.");
    return;
  }

  console.log("Iniciando procesamiento de: " + latestFile.getName());

  const dictData = cargarDatosReporte(latestFile);
  const ssLog = SpreadsheetApp.openById(CONFIG.ID_SHEET_AUDITORIA);
  let sheetLog = ssLog.getSheetByName("Cambios_Consolidado") || ssLog.insertSheet("Cambios_Consolidado");

  const timestamp = new Date();
  const logsParaAñadir = [];
  const alertasCriticas = []; // Para el correo

  CONFIG.HOJAS_REGIONES.forEach(nombreHoja => {
    const sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) return;

    const range = sheet.getDataRange();
    const values = range.getValues();
    const backgrounds = range.getBackgrounds();
    const threshold = (nombreHoja === "CENTRAL" || nombreHoja === "MANAGUA") ? CONFIG.UMBRAL_CENTRAL : CONFIG.UMBRAL_OTROS;

    let cambiosEnEstaHoja = false;

    for (let i = 1; i < values.length; i++) {
      const idVal = values[i][0];
      const sdsSerial = String(values[i][3]).trim();
      const sdsSKU = String(values[i][5]).trim();
      const key = `${sdsSerial}|${sdsSKU}`;

      if (!idVal || !dictData[key]) continue;

      const dataReporte = dictData[key];
      const oldNivel = values[i][6];
      const oldSNToner = String(values[i][7]).trim();
      const currentBG = backgrounds[i][0];

      let filaModificada = false;

      // 1. Verificar Alerta Crítica (Correo)
      if (dataReporte.nivel <= CONFIG.UMBRAL_CRITICO_EMAIL) {
        alertasCriticas.push({
          region: nombreHoja,
          hostname: values[i][1],
          sku: sdsSKU,
          nivel: (dataReporte.nivel * 100).toFixed(0) + "%"
        });
      }

      // 2. Actualizar Nivel Actual
      if (Math.round(oldNivel * 100) !== Math.round(dataReporte.nivel * 100)) {
        values[i][6] = dataReporte.nivel;
        logsParaAñadir.push([timestamp, nombreHoja, i + 1, idVal, values[i][1], sdsSKU, "Nivel", (oldNivel * 100).toFixed(0) + "%", (dataReporte.nivel * 100).toFixed(0) + "%", "Update"]);
        filaModificada = true;
      }

      // 3. Cambio de Toner Detectado
      if (oldSNToner !== "" && dataReporte.snToner !== "" && oldSNToner !== dataReporte.snToner) {
        values[i][7] = dataReporte.snToner;
        values[i][9] = timestamp;
        logsParaAñadir.push([timestamp, nombreHoja, i + 1, idVal, values[i][1], sdsSKU, "S/N Toner", oldSNToner, dataReporte.snToner, "CAMBIO"]);

        if (currentBG === CONFIG.COLORES.VERDE) {
          for (let c = 0; c < 10; c++) backgrounds[i][c] = null;
        }
        filaModificada = true;
      }

      // 4. Lógica de Colores (Rojo/Blanco)
      const esBlanco = (currentBG === '#ffffff' || currentBG === 'rgba(0,0,0,0)');
      if (esBlanco && dataReporte.nivel <= threshold) {
        for (let c = 0; c < 10; c++) backgrounds[i][c] = CONFIG.COLORES.ROJO;
        filaModificada = true;
      } else if (currentBG === CONFIG.COLORES.ROJO && dataReporte.nivel > threshold) {
        for (let c = 0; c < 10; c++) backgrounds[i][c] = null;
        filaModificada = true;
      }

      if (filaModificada) cambiosEnEstaHoja = true;
    }

    if (cambiosEnEstaHoja) {
      range.setValues(values);
      range.setBackgrounds(backgrounds);
    }
  });

  // Guardar Auditoría
  if (logsParaAñadir.length > 0) {
    sheetLog.getRange(sheetLog.getLastRow() + 1, 1, logsParaAñadir.length, logsParaAñadir[0].length).setValues(logsParaAñadir);
    sheetLog.getRange(2, 1, sheetLog.getLastRow() - 1, 10).sort({ column: 1, ascending: false });
  }

  // Enviar Alertas por Correo
  if (alertasCriticas.length > 0) {
    enviarCorreoAlertas(alertasCriticas);
  }

  // Mover archivo a procesados
  moverAProcesados(latestFile);
  console.log("Proceso completado exitosamente.");
}

/**
 * FUNCIONES AUXILIARES
 */

function getUltimoArchivo(folder) {
  const files = folder.getFiles();
  let latest = null;
  let latestDate = 0;
  while (files.hasNext()) {
    const file = files.next();
    const date = file.getDateCreated().getTime();
    if (date > latestDate) {
      latestDate = date;
      latest = file;
    }
  }
  return latest;
}

function cargarDatosReporte(file) {
  const contenido = file.getBlob().getDataAsString();
  const data = Utilities.parseCsv(contenido); // Convierte el texto CSV en una matriz

  const dict = {};
  for (let i = 1; i < data.length; i++) {
    // IMPORTANTE: Verifica que las columnas coincidan con tu CSV
    // Según tu lógica anterior: Col C (índice 2), Col E (índice 4), Col F (índice 5), Col G (índice 6)
    const serie = String(data[i][2]).trim();
    const sku = String(data[i][4]).trim();
    let nivelRaw = String(data[i][5]);

    // Limpiar el nivel si viene con % o como texto
    let nivel = parseFloat(nivelRaw.replace('%', ''));
    if (nivel > 1) nivel = nivel / 100; // Si es 75 lo vuelve 0.75

    if (serie && sku) {
      dict[`${serie}|${sku}`] = {
        nivel: nivel,
        snToner: String(data[i][6]).trim()
      };
    }
  }
  return dict;
}


function enviarCorreoAlertas(alertas) {
  let cuerpo = "<h2>Alerta de Toners Críticos (Menos del 10%)</h2>";
  cuerpo += "<table border='1' style='border-collapse: collapse; padding: 5px;'><tr><th>Región</th><th>Hostname</th><th>SKU</th><th>Nivel</th></tr>";

  alertas.forEach(a => {
    cuerpo += `<tr><td>${a.region}</td><td>${a.hostname}</td><td>${a.sku}</td><td style='color: red;'><b>${a.nivel}</b></td></tr>`;
  });

  cuerpo += "</table><br><p>Por favor, proceda con el cambio de suministros.</p>";

  MailApp.sendEmail({
    to: CONFIG.EMAIL_ALERTAS,
    subject: "⚠️ ALERTA: Niveles de Toner Críticos",
    htmlBody: cuerpo
  });
}

function moverAProcesados(file) {
  const destFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID_PROCESADOS);
  file.moveTo(destFolder);
}
