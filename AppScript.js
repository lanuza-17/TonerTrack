/**
 * CONFIGURACIÓN INTEGRADA
 */
const CONFIG = {
    CARPETA_REPORTES_ID: "11ToHdJxp48WxCKBlji_8KYVwRxvdHyzL", // Tu ID de carpeta
    ID_BITACORA_PRINCIPAL: "ID_DE_TU_BITACORA_REGIONES",
    ID_SHEET_AUDITORIA: "1XrBrjHzU_5T-NfHhYuAC-APuOf21r-QiXPCJicmddWw",

    EMAIL_ALERTAS: 'tu-correo@empresa.com',
    UMBRAL_CRITICO_EMAIL: 0.10, // 10%

    UMBRAL_CENTRAL: 0.3,
    UMBRAL_OTROS: 0.6,
    HOJAS_REGIONES: ["CENTRAL", "MANAGUA", "OCCIDENTE", "SUR", "NORTE", "CENTRO", "RAAN", "RASS"],

    COLORES: { ROJO: '#ff0000', VERDE: '#00b050' }
};

/**
 * FUNCIÓN MAESTRA: Descarga y Actualiza
 * Programa esta función en los activadores (9am, 12md, 4pm)
 */
function procesoCompletoDescargaYActualizacion() {
    console.log("--- 1. BUSCANDO NUEVOS REPORTES EN GMAIL ---");
    guardarAdjuntosCSV();

    console.log("--- 2. VERIFICANDO ARCHIVOS SIN PROCESAR ---");
    // Ahora siempre llamamos a ejecutarProcesoToner, 
    // y dentro validamos si el último archivo ya fue procesado o no.
    ejecutarProcesoToner();
}

/**
 * PARTE 1: DESCARGAR DE GMAIL
 */
function guardarAdjuntosCSV() {
    const carpeta = DriveApp.getFolderById(CONFIG.CARPETA_REPORTES_ID);
    const hilos = GmailApp.search('from:hp-sds-latam@insightportal.net has:attachment');
    let archivosDescargados = 0;

    for (let i = 0; i < hilos.length; i++) {
        const mensajes = hilos[i].getMessages();
        for (let j = 0; j < mensajes.length; j++) {
            const mensaje = mensajes[j];

            if (mensaje.isStarred()) continue; // Evita procesar correos ya usados

            const fechaFormato = Utilities.formatDate(mensaje.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
            const adjuntos = mensaje.getAttachments();

            for (let k = 0; k < adjuntos.length; k++) {
                const archivo = adjuntos[k];
                if (archivo.getName().endsWith(".csv")) {
                    const nuevoNombre = "reporte_" + fechaFormato + ".csv";
                    const existentes = carpeta.getFilesByName(nuevoNombre);

                    if (!existentes.hasNext()) {
                        carpeta.createFile(archivo).setName(nuevoNombre);
                        archivosDescargados++;
                    }
                }
            }
            mensaje.star(); // Marcar como procesado
        }
    }
    return (archivosDescargados > 0);
}

/**
 * PARTE 2: ACTUALIZAR BITÁCORA
 */
function ejecutarProcesoToner() {
    const folder = DriveApp.getFolderById(CONFIG.CARPETA_REPORTES_ID);
    const latestFile = getUltimoArchivo(folder);

    if (!latestFile) {
        console.log("La carpeta de reportes está vacía.");
        return;
    }

    // --- NUEVA LÓGICA: Verificar si ya fue procesado ---
    const prop = PropertiesService.getScriptProperties();
    const ultimoIdProcesado = prop.getProperty('ULTIMO_ARCHIVO_PROCESADO_ID');
    
    if (latestFile.getId() === ultimoIdProcesado) {
        console.log("No hay archivos nuevos sin procesar. Fin del proceso.");
        return;
    }

    console.log("Procesando nuevo archivo: " + latestFile.getName());
    const ss = SpreadsheetApp.openById(CONFIG.ID_BITACORA_PRINCIPAL);
    const dictData = cargarDatosReporte(latestFile);
    const ssLog = SpreadsheetApp.openById(CONFIG.ID_SHEET_AUDITORIA);
    let sheetLog = ssLog.getSheetByName("Cambios_Consolidado") || ssLog.insertSheet("Cambios_Consolidado");

    const timestamp = new Date();
    const logsParaAñadir = [];
    const alertasCriticas = [];

    CONFIG.HOJAS_REGIONES.forEach(nombreHoja => {
        const sheet = ss.getSheetByName(nombreHoja);
        if (!sheet) return;

        const range = sheet.getDataRange();
        const values = range.getValues();
        const backgrounds = range.getBackgrounds();
        const threshold = (nombreHoja === "CENTRAL" || nombreHoja === "MANAGUA") ? CONFIG.UMBRAL_CENTRAL : CONFIG.UMBRAL_OTROS;

        let cambiosEnHoja = false;

        for (let i = 1; i < values.length; i++) {
            const idVal = values[i][0];
            const sdsSerial = String(values[i][3]).trim();
            const sdsSKU = String(values[i][5]).trim();
            const key = `${sdsSerial}|${sdsSKU}`;

            if (!idVal || !dictData[key]) continue;

            const dataReporte = dictData[key];
            const oldNivel = values[i][6];
            const oldSNToner = String(values[i][7]).trim();

            // --- LÓGICA DE ALERTA ÚNICA ---
            // Solo alerta si ANTES era > 10% y AHORA es <= 10%
            if (oldNivel > CONFIG.UMBRAL_CRITICO_EMAIL && dataReporte.nivel <= CONFIG.UMBRAL_CRITICO_EMAIL) {
                alertasCriticas.push({
                    region: nombreHoja, hostname: values[i][1], sku: sdsSKU,
                    nivel: (dataReporte.nivel * 100).toFixed(0) + "%"
                });
            }

            let modificado = false;
            if (Math.round(oldNivel * 100) !== Math.round(dataReporte.nivel * 100)) {
                values[i][6] = dataReporte.nivel;
                logsParaAñadir.push([timestamp, nombreHoja, i + 1, idVal, values[i][1], sdsSKU, "Nivel", (oldNivel * 100).toFixed(0) + "%", (dataReporte.nivel * 100).toFixed(0) + "%", "Update"]);
                modificado = true;
            }

            if (oldSNToner !== "" && dataReporte.snToner !== "" && oldSNToner !== dataReporte.snToner) {
                values[i][7] = dataReporte.snToner;
                values[i][9] = timestamp;
                logsParaAñadir.push([timestamp, nombreHoja, i + 1, idVal, values[i][1], sdsSKU, "S/N Toner", oldSNToner, dataReporte.snToner, "CAMBIO"]);
                if (backgrounds[i][0] === CONFIG.COLORES.VERDE) {
                    for (let c = 0; c < 10; c++) backgrounds[i][c] = null;
                }
                modificado = true;
            }

            // Colores Rojo
            const esBlanco = (backgrounds[i][0] === '#ffffff' || backgrounds[i][0] === 'rgba(0,0,0,0)');
            if (esBlanco && dataReporte.nivel <= threshold) {
                for (let c = 0; c < 10; c++) backgrounds[i][c] = CONFIG.COLORES.ROJO;
                modificado = true;
            } else if (backgrounds[i][0] === CONFIG.COLORES.ROJO && dataReporte.nivel > threshold) {
                for (let c = 0; c < 10; c++) backgrounds[i][c] = null;
                modificado = true;
            }

            if (modificado) cambiosEnHoja = true;
        }

        if (cambiosEnHoja) {
            range.setValues(values);
            range.setBackgrounds(backgrounds);
        }
    });

    if (logsParaAñadir.length > 0) {
        sheetLog.getRange(sheetLog.getLastRow() + 1, 1, logsParaAñadir.length, logsParaAñadir[0].length).setValues(logsParaAñadir);
        sheetLog.getRange(2, 1, sheetLog.getLastRow() - 1, 10).sort({ column: 1, ascending: false });
    }

    if (alertasCriticas.length > 0) enviarCorreoAlertas(alertasCriticas);

    // Marcar el archivo como procesado exitosamente
    prop.setProperty('ULTIMO_ARCHIVO_PROCESADO_ID', latestFile.getId());
    console.log("Archivo procesado y guardado en el historial exitosamente.");
}

// ============================================================================
// FUNCIONES AUXILIARES
// ============================================================================

/**
 * Obtiene el archivo CSV más reciente de una carpeta de Drive.
 */
function getUltimoArchivo(folder) {
    const files = folder.getFilesByType(MimeType.CSV);
    let latestFile = null;
    let maxDate = 0;
    
    while (files.hasNext()) {
        const file = files.next();
        if (file.getDateCreated().getTime() > maxDate) {
            maxDate = file.getDateCreated().getTime();
            latestFile = file;
        }
    }
    return latestFile;
}

/**
 * Lee el archivo CSV y retorna un diccionario con la llave "Serial|SKU"
 */
function cargarDatosReporte(file) {
    const csvString = file.getBlob().getDataAsString();
    const csvData = Utilities.parseCsv(csvString);
    const dictData = {};
    
    if (csvData.length < 2) return dictData;
    
    // Tratamos de encontrar las columnas por el nombre del encabezado
    const headers = csvData[0].map(h => String(h).toLowerCase().trim());
    
    let idxSerial = headers.findIndex(h => h.includes('serial number') || h.includes('printer serial'));
    let idxSKU = headers.findIndex(h => h.includes('part number') || h.includes('sku') || h.includes('consumable part'));
    let idxNivel = headers.findIndex(h => h.includes('level') || h.includes('remaining'));
    let idxSNToner = headers.findIndex(h => h.includes('consumable serial') || h.includes('s/n toner'));
    
    // Fallback: Índices por defecto si no se encuentran las cabeceras
    if (idxSerial === -1) idxSerial = 0; 
    if (idxSKU === -1) idxSKU = 4;
    if (idxNivel === -1) idxNivel = 5;
    if (idxSNToner === -1) idxSNToner = 6;
    
    for (let i = 1; i < csvData.length; i++) {
        const row = csvData[i];
        if (row.length <= Math.max(idxSerial, idxSKU, idxNivel, idxSNToner)) continue;
        
        const serial = String(row[idxSerial]).trim();
        const sku = String(row[idxSKU]).trim();
        const key = `${serial}|${sku}`;
        
        let nivelRaw = String(row[idxNivel]).replace('%', '').trim();
        let nivel = parseFloat(nivelRaw);
        if (nivel > 1) nivel = nivel / 100; // Convertir 10 a 0.10
        
        const snToner = String(row[idxSNToner]).trim();
        
        dictData[key] = {
            nivel: isNaN(nivel) ? 0 : nivel,
            snToner: snToner
        };
    }
    
    return dictData;
}

/**
 * Envía un correo con la tabla de alertas críticas
 */
function enviarCorreoAlertas(alertas) {
    if (!CONFIG.EMAIL_ALERTAS) return;
    
    let htmlBody = "<h3>Alerta Crítica de Tóner (Nivel <= 10%)</h3>";
    htmlBody += "<table border='1' cellpadding='5' style='border-collapse: collapse;'>";
    htmlBody += "<tr><th>Región</th><th>Hostname</th><th>SKU</th><th>Nivel</th></tr>";
    
    alertas.forEach(alerta => {
        htmlBody += `<tr>
            <td>${alerta.region}</td>
            <td>${alerta.hostname}</td>
            <td>${alerta.sku}</td>
            <td style='color:red; font-weight:bold;'>${alerta.nivel}</td>
        </tr>`;
    });
    
    htmlBody += "</table>";
    
    MailApp.sendEmail({
        to: CONFIG.EMAIL_ALERTAS,
        subject: "ALERTA: Tóner Crítico en Impresoras",
        htmlBody: htmlBody
    });
}
