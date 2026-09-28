/**
 * Configuración global del backend
 */
const CONFIG = {
  FOLDER_EVIDENCIAS_ID: "1OG1_geUGHvKck3BrBNAdStMygEhqfliQ",
  SHEETS: {
    ESTUDIANTES: "dicc_estudiantes",
    PRACTICAS: "dicc_practicas",
    ENUNCIADOS: "dicc_enunciados_detallados_x_practica",
    RESPUESTAS: "respuestas",
    EVIDENCIAS: "evidencias"
  }
};

/**
 * Sirve la interfaz web
 */
function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Registro de Prácticas - Ecología CCAA")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Carga inicial de datos para los selectores de cabecera
 */
function obtenerDatosIniciales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Estudiantes: Col A = nombre_apellidos, Col B = NIE
  const sheetEst = ss.getSheetByName(CONFIG.SHEETS.ESTUDIANTES);
  const dataEst = sheetEst.getDataRange().getValues();
  const estudiantes = dataEst.slice(1).map(r => ({
    nombre: String(r[0]).trim(),
    nie: String(r[1]).trim()
  })).filter(e => e.nie !== "");

  // 2. Prácticas: Col A = id_practica, Col B = nombre_practica
  const sheetPrac = ss.getSheetByName(CONFIG.SHEETS.PRACTICAS);
  const dataPrac = sheetPrac.getDataRange().getValues();
  const practicas = dataPrac.slice(1).map(r => ({
    id: String(r[0]).trim(),
    nombre: String(r[1]).trim()
  })).filter(p => p.id !== "");

  return { estudiantes, practicas };
}

/**
 * Obtener las sesiones numéricas disponibles para una práctica
 */
function obtenerSesionesPorPractica(idPractica) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEnun = ss.getSheetByName(CONFIG.SHEETS.ENUNCIADOS);
  const data = sheetEnun.getDataRange().getValues();
  const sesionesSet = new Set();

  for (let i = 1; i < data.length; i++) {
    const rowPractica = String(data[i][1]).trim();
    const rowSesion = data[i][2];
    if (rowPractica === String(idPractica).trim() && rowSesion !== "") {
      sesionesSet.add(Number(rowSesion));
    }
  }

  return Array.from(sesionesSet).sort((a, b) => a - b);
}

/**
 * Consulta enunciados activos filtrados por id_practica y sesion
 */
function obtenerEnunciadosActivos(idPractica, sesion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEnun = ss.getSheetByName(CONFIG.SHEETS.ENUNCIADOS);
  const data = sheetEnun.getDataRange().getValues();
  
  const sesionInt = parseInt(sesion, 10);
  const enunciados = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowPractica = String(row[1]).trim();
    const rowSesion = parseInt(row[2], 10);
    const rowActivo = String(row[5]).trim().toLowerCase();

    if (rowPractica === String(idPractica).trim() && rowSesion === sesionInt && (rowActivo === "sí" || rowActivo === "si")) {
      enunciados.push({
        id_enunciado: String(row[0]).trim(),
        tipo_pregunta: String(row[3]).trim(),
        enunciado: String(row[4]).trim()
      });
    }
  }

  return enunciados;
}

/**
 * Procesa la transacción atómica de envío (Respuestas + Evidencias)
 */
function procesarEnvio(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = new Date();
    const idEnvio = Utilities.getUuid();

    const nie = payload.nie;
    const idPractica = payload.idPractica || payload.id_practica;
    const sesionInt = parseInt(payload.sesion, 10);

    // 1. Guardar Respuestas (6 columnas normalizadas)
    if (payload.respuestas && payload.respuestas.length > 0) {
      const sheetResp = ss.getSheetByName(CONFIG.SHEETS.RESPUESTAS);
      const rowsRespuestas = payload.respuestas.map(r => [
        Utilities.getUuid(),
        idEnvio,
        timestamp,
        nie,
        r.id_enunciado,
        r.respuesta_larga
      ]);

      sheetResp.getRange(
        sheetResp.getLastRow() + 1, 
        1, 
        rowsRespuestas.length, 
        rowsRespuestas[0].length
      ).setValues(rowsRespuestas);
    }

    // 2. Guardar Evidencias
    if (payload.evidencias && payload.evidencias.length > 0) {
      const sheetEvid = ss.getSheetByName(CONFIG.SHEETS.EVIDENCIAS);
      const folderEvidencias = DriveApp.getFolderById(CONFIG.FOLDER_EVIDENCIAS_ID);
      const rowsEvidencias = [];

      for (const ev of payload.evidencias) {
        let driveUrl = "";
        
        const fileObj = ev.archivo || ev.fileData;
        if (fileObj && fileObj.data && fileObj.nombre) {
          const contentType = fileObj.tipo || "application/octet-stream";
          const bytes = Utilities.base64Decode(fileObj.data.split(",")[1] || fileObj.data);
          
          // Se preserva intacto el nombre original del archivo y su extensión
          const originalFilename = String(fileObj.nombre).trim();
          const blob = Utilities.newBlob(bytes, contentType, originalFilename);
          const file = folderEvidencias.createFile(blob);
          driveUrl = file.getUrl();
        }

        rowsEvidencias.push([
          Utilities.getUuid(),
          idEnvio,
          timestamp,
          nie,
          idPractica,
          sesionInt,
          ev.nombre || ev.nombre_evidencia || "",
          ev.descripcion || ev.descripcion_evidencia || "",
          ev.url || "",
          driveUrl
        ]);
      }

      sheetEvid.getRange(
        sheetEvid.getLastRow() + 1, 
        1, 
        rowsEvidencias.length, 
        rowsEvidencias[0].length
      ).setValues(rowsEvidencias);
    }

    return { exito: true, success: true, id_envio: idEnvio, mensaje: "Registro completado con éxito." };

  } catch (error) {
    return { exito: false, success: false, error: error.message, mensaje: error.message };
  } finally {
    lock.releaseLock();
  }
}