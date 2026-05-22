/**
 * SyncPartners.js
 * Sincroniza los Spreadsheet IDs de los socios en la pestaña "Partner / Name"
 * buscando el dominio de cada socio en la base de datos maestra de socios "LATAM_Partner_DB_2026_v3".
 */

function syncPartnerSpreadsheetIds() {
  var ssDest = SpreadsheetApp.getActiveSpreadsheet();
  if (!ssDest) {
    Logger.log("No se encontró una hoja de cálculo activa.");
    return;
  }

  var partnerSheet = ssDest.getSheetByName("Partner / Name");
  if (!partnerSheet) {
    Logger.log("No se encontró la pestaña 'Partner / Name' en la hoja de cálculo activa.");
    return;
  }

  var lastRow = partnerSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No hay filas con datos para procesar en la pestaña 'Partner / Name'.");
    return;
  }

  // Abrir la hoja de cálculo de la Base de Datos de Socios (Partner DB)
  var dbSpreadsheetId = CONFIG.PARTNER_DB_SPREADSHEET_ID;
  if (!dbSpreadsheetId) {
    Logger.log("Error: CONFIG.PARTNER_DB_SPREADSHEET_ID no está configurado en Config.js.");
    return;
  }

  var ssDb;
  try {
    ssDb = SpreadsheetApp.openById(dbSpreadsheetId);
  } catch (e) {
    Logger.log("Error al abrir la base de datos de socios con ID: " + dbSpreadsheetId + ". Detalle: " + e.message);
    return;
  }

  var dbSheetName = "LATAM_Partner_DB_2026_v3";
  var dbSheet = ssDb.getSheetByName(dbSheetName);
  if (!dbSheet) {
    Logger.log("Error: No se encontró la pestaña '" + dbSheetName + "' en la base de datos de socios.");
    return;
  }

  var dbLastRow = dbSheet.getLastRow();
  var dbLastCol = dbSheet.getLastColumn();
  if (dbLastRow < 2 || dbLastCol === 0) {
    Logger.log("La pestaña de base de datos '" + dbSheetName + "' está vacía o no tiene suficientes filas.");
    return;
  }

  var dbValues = dbSheet.getRange(1, 1, dbLastRow, dbLastCol).getValues();
  var dbHeaders = dbValues[0];

  // Determinar los índices de columna para Dominio (H por defecto) e ID de Spreadsheet (K por defecto)
  var domainColIdx = 7; // Columna H (0-indexed)
  var idColIdx = 10;    // Columna K (0-indexed)

  // Intentar buscar de forma dinámica por el nombre del encabezado para máxima robustez
  for (var c = 0; c < dbHeaders.length; c++) {
    var headerName = dbHeaders[c] ? dbHeaders[c].toString().trim().toLowerCase() : "";
    if (headerName === "domain" || headerName === "dominio") {
      domainColIdx = c;
    } else if (
      headerName.indexOf("spreadsheet id") !== -1 || 
      headerName === "spreadsheet_id" || 
      headerName === "spreadsheetid" ||
      headerName === "id de spreadsheet"
    ) {
      idColIdx = c;
    }
  }

  Logger.log("Usando la columna con índice " + domainColIdx + " para Dominios y columna " + idColIdx + " para Spreadsheet IDs en la base de datos.");

  // Construir el mapa de dominio -> spreadsheetId
  var partnerDbMap = {};
  for (var r = 1; r < dbValues.length; r++) {
    var row = dbValues[r];
    var domainVal = row[domainColIdx] ? row[domainColIdx].toString().trim().toLowerCase() : "";
    var spreadsheetIdVal = row[idColIdx] ? row[idColIdx].toString().trim() : "";
    if (domainVal && spreadsheetIdVal) {
      partnerDbMap[domainVal] = spreadsheetIdVal;
    }
  }

  // Leer datos actuales de la pestaña "Partner / Name" (Columna B y C, fila 2 en adelante)
  // Columna B = Dominio, Columna C = Spreadsheet ID
  var range = partnerSheet.getRange(2, 2, lastRow - 1, 2);
  var data = range.getValues();
  var updatedCount = 0;

  for (var i = 0; i < data.length; i++) {
    var partnerDomain = data[i][0] ? data[i][0].toString().trim().toLowerCase() : "";
    var oldSpreadsheetId = data[i][1] ? data[i][1].toString().trim() : "";
    
    if (partnerDomain) {
      if (partnerDbMap.hasOwnProperty(partnerDomain)) {
        var targetId = partnerDbMap[partnerDomain];
        if (oldSpreadsheetId !== targetId) {
          data[i][1] = targetId;
          updatedCount++;
          Logger.log("Socio con dominio '" + partnerDomain + "' actualizado con ID: " + targetId);
        }
      } else {
        Logger.log("Advertencia: No se encontró el dominio '" + partnerDomain + "' en la base de datos de socios.");
      }
    }
  }

  if (updatedCount > 0) {
    range.setValues(data);
    Logger.log("¡Sincronización de socios completada! Se actualizaron " + updatedCount + " fila(s).");
  } else {
    Logger.log("No se encontraron cambios en los Spreadsheet IDs de los socios. Todo al día.");
  }
}
