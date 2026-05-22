/**
 * SyncWorkloadsToPartners.js
 * Distribuye la carga de trabajo consolidada de cada socio a su respectiva hoja de cálculo externa,
 * en una pestaña llamada "Workloads", con diseño limpio y profesional.
 */

function syncWorkloadsToPartners() {
  // Paso 1: Asegurar que la pestaña consolidada esté completamente actualizada con los últimos datos
  Logger.log("Iniciando la consolidación de cargas de trabajo antes de distribuir...");
  try {
    createConsolidatedWorkloads();
  } catch (e) {
    Logger.log("Error al actualizar la pestaña consolidada: " + e.message + ". Se continuará con los datos existentes.");
  }

  var ssDest = SpreadsheetApp.getActiveSpreadsheet();
  if (!ssDest) {
    Logger.log("No se encontró una hoja de cálculo activa.");
    return;
  }

  // Paso 2: Obtener la lista de socios, sus dominios y sus IDs de hoja de cálculo
  var partnerSheet = ssDest.getSheetByName("Partner / Name");
  if (!partnerSheet) {
    Logger.log("No se encontró la pestaña 'Partner / Name' en la hoja de cálculo activa.");
    return;
  }

  var partnerLastRow = partnerSheet.getLastRow();
  if (partnerLastRow < 2) {
    Logger.log("No hay socios configurados en la pestaña 'Partner / Name'.");
    return;
  }

  var partnerData = partnerSheet.getRange(2, 1, partnerLastRow - 1, 3).getValues();
  var partners = [];

  for (var p = 0; p < partnerData.length; p++) {
    var pName = partnerData[p][0] ? partnerData[p][0].toString().trim() : "";
    var pDomain = partnerData[p][1] ? partnerData[p][1].toString().trim().toLowerCase() : "";
    var pSpreadsheetId = partnerData[p][2] ? partnerData[p][2].toString().trim() : "";

    if (pSpreadsheetId) {
      partners.push({
        name: pName,
        domain: pDomain,
        spreadsheetId: pSpreadsheetId,
        workloads: []
      });
    } else {
      Logger.log("Advertencia: El socio '" + pName + "' no tiene un ID de hoja de cálculo configurado. Se omitirá.");
    }
  }

  if (partners.length === 0) {
    Logger.log("No hay socios válidos con ID de hoja de cálculo configurados.");
    return;
  }

  // Paso 3: Leer las cargas de trabajo consolidadas
  var consolSheet = ssDest.getSheetByName("Consolidated Workloads");
  if (!consolSheet) {
    Logger.log("Error: No se encontró la pestaña 'Consolidated Workloads'.");
    return;
  }

  var consolLastRow = consolSheet.getLastRow();
  var consolLastCol = consolSheet.getLastColumn();
  if (consolLastRow < 2 || consolLastCol === 0) {
    Logger.log("La pestaña 'Consolidated Workloads' está vacía.");
    return;
  }

  var consolValues = consolSheet.getRange(1, 1, consolLastRow, consolLastCol).getValues();
  var headers = consolValues[0];

  // Identificar el índice de la columna del Socio/Partner en el Consolidated sheet
  var partnerColIdx = -1;
  var partnerHeaderNames = ["partner", "partner name", "partner / name", "socio", "domain", "partner domain"];
  for (var mh = 0; mh < headers.length; mh++) {
    var hNameLower = headers[mh].toString().trim().toLowerCase();
    if (partnerHeaderNames.indexOf(hNameLower) !== -1) {
      partnerColIdx = mh;
      break;
    }
  }

  if (partnerColIdx === -1) {
    Logger.log("Error: No se pudo encontrar la columna de Socio/Partner en los encabezados de la consolidación.");
    Logger.log("Encabezados encontrados: " + JSON.stringify(headers));
    return;
  }

  Logger.log("Identificada la columna de Socio en la consolidación en el índice: " + partnerColIdx + " (" + headers[partnerColIdx] + ")");

  // Paso 4: Agrupar las filas de la consolidación por socio
  var unmatchedCount = 0;
  for (var r = 1; r < consolValues.length; r++) {
    var row = consolValues[r];
    var workloadPartnerVal = row[partnerColIdx] ? row[partnerColIdx].toString().trim().toLowerCase() : "";
    
    if (!workloadPartnerVal) {
      continue; // Fila sin socio asignado
    }

    var matched = false;
    for (var p = 0; p < partners.length; p++) {
      var partner = partners[p];
      
      // Intentar coincidencia por dominio o por nombre
      var isDomainMatch = partner.domain && (workloadPartnerVal === partner.domain || workloadPartnerVal.indexOf(partner.domain) !== -1);
      var isNameMatch = partner.name && workloadPartnerVal === partner.name.toLowerCase();

      if (isDomainMatch || isNameMatch) {
        partner.workloads.push(row);
        matched = true;
        break;
      }
    }

    if (!matched) {
      unmatchedCount++;
      Logger.log("Advertencia: Fila " + (r + 1) + " con valor de socio '" + row[partnerColIdx] + "' no coincidió con ningún socio configurado.");
    }
  }

  Logger.log("Distribución completada en memoria. Enviando datos a hojas de socios externos...");

  // Paso 5: Para cada socio, abrir su planilla y actualizar la pestaña "Workloads"
  var totalUpdatedPartners = 0;
  for (var p = 0; p < partners.length; p++) {
    var partner = partners[p];
    
    if (partner.workloads.length === 0) {
      Logger.log("El socio '" + partner.name + "' no tiene cargas de trabajo asignadas en esta sincronización.");
      continue;
    }

    Logger.log("Socio '" + partner.name + "': procesando " + partner.workloads.length + " cargas de trabajo...");

    var partnerSs;
    try {
      partnerSs = SpreadsheetApp.openById(partner.spreadsheetId);
    } catch (e) {
      Logger.log("Error crítico: No se pudo abrir la hoja de cálculo del socio '" + partner.name + "' (ID: " + partner.spreadsheetId + "). Detalle: " + e.message);
      continue; // Continuar con el siguiente socio
    }

    var targetSheetName = "Workloads";
    var targetSheet = partnerSs.getSheetByName(targetSheetName);
    
    if (!targetSheet) {
      try {
        targetSheet = partnerSs.insertSheet(targetSheetName);
        Logger.log("Creada nueva pestaña '" + targetSheetName + "' para el socio '" + partner.name + "'.");
      } catch (e) {
        Logger.log("Error al crear la pestaña '" + targetSheetName + "' para el socio '" + partner.name + "'. Detalle: " + e.message);
        continue;
      }
    }

    // Limpiar contenido previo
    try {
      targetSheet.clear();
    } catch (e) {
      Logger.log("Error al limpiar la pestaña '" + targetSheetName + "' para el socio '" + partner.name + "': " + e.message);
      // Recrear en caso de falla catastrófica por bloqueos de tablas de Google
      var sheetPos = targetSheet.getIndex();
      partnerSs.deleteSheet(targetSheet);
      targetSheet = partnerSs.insertSheet(targetSheetName, sheetPos - 1);
    }

    // Construir matriz de datos a escribir
    var valuesToWrite = [headers].concat(partner.workloads);
    var numRows = valuesToWrite.length;
    var numCols = headers.length;

    try {
      // Escribir los valores
      var targetRange = targetSheet.getRange(1, 1, numRows, numCols);
      targetRange.setValues(valuesToWrite);

      // Aplicar diseño Premium y Profesional
      targetSheet.setFrozenRows(1);
      
      var headerRange = targetSheet.getRange(1, 1, 1, numCols);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#f3f3f3"); // Gris limpio
      
      // Asegurar fuente normal para los datos
      targetSheet.getRange(2, 1, numRows - 1, numCols).setFontWeight("normal");

      // Auto-ajustar el ancho de las columnas para perfecta lectura
      for (var col = 1; col <= numCols; col++) {
        targetSheet.autoResizeColumn(col);
      }

      Logger.log("¡Socio '" + partner.name + "' sincronizado con éxito! " + partner.workloads.length + " filas actualizadas.");
      totalUpdatedPartners++;
    } catch (e) {
      Logger.log("Error al guardar datos en la planilla del socio '" + partner.name + "'. Detalle: " + e.message);
    }
  }

  Logger.log("Sincronización de cargas de trabajo a planillas de socios completada.");
  Logger.log("Socios actualizados con éxito: " + totalUpdatedPartners + " de " + partners.length + ". Cargas no coincidentes: " + unmatchedCount);
}
