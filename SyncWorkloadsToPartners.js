/**
 * SyncWorkloadsToPartners.js
 * Distribuye la carga de trabajo consolidada de cada socio a su respectiva hoja de cálculo externa,
 * en una pestaña llamada "Workloads", con diseño limpio y profesional.
 */

function syncWorkloadsToPartners() {
  // Paso 0: Recuperar comentarios y status desde las planillas de los socios (origen de verdad)
  Logger.log("Iniciando la recuperación de comentarios y status desde las planillas de los socios...");
  try {
    pullPartnerUpdates();
  } catch (e) {
    Logger.log("Error al recuperar actualizaciones de socios: " + e.message + ". Se continuará con los datos existentes.");
  }

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
      headerRange.setBackground(null); // Reseteado (sin color de fondo)
      
      // Asegurar fuente normal para los datos
      targetSheet.getRange(2, 1, numRows - 1, numCols).setFontWeight("normal");

      // Auto-ajustar el ancho de las columnas para perfecta lectura
      for (var col = 1; col <= numCols; col++) {
        targetSheet.autoResizeColumn(col);
      }

      // Aplicar anchos de columna manuales y envoltura de texto según requerimientos
      // 1. Columna Y (Comentarios) -> 400px, Wrap
      var comentariosCol = getColIndex(headers, "Comentarios") + 1;
      if (comentariosCol === 0 && numCols >= 25) comentariosCol = 25; // Fallback a Y
      if (comentariosCol > 0 && comentariosCol <= numCols) {
        targetSheet.setColumnWidth(comentariosCol, 400);
        targetSheet.getRange(1, comentariosCol, numRows, 1).setWrap(true);
      }

      // 2. Columna AC (Workload Link) -> 100px
      var wlCol = getColIndex(headers, "Workload Link") + 1;
      if (wlCol === 0 && numCols >= 29) wlCol = 29; // Fallback a AC
      if (wlCol > 0 && wlCol <= numCols) {
        targetSheet.setColumnWidth(wlCol, 100);
      }

      // 3. Columna D (Columna 4) -> 250px, Wrap
      if (numCols >= 4) {
        targetSheet.setColumnWidth(4, 250);
        targetSheet.getRange(1, 4, numRows, 1).setWrap(true);
      }

      // Mostrar todas las columnas primero para asegurar un estado consistente
      try {
        var maxCols = targetSheet.getMaxColumns();
        if (maxCols > 0) {
          targetSheet.showColumns(1, maxCols);
        }
      } catch (e) {
        Logger.log("Advertencia al mostrar todas las columnas: " + e.message);
      }

      // Ocultar columnas solicitadas por el usuario para vista limpia del socio
      // Columnas: A, C, E-F, H-J, L-P, R-X, AA-AB
      hideColumnRangeSafely(targetSheet, 1, 1, numCols);   // A
      hideColumnRangeSafely(targetSheet, 3, 1, numCols);   // C
      hideColumnRangeSafely(targetSheet, 5, 2, numCols);   // E-F
      hideColumnRangeSafely(targetSheet, 8, 3, numCols);   // H-J
      hideColumnRangeSafely(targetSheet, 12, 5, numCols);  // L-P
      hideColumnRangeSafely(targetSheet, 18, 7, numCols);  // R-X
      hideColumnRangeSafely(targetSheet, 27, 2, numCols);  // AA-AB

      Logger.log("¡Socio '" + partner.name + "' sincronizado con éxito! " + partner.workloads.length + " filas actualizadas.");
      totalUpdatedPartners++;
    } catch (e) {
      Logger.log("Error al guardar datos en la planilla del socio '" + partner.name + "'. Detalle: " + e.message);
    }
  }

  Logger.log("Sincronización de cargas de trabajo a planillas de socios completada.");
  Logger.log("Socios actualizados con éxito: " + totalUpdatedPartners + " de " + partners.length + ". Cargas no coincidentes: " + unmatchedCount);
}

/**
 * Recupera los comentarios y status modificados por los socios en sus planillas individuales
 * y los aplica de vuelta a las pestañas principales en el archivo central.
 */
function pullPartnerUpdates() {
  var ssDest = SpreadsheetApp.getActiveSpreadsheet();
  var partnerSheet = ssDest.getSheetByName("Partner / Name");
  if (!partnerSheet) {
    Logger.log("No se encontró la pestaña 'Partner / Name' para recuperar actualizaciones.");
    return;
  }
  
  var partnerLastRow = partnerSheet.getLastRow();
  if (partnerLastRow < 2) {
    Logger.log("No hay socios configurados para recuperar actualizaciones.");
    return;
  }
  
  var partnerData = partnerSheet.getRange(2, 1, partnerLastRow - 1, 3).getValues();
  var partners = [];
  
  for (var p = 0; p < partnerData.length; p++) {
    var pName = partnerData[p][0] ? partnerData[p][0].toString().trim() : "";
    var pSpreadsheetId = partnerData[p][2] ? partnerData[p][2].toString().trim() : "";
    
    if (pSpreadsheetId) {
      partners.push({
        name: pName,
        spreadsheetId: pSpreadsheetId
      });
    }
  }
  
  if (partners.length === 0) {
    Logger.log("No hay socios válidos configurados para recuperar actualizaciones.");
    return;
  }
  
  var updates = {};
  var totalPulled = 0;
  
  for (var p = 0; p < partners.length; p++) {
    var partner = partners[p];
    try {
      var partnerSs = SpreadsheetApp.openById(partner.spreadsheetId);
      var targetSheet = partnerSs.getSheetByName("Workloads");
      if (!targetSheet) {
        Logger.log("Socio '" + partner.name + "': No se encontró la pestaña 'Workloads'.");
        continue;
      }
      
      var lastRow = targetSheet.getLastRow();
      var lastCol = targetSheet.getLastColumn();
      if (lastRow < 2 || lastCol === 0) {
        Logger.log("Socio '" + partner.name + "': La pestaña 'Workloads' está vacía.");
        continue;
      }
      
      var values = targetSheet.getRange(1, 1, lastRow, lastCol).getValues();
      var headers = values[0];
      
      var wIdIdx = 0; // Column A
      var comentariosIdx = getColIndex(headers, "Comentarios");
      var statusIdx = getColIndex(headers, "Status");
      
      if (comentariosIdx === -1) {
        Logger.log("Socio '" + partner.name + "': Advertencia: No se encontró la columna 'Comentarios'. Usando columna Y (index 24) como fallback.");
        comentariosIdx = 24;
      }
      if (statusIdx === -1) {
        Logger.log("Socio '" + partner.name + "': Advertencia: No se encontró la columna 'Status'. Usando columna Z (index 25) como fallback.");
        statusIdx = 25;
      }
      
      var partnerUpdatesCount = 0;
      for (var r = 1; r < values.length; r++) {
        var row = values[r];
        var wId = row[wIdIdx] ? row[wIdIdx].toString().trim() : "";
        if (wId) {
          var comentarios = comentariosIdx < row.length ? row[comentariosIdx].toString().trim() : "";
          var status = statusIdx < row.length ? row[statusIdx].toString().trim() : "";
          
          updates[wId] = {
            comentarios: comentarios,
            status: status,
            partnerName: partner.name
          };
          partnerUpdatesCount++;
          totalPulled++;
        }
      }
      Logger.log("Socio '" + partner.name + "': Leídas " + partnerUpdatesCount + " filas para actualización.");
    } catch (e) {
      Logger.log("Error al recuperar datos del socio '" + partner.name + "': " + e.message);
    }
  }
  
  if (totalPulled === 0) {
    Logger.log("No se encontraron datos para actualizar desde las planillas de los socios.");
    return;
  }
  
  Logger.log("Total de registros leídos para actualizar: " + totalPulled + ". Aplicando a pestañas principales...");
  
  var sheetsToSync = CONFIG.CORE_SHEETS_TO_SYNC;
  var totalUpdatedCells = 0;
  
  for (var i = 0; i < sheetsToSync.length; i++) {
    var sheetName = sheetsToSync[i];
    var sheet = ssDest.getSheetByName(sheetName);
    if (!sheet) continue;
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol === 0) continue;
    
    var range = sheet.getRange(1, 1, lastRow, lastCol);
    var values = range.getValues();
    var headers = values[0];
    
    var wIdIdx = 0; // Column A
    var comentariosIdx = getColIndex(headers, "Comentarios");
    var statusIdx = getColIndex(headers, "Status");
    
    var headersChanged = false;
    if (comentariosIdx === -1) {
      comentariosIdx = headers.length;
      headers.push("Comentarios");
      headersChanged = true;
      Logger.log("Creada columna 'Comentarios' en pestaña principal '" + sheetName + "'");
    }
    if (statusIdx === -1) {
      statusIdx = headers.length;
      headers.push("Status");
      headersChanged = true;
      Logger.log("Creada columna 'Status' en pestaña principal '" + sheetName + "'");
    }
    
    var sheetModified = false;
    var sheetUpdatedCount = 0;
    
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      while (row.length < headers.length) {
        row.push("");
      }
      
      var wId = row[wIdIdx] ? row[wIdIdx].toString().trim() : "";
      if (wId && updates.hasOwnProperty(wId)) {
        var update = updates[wId];
        
        if (update.comentarios !== undefined && row[comentariosIdx].toString().trim() !== update.comentarios) {
          row[comentariosIdx] = update.comentarios;
          sheetModified = true;
          sheetUpdatedCount++;
          totalUpdatedCells++;
        }
        if (update.status !== undefined && row[statusIdx].toString().trim() !== update.status) {
          row[statusIdx] = update.status;
          sheetModified = true;
          sheetUpdatedCount++;
          totalUpdatedCells++;
        }
      }
      values[r] = row;
    }
    
    if (sheetModified || headersChanged) {
      if (headersChanged) {
        values[0] = headers;
      }
      var numRows = values.length;
      var numCols = headers.length;
      sheet.getRange(1, 1, numRows, numCols).setValues(values);
      Logger.log("Pestaña '" + sheetName + "' actualizada con " + sheetUpdatedCount + " cambios de socios.");
    }
  }
  
  Logger.log("Proceso de recuperación completado. Total de celdas actualizadas en origen: " + totalUpdatedCells);
}

/**
 * Oculta un rango de columnas de forma segura, asegurando no exceder el número total de columnas.
 */
function hideColumnRangeSafely(sheet, startCol, numColsToHide, totalCols) {
  if (startCol <= totalCols) {
    var colsToHide = Math.min(numColsToHide, totalCols - startCol + 1);
    try {
      sheet.hideColumns(startCol, colsToHide);
    } catch (e) {
      Logger.log("Error al ocultar columnas desde " + startCol + " (" + colsToHide + "): " + e.message);
    }
  }
}
