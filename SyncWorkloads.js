/**
 * SyncWorkloads.js
 * Sincroniza las 4 pestañas principales (Committed/Uncommitted) desde el origen al destino,
 * preservando comentarios, status y registrando cambios en Change Log.
 */

function syncWorkloadSheets() {
  var ssSource = SpreadsheetApp.openById(CONFIG.SOURCE_SPREADSHEET_ID);
  var ssDest = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ssDest) {
    Logger.log("No se encontró una hoja de cálculo activa.");
    return;
  }
  
  var sheetsToSync = CONFIG.CORE_SHEETS_TO_SYNC;
  var logEntries = [];
  var timestamp = new Date();
  var cutoffTime = timestamp.getTime() - (7 * 24 * 60 * 60 * 1000);
  var activeHighlights = {};
  
  // Get or create Change Log sheet
  var logSheetName = "Change Log";
  var logSheet = ssDest.getSheetByName(logSheetName);
  if (!logSheet) {
    logSheet = ssDest.insertSheet(logSheetName);
    logSheet.appendRow(["Timestamp", "Sheet Name", "Workload ID", "Change Type", "Field", "Old Value", "New Value"]);
    Logger.log("Pestaña de registro de cambios creada: " + logSheetName);
  } else {
    try {
      var lastLogRow = logSheet.getLastRow();
      if (lastLogRow > 1) {
        var logData = logSheet.getRange(2, 1, lastLogRow - 1, 4).getValues();
        for (var l = 0; l < logData.length; l++) {
          var lTime = new Date(logData[l][0]).getTime();
          var lSheet = logData[l][1];
          var lWId = logData[l][2] ? logData[l][2].toString().trim() : "";
          var lType = logData[l][3];
          if (lTime >= cutoffTime && lWId) {
            var key = lSheet + "_" + lWId;
            if (!activeHighlights[key] || lType === "NEW WORKLOAD") {
              activeHighlights[key] = lType;
            }
          }
        }
      }
    } catch (e) {
      Logger.log("Advertencia: No se pudo leer el Change Log para los colores de 7 días: " + e.message);
    }
  }
  
  for (var i = 0; i < sheetsToSync.length; i++) {
    var sheetName = sheetsToSync[i];
    var sourceSheet = ssSource.getSheetByName(sheetName);
    
    if (!sourceSheet) {
      Logger.log("La pestaña origen no existe: " + sheetName);
      var destSheet = ssDest.getSheetByName(sheetName);
      if (!destSheet) {
        destSheet = ssDest.insertSheet(sheetName);
        Logger.log("Pestaña vacía creada en destino (origen faltante): " + sheetName);
      }
      continue;
    }
    
    var lastSourceRow = sourceSheet.getLastRow();
    var lastSourceCol = sourceSheet.getLastColumn();
    
    if (lastSourceRow === 0 || lastSourceCol === 0) {
      Logger.log("Pestaña origen vacía: " + sheetName);
      continue;
    }
    
    var sourceValues = sourceSheet.getRange(1, 1, lastSourceRow, lastSourceCol).getValues();
    var headers = sourceValues[0].slice();
    
    // Ensure no header cell is empty to prevent Google Sheets Table errors
    for (var h = 0; h < headers.length; h++) {
      if (headers[h] === null || headers[h] === undefined || headers[h].toString().trim() === "") {
        headers[h] = "Columna_" + (h + 1);
      }
    }
    
    var workloadIdIdx = 0; // Column A
    var prodDateIdx = getColIndex(headers, "Production Date");
    var arrIdx = getColIndex(headers, "ARR");
    var progressIdx = getColIndex(headers, "Progress");
    
    var filteredValues = [headers];
    for (var r = 1; r < sourceValues.length; r++) {
      var row = sourceValues[r];
      var wId = row[workloadIdIdx] ? row[workloadIdIdx].toString().trim() : "";
      if (wId) {
        filteredValues.push(row.slice());
      }
    }
    
    var destSheet = ssDest.getSheetByName(sheetName);
    var oldWorkloads = {};
    var destHadData = false;
    
    if (!destSheet) {
      destSheet = ssDest.insertSheet(sheetName);
      Logger.log("Pestaña creada en destino: " + sheetName);
    } else {
      try {
        var lastDestRow = destSheet.getLastRow();
        var lastDestCol = destSheet.getLastColumn();
        if (lastDestRow > 1 && lastDestCol > 0) {
          destHadData = true;
          var destValues = destSheet.getRange(1, 1, lastDestRow, lastDestCol).getValues();
          var destHeaders = destValues[0];
          var destWIdIdx = 0; // Column A
          var destProdDateIdx = getColIndex(destHeaders, "Production Date");
          var destArrIdx = getColIndex(destHeaders, "ARR");
          var destProgressIdx = getColIndex(destHeaders, "Progress");
          var destComentariosIdx = getColIndex(destHeaders, "Comentarios");
          var destStatusIdx = getColIndex(destHeaders, "Status");
          
          if (destWIdIdx !== -1) {
            for (var r = 1; r < destValues.length; r++) {
              var row = destValues[r];
              var wId = row[destWIdIdx] ? row[destWIdIdx].toString().trim() : "";
              if (wId) {
                oldWorkloads[wId] = {
                  prodDate: destProdDateIdx !== -1 ? row[destProdDateIdx] : "",
                  arr: destArrIdx !== -1 ? row[destArrIdx] : "",
                  progress: destProgressIdx !== -1 ? row[destProgressIdx] : "",
                  comentarios: destComentariosIdx !== -1 ? row[destComentariosIdx] : "",
                  status: destStatusIdx !== -1 ? row[destStatusIdx] : ""
                };
              }
            }
          }
        }
      } catch (e) {
        Logger.log("Advertencia: No se pudo leer la pestaña destino anterior debido a un error de formato/tabla: " + e.message);
        destHadData = false;
      }
    }
    
    // Ensure Comentarios and Status headers exist in filteredValues
    var outHeaders = filteredValues[0];
    var outComentariosIdx = getColIndex(outHeaders, "Comentarios");
    if (outComentariosIdx === -1) {
      outComentariosIdx = outHeaders.length;
      outHeaders.push("Comentarios");
    }
    var outStatusIdx = getColIndex(outHeaders, "Status");
    if (outStatusIdx === -1) {
      outStatusIdx = outHeaders.length;
      outHeaders.push("Status");
    }
    
    var targetColCount = outHeaders.length;
    for (var r = 1; r < filteredValues.length; r++) {
      var row = filteredValues[r];
      while (row.length < targetColCount) {
        row.push("");
      }
      var wId = row[workloadIdIdx] ? row[workloadIdIdx].toString().trim() : "";
      if (wId && oldWorkloads.hasOwnProperty(wId)) {
        if (oldWorkloads[wId].comentarios !== undefined && oldWorkloads[wId].comentarios !== "") {
          row[outComentariosIdx] = oldWorkloads[wId].comentarios;
        }
        if (oldWorkloads[wId].status !== undefined && oldWorkloads[wId].status !== "") {
          row[outStatusIdx] = oldWorkloads[wId].status;
        }
      }
      filteredValues[r] = row;
    }
    
    try {
      destSheet.clear();
    } catch (e) {
      Logger.log("Error al limpiar destSheet con clear(): " + e.message + ". Recreando la pestaña...");
      var sheetPos = destSheet.getIndex();
      ssDest.deleteSheet(destSheet);
      destSheet = ssDest.insertSheet(sheetName, sheetPos - 1);
    }
    
    if (filteredValues.length > 1) {
      var numRows = filteredValues.length;
      var numCols = filteredValues[0].length;
      destSheet.getRange(1, 1, numRows, numCols).setValues(filteredValues);
      
      var backgrounds = [];
      var headerBg = [];
      for (var c = 0; c < numCols; c++) {
        headerBg.push(null);
      }
      backgrounds.push(headerBg);
      
      for (var r = 1; r < filteredValues.length; r++) {
        var row = filteredValues[r];
        var wId = workloadIdIdx !== -1 && row[workloadIdIdx] ? row[workloadIdIdx].toString().trim() : "";
        var rowBgColor = null;
        
        if (wId && workloadIdIdx !== -1) {
          var key = sheetName + "_" + wId;
          
          if (destHadData && !oldWorkloads.hasOwnProperty(wId)) {
            activeHighlights[key] = "NEW WORKLOAD";
            logEntries.push([timestamp, sheetName, wId, "NEW WORKLOAD", "", "", ""]);
          } else if (destHadData && oldWorkloads.hasOwnProperty(wId)) {
            var oldData = oldWorkloads[wId];
            var newProdDate = prodDateIdx !== -1 ? row[prodDateIdx] : "";
            var newArr = arrIdx !== -1 ? row[arrIdx] : "";
            var newProgress = progressIdx !== -1 ? row[progressIdx] : "";
            
            var changedFields = [];
            if (prodDateIdx !== -1 && !valuesEqual(oldData.prodDate, newProdDate)) {
              changedFields.push({ field: "Production Date", oldVal: oldData.prodDate, newVal: newProdDate });
            }
            if (arrIdx !== -1 && !valuesEqual(oldData.arr, newArr)) {
              changedFields.push({ field: "ARR", oldVal: oldData.arr, newVal: newArr });
            }
            if (progressIdx !== -1 && !valuesEqual(oldData.progress, newProgress)) {
              changedFields.push({ field: "Progress", oldVal: oldData.progress, newVal: newProgress });
            }
            
            if (changedFields.length > 0) {
              activeHighlights[key] = "MODIFIED";
              for (var cf = 0; cf < changedFields.length; cf++) {
                logEntries.push([
                  timestamp,
                  sheetName,
                  wId,
                  "MODIFIED",
                  changedFields[cf].field,
                  formatLogVal(changedFields[cf].oldVal),
                  formatLogVal(changedFields[cf].newVal)
                ]);
              }
            }
          }
          
          if (activeHighlights[key] === "NEW WORKLOAD") {
            rowBgColor = "#d9ead3"; // light green
          } else if (activeHighlights[key] === "MODIFIED") {
            rowBgColor = "#fff2cc"; // light yellow
          }
        }
        
        var rowBgs = [];
        for (var c = 0; c < numCols; c++) {
          rowBgs.push(rowBgColor);
        }
        backgrounds.push(rowBgs);
      }
      
      destSheet.getRange(1, 1, numRows, numCols).setBackgrounds(backgrounds);
      Logger.log("Pestaña sincronizada con éxito: " + sheetName);
    } else {
      destSheet.getRange(1, 1, 1, filteredValues[0].length).setValues([filteredValues[0]]);
      Logger.log("Pestaña origen vacía: " + sheetName);
    }
  }
  
  if (logEntries.length > 0) {
    var lastLogRow = logSheet.getLastRow();
    if (lastLogRow === 0) {
      logSheet.appendRow(["Timestamp", "Sheet Name", "Workload ID", "Change Type", "Field", "Old Value", "New Value"]);
      lastLogRow = 1;
    }
    logSheet.getRange(lastLogRow + 1, 1, logEntries.length, 7).setValues(logEntries);
  }
  
  Logger.log("¡Sincronización de pestañas principales completada!");
}
