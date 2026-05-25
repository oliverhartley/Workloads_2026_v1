/**
 * Consolidation.js
 * Genera o actualiza la pestaña 'Consolidated Workloads' agrupando las 4 pestañas principales.
 * Alinea dinámicamente todas las columnas por nombre de encabezado para evitar desajustes.
 */

function createConsolidatedWorkloads() {
  var ssDest = SpreadsheetApp.getActiveSpreadsheet();
  if (!ssDest) {
    Logger.log("No se encontró una hoja de cálculo activa.");
    return;
  }
  
  var sheetsToSync = CONFIG.CORE_SHEETS_TO_SYNC;
  var consolidatedSheetName = "Consolidated Workloads";
  var allSheetsData = [];
  
  for (var i = 0; i < sheetsToSync.length; i++) {
    var sheetName = sheetsToSync[i];
    var sheet = ssDest.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log("La pestaña " + sheetName + " no existe en el documento actual.");
      continue;
    }
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) {
      continue;
    }
    
    var range = sheet.getRange(1, 1, lastRow, lastCol);
    var values = range.getValues();
    var backgrounds = range.getBackgrounds();
    
    allSheetsData.push({
      sheetName: sheetName,
      headers: values[0],
      rows: values.slice(1),
      backgrounds: backgrounds.slice(1)
    });
  }
  
  if (allSheetsData.length > 0) {
    var masterHeaders = [];
    for (var s = 0; s < allSheetsData.length; s++) {
      var sHeaders = allSheetsData[s].headers;
      for (var h = 0; h < sHeaders.length; h++) {
        var hName = sHeaders[h] !== null && sHeaders[h] !== undefined ? sHeaders[h].toString().trim() : "";
        if (hName && masterHeaders.indexOf(hName) === -1) {
          masterHeaders.push(hName);
        }
      }
    }
    masterHeaders.push("Commit Status");
    masterHeaders.push("Workload Link");
    
    var allConsolidatedRows = [masterHeaders];
    var allConsolidatedBackgrounds = [];
    var masterHeaderBg = [];
    for (var mh = 0; mh < masterHeaders.length; mh++) {
      masterHeaderBg.push(null);
    }
    allConsolidatedBackgrounds.push(masterHeaderBg);
    
    var commitStatusIdx = masterHeaders.indexOf("Commit Status");
    var workloadLinkIdx = masterHeaders.indexOf("Workload Link");
    
    // Identificar el índice de la columna del Socio/Partner en masterHeaders
    var partnerColIdx = -1;
    var partnerHeaderNames = ["partner", "partner name", "partner / name", "socio", "domain", "partner domain"];
    for (var mh = 0; mh < masterHeaders.length; mh++) {
      var hNameLower = masterHeaders[mh].toString().trim().toLowerCase();
      if (partnerHeaderNames.indexOf(hNameLower) !== -1) {
        partnerColIdx = mh;
        break;
      }
    }
    
    for (var s = 0; s < allSheetsData.length; s++) {
      var sData = allSheetsData[s];
      var sName = sData.sheetName;
      var sHeaders = sData.headers;
      
      var commitStatus = "";
      if (sName.indexOf("Uncommitted") !== -1) {
        commitStatus = "Uncommitted";
      } else if (sName.indexOf("Committed") !== -1) {
        commitStatus = "Committed";
      }
      
      for (var r = 0; r < sData.rows.length; r++) {
        var sRow = sData.rows[r];
        var wId = sRow[0] ? sRow[0].toString().trim() : "";
        if (!wId) {
          continue;
        }
        var sBg = sData.backgrounds[r][0]; // Row background
        
        var consolRow = [];
        var consolBg = [];
        for (var mh = 0; mh < masterHeaders.length; mh++) {
          consolRow.push("");
          consolBg.push(sBg);
        }
        
        for (var c = 0; c < sHeaders.length; c++) {
          var hName = sHeaders[c] !== null && sHeaders[c] !== undefined ? sHeaders[c].toString().trim() : "";
          if (hName) {
            var mIdx = masterHeaders.indexOf(hName);
            if (mIdx !== -1) {
              consolRow[mIdx] = sRow[c] !== null && sRow[c] !== undefined ? sRow[c] : "";
            }
          }
        }
        
        consolRow[commitStatusIdx] = commitStatus;
        
        var workloadLink = "https://vector.lightning.force.com/lightning/r/Workload__c/" + wId + "/view";
        consolRow[workloadLinkIdx] = workloadLink;
        
        // Filtrar si la columna de socio está vacía
        if (partnerColIdx !== -1) {
          var partnerVal = consolRow[partnerColIdx] ? consolRow[partnerColIdx].toString().trim() : "";
          if (!partnerVal) {
            continue; // Ignorar filas sin socio asignado
          }
        }
        
        allConsolidatedRows.push(consolRow);
        allConsolidatedBackgrounds.push(consolBg);
      }
    }
    
    var consolSheet = ssDest.getSheetByName(consolidatedSheetName);
    if (!consolSheet) {
      consolSheet = ssDest.insertSheet(consolidatedSheetName);
      Logger.log("Pestaña consolidada creada: " + consolidatedSheetName);
    }
    
    try {
      consolSheet.clear();
    } catch (e) {
      Logger.log("Error al limpiar consolSheet con clear(): " + e.message + ". Recreando la pestaña consolidada...");
      var consolPos = consolSheet.getIndex();
      ssDest.deleteSheet(consolSheet);
      consolSheet = ssDest.insertSheet(consolidatedSheetName, consolPos - 1);
    }
    
    var totalConsolRows = allConsolidatedRows.length;
    var totalConsolCols = masterHeaders.length;
    consolSheet.getRange(1, 1, totalConsolRows, totalConsolCols).setValues(allConsolidatedRows);
    consolSheet.getRange(1, 1, totalConsolRows, totalConsolCols).setBackgrounds(allConsolidatedBackgrounds);
    Logger.log("Pestaña consolidada actualizada con éxito.");
  } else {
    Logger.log("No hay datos en las pestañas individuales para consolidar.");
  }
}
