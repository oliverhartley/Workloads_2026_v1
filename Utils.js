/**
 * Utils.gs
 * Funciones auxiliares compartidas para la búsqueda de columnas y comparación de valores.
 */

function getColIndex(headerList, name) {
  for (var c = 0; c < headerList.length; c++) {
    if (headerList[c] && headerList[c].toString().trim().toLowerCase() === name.toLowerCase()) {
      return c;
    }
  }
  return -1;
}

function valuesEqual(val1, val2) {
  if (val1 === null || val1 === undefined) val1 = "";
  if (val2 === null || val2 === undefined) val2 = "";
  
  if (val1 instanceof Date && val2 instanceof Date) {
    return val1.getTime() === val2.getTime();
  }
  return val1.toString().trim() === val2.toString().trim();
}

function formatLogVal(val) {
  if (val instanceof Date) {
    try {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } catch (e) {
      return val.toString();
    }
  }
  return val !== null && val !== undefined ? val.toString() : "";
}
