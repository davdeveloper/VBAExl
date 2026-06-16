function ogoNormalizeText_(value) {
  return String(value == null ? "" : value).normalize("NFC").trim();
}

function ogoFoldText_(value) {
  return ogoNormalizeText_(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/ß/g, "ss").toLowerCase().trim();
}

function ogoExtractDriveId_(value) {
  const text = ogoNormalizeText_(value);
  if (!text) return "";

  const patterns = [
    /\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/,
    /^([a-zA-Z0-9_-]{20,})$/,
    /([a-zA-Z0-9_-]{25,})/
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match && match[1]) return match[1];
  }

  return "";
}

function ogoFindSettingsSource_(settingsSheet, searchFileName) {
  const values = settingsSheet.getRange("A1:F" + settingsSheet.getLastRow()).getValues();
  const searchBase = ogoFoldText_(searchFileName);
  const searchExcel = ogoFoldText_(searchFileName + ".xlsx");

  for (let i = 0; i < values.length; i++) {
    const cellA = ogoFoldText_(values[i][0]);
    const cellB = ogoFoldText_(values[i][1]);

    if (cellA === searchBase || cellB === searchBase || cellA === searchExcel || cellB === searchExcel ||
        cellA.includes(searchBase) || cellB.includes(searchBase)) {
      const fileId = ogoExtractDriveId_(values[i][4]) || ogoExtractDriveId_(values[i][5]);
      if (!fileId) return null;

      const rowNumber = i + 1;
      settingsSheet.getRange(rowNumber, 5).setValue(fileId);
      settingsSheet.getRange(rowNumber, 6).setValue(`https://docs.google.com/spreadsheets/d/${fileId}/edit`);
      return { rowNumber, fileId };
    }
  }

  return null;
}

function ogoCleanWgText_(value) {
  let text = ogoNormalizeText_(value).replace(/"/g, "").replace(/\u00A0/g, " ").trim();
  while (text.length > 5 && text.endsWith("0")) text = text.slice(0, -1);
  return text;
}

function ogoClearContentFromRow_(sheet, startRow, columnCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow - startRow + 1, columnCount).clearContent();
  }
}

/**
 * Button function: importOGOther
 * Imports "8WS_O+G_DE_Sonstiges" into "O+G Sonstiges".
 */
function importOGOther() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const targetSheet = ss.getSheetByName("O+G Sonstiges");

  if (!settingsSheet || !targetSheet) {
    ui.alert("Fehler", 'Benötigte Tabellenblätter fehlen: "Einstellungen" oder "O+G Sonstiges".', ui.ButtonSet.OK);
    return;
  }

  const searchFileName = "8WS_O+G_DE_Sonstiges";
  const sourceInfo = ogoFindSettingsSource_(settingsSheet, searchFileName);
  if (!sourceInfo) {
    ui.alert("Konfigurationsfehler", `Quelle "${searchFileName}" wurde in Einstellungen nicht gefunden oder enthält keine gültige ID/URL.`, ui.ButtonSet.OK);
    return;
  }

  let sourceSs;
  try {
    sourceSs = SpreadsheetApp.openById(sourceInfo.fileId);
  } catch (e) {
    ui.alert("Dateifehler", `Quelldatei konnte nicht geöffnet werden.\n\n${e.message}`, ui.ButtonSet.OK);
    return;
  }

  const sourceDataSheet = sourceSs.getSheets()[0];
  const sourceLastRow = sourceDataSheet.getLastRow();
  if (sourceLastRow < 3) {
    ui.alert("Keine Daten", "Die Quelldatei enthält ab Zeile 3 keine Daten.", ui.ButtonSet.OK);
    return;
  }

  const sourceValues = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 1)
    .getDisplayValues()
    .map(row => [ogoCleanWgText_(row[0])])
    .filter(row => row[0] !== "");

  if (!sourceValues.length) {
    ui.alert("Keine Daten", "Nach Bereinigung wurden keine importierbaren Daten gefunden.", ui.ButtonSet.OK);
    return;
  }

  ogoClearContentFromRow_(targetSheet, 7, 11);

  const destinationRange = targetSheet.getRange(7, 1, sourceValues.length, 1);
  destinationRange.setNumberFormat("@");
  destinationRange.setValues(sourceValues);

  targetSheet.getRange("A1").activate();
  ss.toast(`${sourceValues.length} Zeilen importiert.`, "O+G Sonstiges", 8);
}
