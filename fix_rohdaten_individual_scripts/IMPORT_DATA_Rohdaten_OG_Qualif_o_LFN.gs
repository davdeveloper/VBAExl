function ogqlNormalizeText_(value) {
  return String(value == null ? "" : value).normalize("NFC").trim();
}

function ogqlFoldText_(value) {
  return ogqlNormalizeText_(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/ß/g, "ss").toLowerCase().trim();
}

function ogqlExtractDriveId_(value) {
  const text = ogqlNormalizeText_(value);
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

function ogqlFindSettingsSource_(settingsSheet, searchFileName) {
  const values = settingsSheet.getRange("A1:F" + settingsSheet.getLastRow()).getValues();
  const searchBase = ogqlFoldText_(searchFileName);
  const searchExcel = ogqlFoldText_(searchFileName + ".xlsx");

  for (let i = 0; i < values.length; i++) {
    const cellA = ogqlFoldText_(values[i][0]);
    const cellB = ogqlFoldText_(values[i][1]);

    if (cellA === searchBase || cellB === searchBase || cellA === searchExcel || cellB === searchExcel ||
        cellA.includes(searchBase) || cellB.includes(searchBase)) {
      const fileId = ogqlExtractDriveId_(values[i][4]) || ogqlExtractDriveId_(values[i][5]);
      if (!fileId) return null;

      const rowNumber = i + 1;
      settingsSheet.getRange(rowNumber, 5).setValue(fileId);
      settingsSheet.getRange(rowNumber, 6).setValue(`https://docs.google.com/spreadsheets/d/${fileId}/edit`);
      return { rowNumber, fileId };
    }
  }

  return null;
}

function ogqlCleanText_(value) {
  return ogqlNormalizeText_(value).replace(/"/g, "").replace(/\u00A0/g, " ").trim();
}

function ogqlCleanWgText_(value) {
  let text = ogqlCleanText_(value);
  while (text.length > 5 && text.endsWith("0")) text = text.slice(0, -1);
  return text;
}

function ogqlClearContentFromRow_(sheet, startRow, columnCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow - startRow + 1, columnCount).clearContent();
  }
}

/**
 * Button function: importRawOGQualityCasesNoLFN
 * Imports "8WS_O+G_DE_Qualitätsfälle (ohne LFN)" into "Rohdaten O+G Qualitätsf. o. LFN".
 */
function importRawOGQualityCasesNoLFN() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const targetSheet = ss.getSheetByName("Rohdaten O+G Qualitätsf. o. LFN");

  if (!settingsSheet || !targetSheet) {
    ui.alert("Fehler", 'Benötigte Tabellenblätter fehlen: "Einstellungen" oder "Rohdaten O+G Qualitätsf. o. LFN".', ui.ButtonSet.OK);
    return;
  }

  const searchFileName = "8WS_O+G_DE_Qualitätsfälle (ohne LFN)";
  const sourceInfo = ogqlFindSettingsSource_(settingsSheet, searchFileName);
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

  const rawValues = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 6).getDisplayValues();
  const sourceValues = rawValues
    .map(row => row.map((value, index) => index === 0 ? ogqlCleanWgText_(value) : ogqlCleanText_(value)))
    .filter(row => row.some(value => value !== ""));

  if (!sourceValues.length) {
    ui.alert("Keine Daten", "Nach Bereinigung wurden keine importierbaren Daten gefunden.", ui.ButtonSet.OK);
    return;
  }

  ogqlClearContentFromRow_(targetSheet, 8, 17);

  const destinationRange = targetSheet.getRange(8, 2, sourceValues.length, sourceValues[0].length);
  destinationRange.setNumberFormat("@");
  destinationRange.setValues(sourceValues);

  targetSheet.getRange("E2").activate();
  ss.toast(`${sourceValues.length} Zeilen importiert.`, "Rohdaten O+G Qualitätsf. o. LFN", 8);
}
