function rfqNormalizeText_(value) {
  return String(value == null ? "" : value).normalize("NFC").trim();
}

function rfqFoldText_(value) {
  return rfqNormalizeText_(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/ß/g, "ss").toLowerCase().trim();
}

function rfqExtractDriveId_(value) {
  const text = rfqNormalizeText_(value);
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

function rfqFindSettingsSource_(settingsSheet, searchFileName) {
  const values = settingsSheet.getRange("A1:F" + settingsSheet.getLastRow()).getValues();
  const searchBase = rfqFoldText_(searchFileName);
  const searchExcel = rfqFoldText_(searchFileName + ".xlsx");

  for (let i = 0; i < values.length; i++) {
    const cellA = rfqFoldText_(values[i][0]);
    const cellB = rfqFoldText_(values[i][1]);

    if (cellA === searchBase || cellB === searchBase || cellA === searchExcel || cellB === searchExcel ||
        cellA.includes(searchBase) || cellB.includes(searchBase)) {
      const fileId = rfqExtractDriveId_(values[i][4]) || rfqExtractDriveId_(values[i][5]);
      if (!fileId) return null;

      const rowNumber = i + 1;
      settingsSheet.getRange(rowNumber, 5).setValue(fileId);
      settingsSheet.getRange(rowNumber, 6).setValue(`https://docs.google.com/spreadsheets/d/${fileId}/edit`);
      return { rowNumber, fileId };
    }
  }

  return null;
}

function rfqCleanText_(value) {
  return rfqNormalizeText_(value).replace(/"/g, "").replace(/\u00A0/g, " ").trim();
}

function rfqCleanWgText_(value) {
  let text = rfqCleanText_(value);
  while (text.length > 5 && text.endsWith("0")) text = text.slice(0, -1);
  return text;
}

function rfqFindHeaderColumn_(sheet, headerText, headerRows, maxColumns) {
  const headers = sheet.getRange(1, 1, headerRows, maxColumns).getDisplayValues();
  const wanted = rfqFoldText_(headerText);

  for (let r = 0; r < headers.length; r++) {
    for (let c = 0; c < headers[r].length; c++) {
      if (rfqFoldText_(headers[r][c]) === wanted) return c;
    }
  }

  return -1;
}

function rfqClearContentFromRow_(sheet, startRow, columnCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow - startRow + 1, columnCount).clearContent();
  }
}

/**
 * Button function: importRawFoodQualityCases
 * Imports "8WS_Food_DE_Qualitätsfälle" into "Rohdaten Food Qualitätsfälle".
 */
function importRawFoodQualityCases() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const targetSheet = ss.getSheetByName("Rohdaten Food Qualitätsfälle");

  if (!settingsSheet || !targetSheet) {
    ui.alert("Fehler", 'Benötigte Tabellenblätter fehlen: "Einstellungen" oder "Rohdaten Food Qualitätsfälle".', ui.ButtonSet.OK);
    return;
  }

  const searchFileName = "8WS_Food_DE_Qualitätsfälle";
  const sourceInfo = rfqFindSettingsSource_(settingsSheet, searchFileName);
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

  const sourceDataSheet = sourceSs.getSheetByName("In_Development_8_weeks_Report_F") || sourceSs.getSheets()[0];
  const sourceLastRow = sourceDataSheet.getLastRow();
  if (sourceLastRow < 3) {
    ui.alert("Keine Daten", "Die Quelldatei enthält ab Zeile 3 keine Daten.", ui.ButtonSet.OK);
    return;
  }

  let wgColumnIndex = rfqFindHeaderColumn_(sourceDataSheet, "WG", 3, 10);
  if (wgColumnIndex < 0 || wgColumnIndex > 7) wgColumnIndex = 0;

  const rawValues = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 8).getDisplayValues();
  const sourceValues = rawValues
    .map(row => row.map((value, index) => index === wgColumnIndex ? rfqCleanWgText_(value) : rfqCleanText_(value)))
    .filter(row => row.some(value => value !== ""));

  if (!sourceValues.length) {
    ui.alert("Keine Daten", "Nach Bereinigung wurden keine importierbaren Daten gefunden.", ui.ButtonSet.OK);
    return;
  }

  rfqClearContentFromRow_(targetSheet, 8, 25);

  const destinationRange = targetSheet.getRange(8, 7, sourceValues.length, sourceValues[0].length);
  destinationRange.setNumberFormat("@");
  destinationRange.setValues(sourceValues);

  targetSheet.getRange("E2").activate();
  ss.toast(`${sourceValues.length} Zeilen importiert.`, "Rohdaten Food Qualitätsfälle", 8);
}
