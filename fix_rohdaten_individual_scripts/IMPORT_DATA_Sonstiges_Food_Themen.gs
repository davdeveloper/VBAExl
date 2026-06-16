function sftNormalizeText_(value) {
  return String(value == null ? "" : value).normalize("NFC").trim();
}

function sftFoldText_(value) {
  return sftNormalizeText_(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/ß/g, "ss").toLowerCase().trim();
}

function sftExtractDriveId_(value) {
  const text = sftNormalizeText_(value);
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

function sftFindSettingsSource_(settingsSheet, searchFileName) {
  const values = settingsSheet.getRange("A1:F" + settingsSheet.getLastRow()).getValues();
  const searchBase = sftFoldText_(searchFileName);
  const searchExcel = sftFoldText_(searchFileName + ".xlsx");

  for (let i = 0; i < values.length; i++) {
    const cellA = sftFoldText_(values[i][0]);
    const cellB = sftFoldText_(values[i][1]);

    if (cellA === searchBase || cellB === searchBase || cellA === searchExcel || cellB === searchExcel ||
        cellA.includes(searchBase) || cellB.includes(searchBase)) {
      const fileId = sftExtractDriveId_(values[i][4]) || sftExtractDriveId_(values[i][5]);
      if (!fileId) return null;

      const rowNumber = i + 1;
      settingsSheet.getRange(rowNumber, 5).setValue(fileId);
      settingsSheet.getRange(rowNumber, 6).setValue(`https://docs.google.com/spreadsheets/d/${fileId}/edit`);
      return { rowNumber, fileId };
    }
  }

  return null;
}

function sftCleanText_(value) {
  return sftNormalizeText_(value).replace(/"/g, "").replace(/\u00A0/g, " ").trim();
}

function sftCleanWgText_(value) {
  let text = sftCleanText_(value);
  while (text.length > 7 && text.endsWith("0")) text = text.slice(0, -1);
  return text;
}

function sftFindHeaderColumn_(sheet, headerText, headerRows, maxColumns) {
  const headers = sheet.getRange(1, 1, headerRows, maxColumns).getDisplayValues();
  const wanted = sftFoldText_(headerText);

  for (let r = 0; r < headers.length; r++) {
    for (let c = 0; c < headers[r].length; c++) {
      if (sftFoldText_(headers[r][c]) === wanted) return c;
    }
  }

  return -1;
}

function sftClearContentFromRow_(sheet, startRow, columnCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow - startRow + 1, columnCount).clearContent();
  }
}

/**
 * Button function: importRawOtherFoodTopics
 * Imports "8WS_Food_DE_Sonstiges und aktuelle Themen" into "Sonstiges Food u Akt. Themen".
 */
function importRawOtherFoodTopics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const targetSheet = ss.getSheetByName("Sonstiges Food u Akt. Themen");

  if (!settingsSheet || !targetSheet) {
    ui.alert("Fehler", 'Benötigte Tabellenblätter fehlen: "Einstellungen" oder "Sonstiges Food u Akt. Themen".', ui.ButtonSet.OK);
    return;
  }

  const searchFileName = "8WS_Food_DE_Sonstiges und aktuelle Themen";
  const sourceInfo = sftFindSettingsSource_(settingsSheet, searchFileName);
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

  let wgColumnIndex = sftFindHeaderColumn_(sourceDataSheet, "WG", 3, 10);
  if (wgColumnIndex !== 0) wgColumnIndex = 0;

  const sourceValues = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 1)
    .getDisplayValues()
    .map(row => [wgColumnIndex === 0 ? sftCleanWgText_(row[0]) : sftCleanText_(row[0])])
    .filter(row => row[0] !== "");

  if (!sourceValues.length) {
    ui.alert("Keine Daten", "Nach Bereinigung wurden keine importierbaren Daten gefunden.", ui.ButtonSet.OK);
    return;
  }

  sftClearContentFromRow_(targetSheet, 7, 18);

  const destinationRange = targetSheet.getRange(7, 5, sourceValues.length, 1);
  destinationRange.setNumberFormat("@");
  destinationRange.setValues(sourceValues);

  targetSheet.getRange("E2").activate();
  ss.toast(`${sourceValues.length} Zeilen importiert. Bitte Spalten P & Q prüfen.`, "Sonstiges Food", 10);
}
