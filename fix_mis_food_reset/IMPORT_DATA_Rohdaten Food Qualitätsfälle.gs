/**
 * Shared helpers for the migrated Apps Script import modules.
 * Keep these helpers in the project if other split all2 scripts use all2... functions.
 */
function all2NormalizeText_(value) {
  return String(value == null ? "" : value).normalize("NFC").trim();
}

function all2FindSettingsRow_(settingsSheet, searchFileName) {
  const settingsValues = settingsSheet.getRange("A1:F" + settingsSheet.getLastRow()).getValues();
  const searchBase = all2NormalizeText_(searchFileName);
  const searchExcel = all2NormalizeText_(searchFileName + ".xlsx");
  const options = { sensitivity: "base" };

  for (let i = 0; i < settingsValues.length; i++) {
    const cellA = all2NormalizeText_(settingsValues[i][0]);
    const cellB = all2NormalizeText_(settingsValues[i][1]);

    if ((cellA && cellA.localeCompare(searchBase, undefined, options) === 0) ||
        (cellB && cellB.localeCompare(searchExcel, undefined, options) === 0)) {
      return {
        rowNumber: i + 1,
        fileId: all2ExtractDriveId_(settingsValues[i][4]),
        fileUrl: all2NormalizeText_(settingsValues[i][5])
      };
    }
  }

  return null;
}

function all2ExtractDriveId_(value) {
  const text = all2NormalizeText_(value);
  if (!text) return "";

  const patterns = [
    /\/d\/([a-zA-Z0-9_-]+)/,
    /\/folders\/([a-zA-Z0-9_-]+)/,
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

function all2SanitizeDisplayValues_(values) {
  return values.map(row => row.map(value => {
    const text = all2NormalizeText_(value).replace(/"/g, "").replace(/\u00A0/g, " ").trim();
    return text === "" ? "" : text;
  }));
}

function all2ClearContentFromRow_(sheet, startRow, columnCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow - startRow + 1, columnCount).clearContent();
  }
}

function all2CleanUpTempSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  for (const sheet of sheets) {
    if (sheet.getName().toLowerCase().startsWith("temp_") && ss.getSheets().length > 1) {
      try {
        ss.deleteSheet(sheet);
      } catch (e) {
        console.error(`Konnte ${sheet.getName()} nicht löschen: ${e.message}`);
      }
    }
  }
}

function all2FindRowContainingAny_(sheet, searchTerms) {
  const values = sheet.getDataRange().getValues();
  const normalizedTerms = searchTerms.map(term => all2NormalizeText_(term).toLowerCase());

  for (let i = 0; i < values.length; i++) {
    const rowText = values[i].map(cell => all2NormalizeText_(cell).toLowerCase()).join(" ");
    if (normalizedTerms.some(term => term && rowText.includes(term))) {
      return i + 1;
    }
  }

  return -1;
}

function all2FindHeaderColumn_(sheet, headerText, headerRows, maxColumns) {
  const headers = sheet.getRange(1, 1, headerRows, maxColumns).getDisplayValues();
  const wanted = all2NormalizeText_(headerText).toUpperCase();

  for (let r = 0; r < headers.length; r++) {
    for (let c = 0; c < headers[r].length; c++) {
      if (all2NormalizeText_(headers[r][c]).toUpperCase() === wanted) {
        return c;
      }
    }
  }

  return -1;
}

function all2CleanWgText_(value) {
  let text = all2NormalizeText_(value).replace(/"/g, "").replace(/\u00A0/g, " ").trim();

  while (text.length > 5 && text.endsWith("0")) {
    text = text.slice(0, -1);
  }

  return text;
}

/**
 * Imports raw data for "Rohdaten Food Qualitätsfälle".
 *
 * Restores the old V4 business logic:
 * - Scans rows 1-3 for the WG column.
 * - Sanitizes the WG column without changing the source file.
 *
 * Keeps the newer safety fix:
 * - Reads display values and writes target values as plain text only.
 *
 * Button function: importRawFoodQualityCases
 */
function importRawFoodQualityCases() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const targetSheet = ss.getSheetByName("Rohdaten Food Qualitätsfälle");

  Logger.log("--- Starting importRawFoodQualityCases (V4 WG logic + target-only plain text) ---");

  if (!settingsSheet || !targetSheet) {
    ui.alert('Fehler', 'Benötigte Tabellenblätter fehlen: "Einstellungen" oder "Rohdaten Food Qualitätsfälle".', ui.ButtonSet.OK);
    return;
  }

  const searchFileName = "8WS_Food_DE_Qualitätsfälle";
  const settingMatch = all2FindSettingsRow_(settingsSheet, searchFileName);
  const sourceFileId = settingMatch ? settingMatch.fileId : "";

  if (!sourceFileId) {
    ui.alert('Konfigurationsfehler', `File ID für "${searchFileName}" wurde nicht gefunden.`, ui.ButtonSet.OK);
    return;
  }

  all2ClearContentFromRow_(targetSheet, 8, 25);

  let sourceSs;
  try {
    sourceSs = SpreadsheetApp.openById(sourceFileId);
  } catch (e) {
    ui.alert('Dateifehler', `Quelldatei konnte nicht geöffnet werden.\n\n${e.message}`, ui.ButtonSet.OK);
    return;
  }

  const sourceDataSheet = sourceSs.getSheetByName("In_Development_8_weeks_Report_F");
  if (!sourceDataSheet) {
    ui.alert('Tabellenblattfehler', 'Tab "In_Development_8_weeks_Report_F" wurde in der Quelldatei nicht gefunden.', ui.ButtonSet.OK);
    return;
  }

  const sourceLastRow = sourceDataSheet.getLastRow();
  if (sourceLastRow < 3) {
    ss.toast("Keine Rohdaten für Food Qualitätsfälle gefunden.", "Hinweis", 5);
    return;
  }

  let wgColIndex = all2FindHeaderColumn_(sourceDataSheet, "WG", 3, 10);
  if (wgColIndex < 0 || wgColIndex > 7) {
    Logger.log(`WG-Spalte nicht in A:H gefunden, nutze Fallback Spalte A. Gefunden: ${wgColIndex}`);
    wgColIndex = 0;
  }

  const sourceRange = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 8);
  const sourceValues = sourceRange.getDisplayValues().map(row => {
    return row.map((value, index) => {
      if (index === wgColIndex) return all2CleanWgText_(value);

      const text = all2NormalizeText_(value).replace(/"/g, "").replace(/\u00A0/g, " ").trim();
      return text === "" ? "" : text;
    });
  });

  if (!sourceValues.length || !sourceValues[0].length) {
    ss.toast("Keine importierbaren Rohdaten für Food Qualitätsfälle gefunden.", "Hinweis", 5);
    return;
  }

  const destinationRange = targetSheet.getRange(8, 7, sourceValues.length, sourceValues[0].length);
  destinationRange.setNumberFormat("@");
  destinationRange.setValues(sourceValues);

  targetSheet.getRange("E2").activate();
  ss.toast('Rohdaten Food Qualitätsfälle erfolgreich importiert.', 'Erfolg', 5);
}
