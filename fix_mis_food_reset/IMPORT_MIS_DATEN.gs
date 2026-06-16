/**
 * MIS import with text-safe article/key handling.
 *
 * Fixes the Excel-to-Sheets migration issue where displayed values such as
 * 190.164 were imported as raw numbers such as 190.163.788.
 */
function misNormalizeText_(value) {
  return String(value == null ? "" : value).normalize("NFC").trim();
}

function misExtractDriveId_(value) {
  const text = misNormalizeText_(value);
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

function misSanitizeDisplayValues_(values) {
  return values.map(row => row.map(value => {
    const text = misNormalizeText_(value).replace(/"/g, "").replace(/\u00A0/g, " ").trim();
    return text === "" ? "" : text;
  }));
}

function misClearContentFromRow_(sheet, startRow, columnCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow - startRow + 1, columnCount).clearContent();
  }
}

/**
 * Imports MIS data and updates the lookup formula in "MIS Final".
 *
 * Button function: runMisImportSmarter
 */
function runMisImportSmarter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const ui = SpreadsheetApp.getUi();

  if (!settingsSheet) {
    ui.alert('Fehler', 'Tabellenblatt "Einstellungen" wurde nicht gefunden.', ui.ButtonSet.OK);
    return;
  }

  const values = settingsSheet.getRange("A1:F" + settingsSheet.getLastRow()).getValues();
  const searchTextData = "WE-Menge pro Artikel-Lieferant";
  const searchTextFormula = "Abverkauf bereinigt pro Artikel";

  let misFileId = "";
  let formulaRowIndex = -1;

  for (let i = 0; i < values.length; i++) {
    const cellA = misNormalizeText_(values[i][0]);
    const cellB = misNormalizeText_(values[i][1]);
    const rowText = `${cellA} ${cellB}`;

    if (rowText.includes(searchTextData)) {
      misFileId = misExtractDriveId_(values[i][4]);
      Logger.log(`MIS Quelldatei gefunden in Zeile ${i + 1}: ${misFileId}`);
    }

    if (rowText.includes(searchTextFormula)) {
      formulaRowIndex = i + 1;
      Logger.log(`MIS Formelquelle gefunden in Zeile ${formulaRowIndex}`);
    }
  }

  if (!misFileId || formulaRowIndex < 1) {
    ui.alert(
      'Konfigurationsfehler',
      'In "Einstellungen" wurden nicht beide MIS-Zeilen gefunden: "WE-Menge pro Artikel-Lieferant" und "Abverkauf bereinigt pro Artikel".',
      ui.ButtonSet.OK
    );
    return;
  }

  const misFinalSheet = ss.getSheetByName("MIS Final");
  if (!misFinalSheet) {
    ui.alert('Fehler', 'Tabellenblatt "MIS Final" wurde nicht gefunden.', ui.ButtonSet.OK);
    return;
  }

  const targetTabName = "8 W S DE Food - Abverkauf berei";

  // Textvergleich statt A4*1/IMPORTRANGE*1: Punkte in Artikelnummern/Werten bleiben unverändert.
  const newFormula = `=XLOOKUP(TO_TEXT(A4); ARRAYFORMULA(TO_TEXT(IMPORTRANGE(Einstellungen!$F$${formulaRowIndex}; "'${targetTabName}'!A:A"))); IMPORTRANGE(Einstellungen!$F$${formulaRowIndex}; "'${targetTabName}'!C:C"); 0; 0; 1)`;

  try {
    misFinalSheet.getRange("F4").setFormula(newFormula);
  } catch (e) {
    ui.alert('Formelfehler', `Die Formel in "MIS Final"!F4 konnte nicht gesetzt werden.\n\n${e.message}`, ui.ButtonSet.OK);
    return;
  }

  importAndProcessMisData(misFileId);

  if (typeof updateDeckblattFoodFormulas === "function") updateDeckblattFoodFormulas();
  if (typeof updateDeckblattNonFoodFormulas === "function") updateDeckblattNonFoodFormulas();
  if (typeof updateDeckblatt_O_G_Formulas === "function") updateDeckblatt_O_G_Formulas();
}

/**
 * Imports MIS source data into "MIS Final" as displayed text.
 *
 * @param {string} sourceFileId Google Drive ID of the converted source Google Sheet.
 */
function importAndProcessMisData(sourceFileId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("MIS Final");
  const ui = SpreadsheetApp.getUi();

  if (!targetSheet) {
    ui.alert('Fehler', 'Tabellenblatt "MIS Final" wurde nicht gefunden.', ui.ButtonSet.OK);
    return;
  }

  let sourceSs;
  try {
    sourceSs = SpreadsheetApp.openById(misExtractDriveId_(sourceFileId));
  } catch (e) {
    ui.alert('Dateifehler', `Die MIS-Quelldatei konnte nicht geöffnet werden.\n\n${e.message}`, ui.ButtonSet.OK);
    return;
  }

  const sourceSheet = sourceSs.getSheets()[0];
  const sourceLastRow = sourceSheet.getLastRow();

  if (sourceLastRow < 2) {
    ui.alert('Keine Daten', 'Die MIS-Quelldatei enthält keine importierbaren Daten.', ui.ButtonSet.OK);
    return;
  }

  const sourceRange = sourceSheet.getRange(2, 1, sourceLastRow - 1, 4);
  const sourceValues = misSanitizeDisplayValues_(sourceRange.getDisplayValues());

  if (!sourceValues.length || !sourceValues[0].length) {
    ui.alert('Keine Daten', 'Die MIS-Quelldatei enthält keine importierbaren Daten.', ui.ButtonSet.OK);
    return;
  }

  misClearContentFromRow_(targetSheet, 6, 7);

  const targetDataRange = targetSheet.getRange(6, 1, sourceValues.length, sourceValues[0].length);
  targetDataRange.setNumberFormat("@");
  targetDataRange.setValues(sourceValues);

  const newLastRow = 5 + sourceValues.length;

  if (newLastRow >= 6) {
    const formulaTemplateRange = targetSheet.getRange("E4:G4");
    const formulaDestinationRange = targetSheet.getRange(6, 5, newLastRow - 5, 3);

    formulaTemplateRange.copyTo(formulaDestinationRange);
    SpreadsheetApp.flush();

    const calculatedValues = formulaDestinationRange.getDisplayValues();
    formulaDestinationRange.setNumberFormat("@");
    formulaDestinationRange.setValues(calculatedValues);
    formulaDestinationRange.setFontColor("black");
    formulaDestinationRange.setBackground("white");
  }

  targetSheet.getRange("D1").activate();

  if (typeof navigateOnSuccess === "function") {
    navigateOnSuccess();
  } else {
    ss.toast("MIS Import erfolgreich abgeschlossen.", "Erfolg", 5);
  }
}
