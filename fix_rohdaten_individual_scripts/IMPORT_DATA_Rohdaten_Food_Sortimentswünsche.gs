function rfsNormalizeText_(value) {
  return String(value == null ? "" : value).normalize("NFC").trim();
}

function rfsFoldText_(value) {
  return rfsNormalizeText_(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/ß/g, "ss").toLowerCase().trim();
}

function rfsExtractDriveId_(value) {
  const text = rfsNormalizeText_(value);
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

function rfsFindSettingsSource_(settingsSheet, searchFileName) {
  const values = settingsSheet.getRange("A1:F" + settingsSheet.getLastRow()).getValues();
  const searchBase = rfsFoldText_(searchFileName);
  const searchExcel = rfsFoldText_(searchFileName + ".xlsx");

  for (let i = 0; i < values.length; i++) {
    const cellA = rfsFoldText_(values[i][0]);
    const cellB = rfsFoldText_(values[i][1]);

    if (cellA === searchBase || cellB === searchBase || cellA === searchExcel || cellB === searchExcel ||
        cellA.includes(searchBase) || cellB.includes(searchBase)) {
      const fileId = rfsExtractDriveId_(values[i][4]) || rfsExtractDriveId_(values[i][5]);
      if (!fileId) return null;

      const rowNumber = i + 1;
      settingsSheet.getRange(rowNumber, 5).setValue(fileId);
      settingsSheet.getRange(rowNumber, 6).setValue(`https://docs.google.com/spreadsheets/d/${fileId}/edit`);
      return { rowNumber, fileId };
    }
  }

  return null;
}

function rfsCleanText_(value) {
  return rfsNormalizeText_(value).replace(/"/g, "").replace(/\u00A0/g, " ").trim();
}

function rfsCleanWgText_(value) {
  let text = rfsCleanText_(value);
  while (text.length > 5 && text.endsWith("0")) text = text.slice(0, -1);
  return text;
}

function rfsClearContentFromRow_(sheet, startRow, columnCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow - startRow + 1, columnCount).clearContent();
  }
}

/**
 * Button function: importFoodAssortmentWishes
 * Imports "8WS_Food_DE_Sortimentswünsche" into "Rohdaten Food Sortimentswünsche".
 */
function importFoodAssortmentWishes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const targetSheet = ss.getSheetByName("Rohdaten Food Sortimentswünsche");

  if (!settingsSheet || !targetSheet) {
    ui.alert("Fehler", 'Benötigte Tabellenblätter fehlen: "Einstellungen" oder "Rohdaten Food Sortimentswünsche".', ui.ButtonSet.OK);
    return;
  }

  const searchFileName = "8WS_Food_DE_Sortimentswünsche";
  const sourceInfo = rfsFindSettingsSource_(settingsSheet, searchFileName);
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

  const rawValues = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 4).getDisplayValues();
  const sourceValues = rawValues
    .map(row => row.map((value, index) => index === 0 ? rfsCleanWgText_(value) : rfsCleanText_(value)))
    .filter(row => row.some(value => value !== ""));

  if (!sourceValues.length) {
    ui.alert("Keine Daten", "Nach Bereinigung wurden keine importierbaren Daten gefunden.", ui.ButtonSet.OK);
    return;
  }

  rfsClearContentFromRow_(targetSheet, 8, 18);

  const destinationRange = targetSheet.getRange(8, 6, sourceValues.length, sourceValues[0].length);
  destinationRange.setNumberFormat("@");
  destinationRange.setValues(sourceValues);

  updateRohdatenFoodFormulas(sourceInfo.rowNumber);

  targetSheet.getRange("E2").activate();
  ss.toast(`${sourceValues.length} Zeilen importiert.`, "Rohdaten Food Sortimentswünsche", 8);
}

function updateRohdatenFoodFormulas(foundRowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("Rohdaten Food Sortimentswünsche");
  if (!targetSheet || !foundRowNumber) return;

  const newFormulaForJ6 = `=ARRAY_CONSTRAIN(ARRAYFORMULA(IFERROR(IF(J$5=J$4;INDEX(IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$3:$Q$1000");MATCH(1;($G6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$B$3:$B$1000"))*($H6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$C$3:$C$1000"));0);MATCH(J$5;IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$1:$Q$1");0));IFERROR(NUMBERVALUE(INDEX(IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$3:$Q$1000");MATCH(1;($G6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$B$3:$B$1000"))*($H6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$C$3:$C$1000"));0);MATCH(J$5;IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$1:$Q$1");0)));0)+IFERROR(NUMBERVALUE(INDEX(IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$3:$Q$1000");MATCH(1;($G6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$B$3:$B$1000"))*($H6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$C$3:$C$1000"));0);MATCH(J$4;IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$1:$Q$1");0)));0));"")); 1; 1)`;

  const templateFormulaCell = targetSheet.getRange("J6");
  templateFormulaCell.setFormula(newFormulaForJ6);
  templateFormulaCell.autoFill(targetSheet.getRange("J6:Q6"), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}
