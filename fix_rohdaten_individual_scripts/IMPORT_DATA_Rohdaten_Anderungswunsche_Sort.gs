function rcasNormalizeText_(value) {
  return String(value == null ? "" : value).normalize("NFC").trim();
}

function rcasFoldText_(value) {
  return rcasNormalizeText_(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/ß/g, "ss").toLowerCase().trim();
}

function rcasExtractDriveId_(value) {
  const text = rcasNormalizeText_(value);
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

function rcasFindSettingsSource_(settingsSheet, searchFileName) {
  const values = settingsSheet.getRange("A1:F" + settingsSheet.getLastRow()).getValues();
  const searchBase = rcasFoldText_(searchFileName);
  const searchExcel = rcasFoldText_(searchFileName + ".xlsx");

  for (let i = 0; i < values.length; i++) {
    const cellA = rcasFoldText_(values[i][0]);
    const cellB = rcasFoldText_(values[i][1]);

    if (cellA === searchBase || cellB === searchBase || cellA === searchExcel || cellB === searchExcel ||
        cellA.includes(searchBase) || cellB.includes(searchBase)) {
      const fileId = rcasExtractDriveId_(values[i][4]) || rcasExtractDriveId_(values[i][5]);
      if (!fileId) return null;

      const rowNumber = i + 1;
      settingsSheet.getRange(rowNumber, 5).setValue(fileId);
      settingsSheet.getRange(rowNumber, 6).setValue(`https://docs.google.com/spreadsheets/d/${fileId}/edit`);
      return { rowNumber, fileId };
    }
  }

  return null;
}

function rcasCleanText_(value) {
  return rcasNormalizeText_(value).replace(/"/g, "").replace(/\u00A0/g, " ").trim();
}

function rcasCleanWgText_(value) {
  let text = rcasCleanText_(value);
  while (text.length > 5 && text.endsWith("0")) text = text.slice(0, -1);
  return text;
}

function rcasClearContentFromRow_(sheet, startRow, columnCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow - startRow + 1, columnCount).clearContent();
  }
}

/**
 * Button function: importChangeWishesSorting
 * Imports "8WS_Food_DE_Änderungswünsche Sortierung" into "Rohdaten Änderungswünsche Sort.".
 */
function importChangeWishesSorting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const targetSheet = ss.getSheetByName("Rohdaten Änderungswünsche Sort.");

  if (!settingsSheet || !targetSheet) {
    ui.alert("Fehler", 'Benötigte Tabellenblätter fehlen: "Einstellungen" oder "Rohdaten Änderungswünsche Sort.".', ui.ButtonSet.OK);
    return;
  }

  const searchFileName = "8WS_Food_DE_Änderungswünsche Sortierung";
  const sourceInfo = rcasFindSettingsSource_(settingsSheet, searchFileName);
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
    .map(row => row.map((value, index) => index === 0 ? rcasCleanWgText_(value) : rcasCleanText_(value)))
    .filter(row => row.some(value => value !== ""));

  if (!sourceValues.length) {
    ui.alert("Keine Daten", "Nach Bereinigung wurden keine importierbaren Daten gefunden.", ui.ButtonSet.OK);
    return;
  }

  rcasClearContentFromRow_(targetSheet, 8, 18);

  const destinationRange = targetSheet.getRange(8, 6, sourceValues.length, sourceValues[0].length);
  destinationRange.setNumberFormat("@");
  destinationRange.setValues(sourceValues);

  updateRohdatenChangeWishesFormulas(sourceInfo.rowNumber);

  targetSheet.getRange("E2").activate();
  ss.toast(`${sourceValues.length} Zeilen importiert.`, "Rohdaten Änderungswünsche Sort.", 8);
}

function updateRohdatenChangeWishesFormulas(foundRowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("Rohdaten Änderungswünsche Sort.");
  if (!targetSheet || !foundRowNumber) return;

  const urlCellReference = `Einstellungen!$F$${foundRowNumber}`;
  const sheetName = "In_Development_8_weeks_Report_F";

  const newFormulaForJ6 = `=LET(` +
    `url;${urlCellReference};` +
    `sheet_name;"${sheetName}";` +
    `data_range;IMPORTRANGE(url;"'"&sheet_name&"'!E3:Q1000");` +
    `criteria1_range;IMPORTRANGE(url;"'"&sheet_name&"'!B3:B1000");` +
    `criteria2_range;IMPORTRANGE(url;"'"&sheet_name&"'!C3:C1000");` +
    `header_range;IMPORTRANGE(url;"'"&sheet_name&"'!E1:Q1");` +
    `filtered_row;IFERROR(FILTER(data_range;TO_TEXT(criteria1_range)=TO_TEXT($G6);TO_TEXT(criteria2_range)=TO_TEXT($H6)));` +
    `IF(ISBLANK(filtered_row);"";IFERROR(IF(J$5=J$4;INDEX(filtered_row;1;MATCH(J$5;header_range;0));N(INDEX(filtered_row;1;MATCH(J$5;header_range;0)))+N(INDEX(filtered_row;1;MATCH(J$4;header_range;0))))))` +
    `)`;

  const templateFormulaCell = targetSheet.getRange("J6");
  templateFormulaCell.setFormula(newFormulaForJ6);
  templateFormulaCell.autoFill(targetSheet.getRange("J6:Q6"), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}
