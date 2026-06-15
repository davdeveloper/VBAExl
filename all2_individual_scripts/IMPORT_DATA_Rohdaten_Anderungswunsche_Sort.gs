
// /**
//  * 🔎 Imports raw data for Sorting Arrangement Wishes ('Änderungswünsche Sortierung').
//  * V7: HYBRID - Uses V4 structure (Target Col F) + V6 WG Validation (>6 logic).
//  */
// function importChangeWishesSorting() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const ui = SpreadsheetApp.getUi();
//   const settingsSheet = ss.getSheetByName("Einstellungen");
//   const targetSheet = ss.getSheetByName("Rohdaten Änderungswünsche Sort.");

//   Logger.log("--- Starting importChangeWishesSorting (V7 - Hybrid) ---");

//   if (!settingsSheet || !targetSheet) {
//     ui.alert('Error', 'Required sheets missing.', ui.ButtonSet.OK);
//     return;
//   }

//   // --- Find Source File ID ---
//   const settingsValues = settingsSheet.getRange("A1:F" + settingsSheet.getLastRow()).getValues();
//   const searchFileName = "8WS_Food_DE_Änderungswünsche Sortierung";
//   let sourceFileId = null;
//   let foundRowNumber = -1;

//   for (let i = 0; i < settingsValues.length; i++) {
//     const cellA = settingsValues[i][0].toString().trim();
//     const cellB = settingsValues[i][1].toString().trim();
//     const excelFileName = searchFileName + ".xlsx";
//     const options = { sensitivity: 'base' };
//     if ((cellA && cellA.localeCompare(searchFileName, undefined, options) === 0) ||
//       (cellB && cellB.localeCompare(excelFileName, undefined, options) === 0)) {
//       sourceFileId = settingsValues[i][4]; // File ID in Column E
//       foundRowNumber = i + 1;
//       break;
//     }
//   }

//   if (!sourceFileId) {
//     ui.alert('Configuration Error', `File ID for "${searchFileName}" not found.`, ui.ButtonSet.OK);
//     return;
//   }

//   // --- Prepare Target Sheet ---
//   const lastRow = targetSheet.getLastRow();
//   // Clears content starting from Row 8, Columns F:I (Target is Col 6, width 4) based on V4 logic
//   // Actually V4 cleared A8:R, let's stick to the V4 clearing logic to be safe:
//   if (lastRow >= 8) {
//     targetSheet.getRange(8, 1, lastRow - 7, 18).clearContent();
//   }

//   // --- Get Source Data ---
//   let sourceSs;
//   try {
//     sourceSs = SpreadsheetApp.openById(sourceFileId);
//   } catch (e) {
//     ui.alert('File Access Error', `Could not open source file.`, ui.ButtonSet.OK);
//     return;
//   }

//   const sourceDataSheet = sourceSs.getSheetByName("In_Development_8_weeks_Report_F");
//   if (!sourceDataSheet) {
//     ui.alert('Sheet Error', 'Tab "In_Development_8_weeks_Report_F" not found.', ui.ButtonSet.OK);
//     return;
//   }

//   const sourceLastRow = sourceDataSheet.getLastRow();
//   if (sourceLastRow < 3) {
//     ui.alert('No Data', 'Source file has no data.', ui.ButtonSet.OK);
//     return;
//   }

//   // --- Read Data ---
//   // Using getDisplayValues() is safer for the WG string logic than getValues()
//   const sourceRange = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 4);
//   const sourceValues = sourceRange.getDisplayValues(); 

//   // --- HYBRID CLEANING LOOP ---
//   for (let r = 0; r < sourceValues.length; r++) {
//     for (let c = 0; c < sourceValues[r].length; c++) {
      
//       let val = sourceValues[r][c];

//       // --- WG Column Logic (Column 0 / A) ---
//       if (c === 0) {
//         if (!val) {
//           sourceValues[r][c] = "";
//         } else {
//           // 1. Force String & Clean Garbage (Quotes / Non-breaking spaces)
//           let strVal = String(val).replace(/"/g, "").replace(/\u00A0/g, " ").trim();

//           // 2. Apply V6 Logic: Trim zeros if length > 6
//           while (strVal.length > 6 && strVal.endsWith('0')) {
//              strVal = strVal.slice(0, -1);
//           }
//           sourceValues[r][c] = strVal; 
//         }
//       } 
//       // --- Other Columns Logic ---
//       else {
//         sourceValues[r][c] = String(val).replace(/"/g, "").replace(/\u00A0/g, " ").trim();
//       }
//     }
//   }

//   // --- Write clean data BACK to the source file ---
//   // (As requested from V4 logic)
//   sourceRange.setNumberFormat("@"); 
//   sourceRange.setValues(sourceValues); 

//   // --- Write to Target Sheet (Column F / 6) ---
//   // V4 pasted into Column 6 (F), V6 pasted into Column 1. We keep V4 (F).
//   const destinationRange = targetSheet.getRange(8, 6, sourceValues.length, sourceValues[0].length);
//   destinationRange.setNumberFormat("@");
//   destinationRange.setValues(sourceValues);

//   updateRohdatenChangeWishesFormulas(foundRowNumber);

//   targetSheet.getRange("E2").activate();
//   ss.toast('Import successful!', 'Success', 5);
// }

// /**
//  * 🔄 Dynamically updates the template formula in 'Rohdaten Änderungswünsche Sort.' row 6.
//  * @param {number} foundRowNumber The row from 'Einstellungen' with the source file config.
//  */
// function updateRohdatenChangeWishesFormulas(foundRowNumber) {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const targetSheet = ss.getSheetByName("Rohdaten Änderungswünsche Sort.");
//   if (!targetSheet) return;

//   const urlCellReference = `Einstellungen!$F$${foundRowNumber}`;
//   const sheetName = 'In_Development_8_weeks_Report_F';

//   const newFormulaForJ6 = `=LET(` +
//     `url;${urlCellReference};` +
//     `sheet_name;"${sheetName}";` +
//     `data_range;IMPORTRANGE(url;"'"&sheet_name&"'!E3:Q1000");` +
//     `criteria1_range;IMPORTRANGE(url;"'"&sheet_name&"'!B3:B1000");` +
//     `criteria2_range;IMPORTRANGE(url;"'"&sheet_name&"'!C3:C1000");` +
//     `header_range;IMPORTRANGE(url;"'"&sheet_name&"'!E1:Q1");` +
//     `filtered_row;IFERROR(FILTER(data_range;TO_TEXT(criteria1_range)=TO_TEXT($G6);criteria2_range=$H6));` +
//     `IF(ISBLANK(filtered_row);"";IFERROR(IF(J$5=J$4;INDEX(filtered_row;1;MATCH(J$5;header_range;0));N(INDEX(filtered_row;1;MATCH(J$5;header_range;0)))+N(INDEX(filtered_row;1;MATCH(J$4;header_range;0))))))` +
//     `)`;

//   const templateFormulaCell = targetSheet.getRange("J6");
//   templateFormulaCell.setFormula(newFormulaForJ6);

//   const fullTemplateRange = targetSheet.getRange("J6:Q6");
//   templateFormulaCell.autoFill(fullTemplateRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
// }


/**
 * 🔎 Imports raw data for Sorting Arrangement Wishes ('Änderungswünsche Sortierung').
 * FIX: Plain Text method + Dynamic Border Formatting.
 * UPDATE: Aggressively clears ALL borders and content from Row 8 to the very bottom of the sheet.
 */
function importChangeWishesSorting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const targetSheet = ss.getSheetByName("Rohdaten Änderungswünsche Sort.");

  Logger.log("--- Starting importChangeWishesSorting (Strict Plain Text + MaxRow Clear) ---");

  if (!settingsSheet || !targetSheet) {
    ui.alert('Error', 'Required sheets missing.', ui.ButtonSet.OK);
    return;
  }

  // --- Find Source File ID ---
  const searchFileName = "8WS_Food_DE_Änderungswünsche Sortierung";
  const settingMatch = all2FindSettingsRow_(settingsSheet, searchFileName);
  const sourceFileId = settingMatch ? settingMatch.fileId : "";
  const foundRowNumber = settingMatch ? settingMatch.rowNumber : -1;

  if (!sourceFileId) {
    ui.alert('Configuration Error', `File ID for "${searchFileName}" not found.`, ui.ButtonSet.OK);
    return;
  }

  // --- Prepare Target Sheet (THE BORDER FIX) ---
  all2ClearContentFromRow_(targetSheet, 8, 18);

  // --- Get Source Data ---
  let sourceSs;
  try {
    sourceSs = SpreadsheetApp.openById(sourceFileId);
  } catch (e) {
    ui.alert('File Access Error', `Could not open source file.`, ui.ButtonSet.OK);
    return;
  }

  const sourceDataSheet = sourceSs.getSheetByName("In_Development_8_weeks_Report_F");
  if (!sourceDataSheet) {
    ui.alert('Sheet Error', 'Tab "In_Development_8_weeks_Report_F" not found.', ui.ButtonSet.OK);
    return;
  }

  const sourceLastRow = sourceDataSheet.getLastRow();
  if (sourceLastRow < 3) {
    ui.alert('No Data', 'Source file has no data.', ui.ButtonSet.OK);
    return;
  }

  // --- Read Data (using getDisplayValues to capture exact text) ---
  const sourceRange = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 4);
  const sourceValues = all2SanitizeDisplayValues_(sourceRange.getDisplayValues()); 

  // --- Write to Target Sheet (Column F / 6) as Plain Text ---
  const destinationRange = targetSheet.getRange(8, 6, sourceValues.length, sourceValues[0].length);
  destinationRange.setNumberFormat("@");
  destinationRange.setValues(sourceValues);

  // --- 3. Apply Formatting (Borders/Fonts/Colors) from Row 6 ---
  const formatTemplate = targetSheet.getRange(6, 1, 1, 18); // A6:R6 template row
  const targetFormatRange = targetSheet.getRange(8, 1, sourceValues.length, 18);
  formatTemplate.copyTo(targetFormatRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  updateRohdatenChangeWishesFormulas(foundRowNumber);

  targetSheet.getRange("E2").activate();
  ss.toast('Import successful!', 'Success', 5);
}

/**
 * 🔄 Dynamically updates the template formula in 'Rohdaten Änderungswünsche Sort.'
 */
function updateRohdatenChangeWishesFormulas(foundRowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("Rohdaten Änderungswünsche Sort.");
  if (!targetSheet) return;

  const urlCellReference = `Einstellungen!$F$${foundRowNumber}`;
  const sheetName = 'In_Development_8_weeks_Report_F';

  const newFormulaForJ6 = `=LET(` +
    `url;${urlCellReference};` +
    `sheet_name;"${sheetName}";` +
    `data_range;IMPORTRANGE(url;"'"&sheet_name&"'!E3:Q1000");` +
    `criteria1_range;IMPORTRANGE(url;"'"&sheet_name&"'!B3:B1000");` +
    `criteria2_range;IMPORTRANGE(url;"'"&sheet_name&"'!C3:C1000");` +
    `header_range;IMPORTRANGE(url;"'"&sheet_name&"'!E1:Q1");` +
    `filtered_row;IFERROR(FILTER(data_range;TO_TEXT(criteria1_range)=TO_TEXT($G6);criteria2_range=$H6));` +
    `IF(ISBLANK(filtered_row);"";IFERROR(IF(J$5=J$4;INDEX(filtered_row;1;MATCH(J$5;header_range;0));N(INDEX(filtered_row;1;MATCH(J$5;header_range;0)))+N(INDEX(filtered_row;1;MATCH(J$4;header_range;0))))))` +
    `)`;

  const templateFormulaCell = targetSheet.getRange("J6");
  templateFormulaCell.setFormula(newFormulaForJ6);

  const fullTemplateRange = targetSheet.getRange("J6:Q6");
  templateFormulaCell.autoFill(fullTemplateRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}
