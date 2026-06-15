// /**
//  * 🔎 Imports raw data for Food Assortment Wishes and updates template formulas.
//  * V3: Implements perfected pre-import cleaning logic and source write-back.
//  */
// function importFoodAssortmentWishes() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const ui = SpreadsheetApp.getUi();
//   const settingsSheet = ss.getSheetByName("Einstellungen");
//   const targetSheet = ss.getSheetByName("Rohdaten Food Sortimentswünsche");

//   Logger.log("--- Starting importFoodAssortmentWishes (V3 - Perfected Clean) ---");

//   if (!settingsSheet || !targetSheet) {
//     ui.alert('Error', 'Could not find "Einstellungen" or "Rohdaten Food Sortimentswünsche" sheet.', ui.ButtonSet.OK);
//     Logger.log('Error: A required sheet is missing.');
//     return;
//   }

//   // --- Find Source File ID ---
//   const settingsValues = settingsSheet.getRange("A1:F" + settingsSheet.getLastRow()).getValues();
//   const searchFileName = "8WS_Food_DE_Sortimentswünsche";
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
//       Logger.log(`Source file config found on row ${foundRowNumber} with ID: ${sourceFileId}`);
//       break;
//     }
//   }

//   if (!sourceFileId) {
//     ui.alert('Configuration Error', `Could not find the file ID for "${searchFileName}".`, ui.ButtonSet.OK);
//     Logger.log(`Error: File ID for "${searchFileName}" not found.`);
//     return;
//   }

//   // --- Prepare Target Sheet ---
//   const lastRow = targetSheet.getLastRow();
//   if (lastRow >= 8) {
//     targetSheet.getRange(8, 1, lastRow - 7, 18).clearContent();
//     Logger.log(`Cleared old data from A8:R${lastRow}`);
//   }

//   // --- Get Source Data ---
//   let sourceSs;
//   try {
//     sourceSs = SpreadsheetApp.openById(sourceFileId);
//   } catch (e) {
//     ui.alert('File Access Error', `Could not open source file: ${e.message}`, ui.ButtonSet.OK);
//     return;
//   }
//   const sourceDataSheet = sourceSs.getSheetByName("In_Development_8_weeks_Report_F");
//   if (!sourceDataSheet) {
//     ui.alert('Sheet Error', 'Could not find "In_Development_8_weeks_Report_F" in source.', ui.ButtonSet.OK);
//     return;
//   }

//   const sourceLastRow = sourceDataSheet.getLastRow();
//   if (sourceLastRow < 3) {
//     ui.alert('No Data', 'Source file has no data to import.', ui.ButtonSet.OK);
//     return;
//   }

//   // --- Copy data from source and paste to target ---
//   try {
//     const sourceRange = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 4); // Reads columns A, B, C, D
//     const sourceValues = sourceRange.getValues();
//     Logger.log(`Source values read. Starting cleaning...`);


//     // --- START: Perfected Cleaning Loop ---
//     for (let r = 0; r < sourceValues.length; r++) {
//       for (let c = 0; c < sourceValues[r].length; c++) {

//         let cellValue = sourceValues[r][c];

//         // 1. Specific fix for Col G (WG), which is c === 0 (the first column)
//         // This is source column A, which maps to target column F.
//         if (c === 0) {
//           // First, clean it up if it's a string to remove spaces, quotes, etc.
//           if (typeof cellValue === 'string') {
//             cellValue = cellValue.replace(/"/g, "")
//               .replace(/\u00A0/g, " ")
//               .trim();
//           }

//           // Now, check if the (potentially cleaned) value is numeric
//           let numericValue;
//           if (typeof cellValue === 'number') {
//             numericValue = cellValue;
//           } else if (typeof cellValue === 'string' && cellValue !== "") {
//             // Replace comma with dot for robust parsing, then parse
//             const parsableString = cellValue.replace(',', '.');
//             numericValue = parseFloat(parsableString);
//           } else {
//             // It's an empty string or some other type, not a number
//             numericValue = NaN;
//           }

//           // If it's a valid finite number, format it to 2 decimal places.
//           // isFinite() handles NaN and Infinity
//           if (isFinite(numericValue)) {
//             sourceValues[r][c] = numericValue.toFixed(2); // e.g., 30.1 -> "30.10"
//           } else {
//             // It's not a number (e.g., "N/A" or was empty), just assign the cleaned string
//             // If it was a number originally (like Infinity), convert to string
//             sourceValues[r][c] = (typeof cellValue === 'string') ? cellValue : String(cellValue);
//           }
//         }
//         // 2. Clean all other string values (c > 0)
//         else if (typeof cellValue === 'string') {
//           sourceValues[r][c] = cellValue.replace(/"/g, "")         // 1. Remove all quotes
//             .replace(/\u00A0/g, " ") // 2. Replace non-breaking spaces
//             .trim();                 // 3. Remove all leading/trailing spaces
//         }
//       }
//     }
//     Logger.log("Data cleaning loop finished.");
//     // --- END: Perfected Cleaning Loop ---

//     // --- Write clean data BACK to the source file ---
//     Logger.log("Writing clean data back to source file...");
//     sourceRange.setNumberFormat("@"); // Set source columns A:D to Plain Text
//     sourceRange.setValues(sourceValues); // Paste clean values back
//     Logger.log("Source file successfully cleaned.");
//     // --- End of "write-back" code ---

//     const destinationRange = targetSheet.getRange(8, 6, sourceValues.length, sourceValues[0].length); // F8:I...

//     // --- Force target columns F:I to Plain Text ---
//     Logger.log("Setting destination columns F:I to Plain Text format.");
//     destinationRange.setNumberFormat("@");

//     destinationRange.setValues(sourceValues);
//     Logger.log(`Successfully imported ${sourceValues.length} rows of clean data.`);

//     updateRohdatenFoodFormulas(foundRowNumber); // Update template formula in row 6

//     targetSheet.getRange("E2").activate();
//     ss.toast('Die Daten wurden erfolgreich importiert! Klicke jetzt auf "DATEN GENERIEREN".', 'Erfolg!', 10);
//   } catch (e) {
//     ui.alert('Import Error', `An error occurred: ${e.message}`, ui.ButtonSet.OK);
//   }
//   Logger.log("--- Finished importFoodAssortmentWishes ---");
// }


// /**
//  * 🔄 Dynamically updates the template formula in 'Rohdaten Food Sortimentswünsche' row 6.
//  * @param {number} foundRowNumber The row from 'Einstellungen' with the source file config.
//  */
// function updateRohdatenFoodFormulas(foundRowNumber) {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const targetSheet = ss.getSheetByName("Rohdaten Food Sortimentswünsche");
//   if (!targetSheet) return;

//   const urlCellReference = `Einstellungen!$F$${foundRowNumber}`;
//   const sheetName = 'In_Development_8_weeks_Report_F';

//   // Same robust formula, but referencing $G6 for the template row
//   const newFormulaForJ6 = `=ARRAY_CONSTRAIN(ARRAYFORMULA(IFERROR(IF(J$5=J$4;INDEX(IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$3:$Q$1000");MATCH(1;($G6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$B$3:$B$1000"))*($H6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$C$3:$C$1000"));0);MATCH(J$5;IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$1:$Q$1");0));IFERROR(NUMBERVALUE(INDEX(IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$3:$Q$1000");MATCH(1;($G6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$B$3:$B$1000"))*($H6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$C$3:$C$1000"));0);MATCH(J$5;IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$1:$Q$1");0)));0)+IFERROR(NUMBERVALUE(INDEX(IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$3:$Q$1000");MATCH(1;($G6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$B$3:$B$1000"))*($H6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$C$3:$C$1000"));0);MATCH(J$4;IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$1:$Q$1");0)));0));"")); 1; 1)`;

//   const templateFormulaCell = targetSheet.getRange("J6");
//   templateFormulaCell.setFormula(newFormulaForJ6);

//   const fullTemplateRange = targetSheet.getRange("J6:Q6");
//   templateFormulaCell.autoFill(fullTemplateRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

//   Logger.log(`Successfully updated template formula in J6:Q6 to reference ${urlCellReference}.`);
// }

/**
 * 🔎 Imports raw data for Food Assortment Wishes.
 * FIX: Uses getDisplayValues() and Plain Text format ("@") on both SOURCE and TARGET 
 * to ensure WG and CW columns remain exactly as they are without zero or date manipulation.
 */
function importFoodAssortmentWishes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const targetSheet = ss.getSheetByName("Rohdaten Food Sortimentswünsche");
  
  Logger.log("--- Starting importFoodAssortmentWishes (Strict Plain Text) ---");

  if (!settingsSheet || !targetSheet) {
    ui.alert('Error', 'Required sheets missing.', ui.ButtonSet.OK);
    return;
  }

  // --- Find Source File ID ---
  const searchFileName = "8WS_Food_DE_Sortimentswünsche";
  const settingMatch = all2FindSettingsRow_(settingsSheet, searchFileName);
  const sourceFileId = settingMatch ? settingMatch.fileId : "";
  const foundRowNumber = settingMatch ? settingMatch.rowNumber : -1;

  if (!sourceFileId) {
    ui.alert('Configuration Error', `File ID for "${searchFileName}" not found in Settings.`, ui.ButtonSet.OK);
    return;
  }

  // --- Prepare Target ---
  all2ClearContentFromRow_(targetSheet, 8, 18);

  // --- Get Source Data ---
  let sourceSs;
  try {
    sourceSs = SpreadsheetApp.openById(sourceFileId);
  } catch (e) {
    ui.alert('File Access Error', `Could not open source file: ${e.message}`, ui.ButtonSet.OK);
    return;
  }
  
  const sourceDataSheet = sourceSs.getSheetByName("In_Development_8_weeks_Report_F");
  if (!sourceDataSheet) {
    ui.alert('Sheet Error', 'Tab "In_Development_8_weeks_Report_F" not found.', ui.ButtonSet.OK);
    return;
  }

  const sourceLastRow = sourceDataSheet.getLastRow();
  if (sourceLastRow < 3) return;

  // Read Data using DisplayValues to capture exact text representation (preserves "30.10" and "2026-01")
  const sourceRange = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 4);
  const sourceValues = all2SanitizeDisplayValues_(sourceRange.getDisplayValues()); 

  // --- Write to Target as Plain Text ---
  const destinationRange = targetSheet.getRange(8, 6, sourceValues.length, sourceValues[0].length); 
  destinationRange.setNumberFormat("@"); 
  destinationRange.setValues(sourceValues);
  
  updateRohdatenFoodFormulas(foundRowNumber);

  targetSheet.getRange("E2").activate();
  ss.toast('Import successful! WG and CW columns sanitized as Plain Text.', 'Success', 5);
}

function updateRohdatenFoodFormulas(foundRowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("Rohdaten Food Sortimentswünsche");
  if (!targetSheet) return;

  const newFormulaForJ6 = `=ARRAY_CONSTRAIN(ARRAYFORMULA(IFERROR(IF(J$5=J$4;INDEX(IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$3:$Q$1000");MATCH(1;($G6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$B$3:$B$1000"))*($H6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$C$3:$C$1000"));0);MATCH(J$5;IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$1:$Q$1");0));IFERROR(NUMBERVALUE(INDEX(IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$3:$Q$1000");MATCH(1;($G6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$B$3:$B$1000"))*($H6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$C$3:$C$1000"));0);MATCH(J$5;IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$1:$Q$1");0)));0)+IFERROR(NUMBERVALUE(INDEX(IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$3:$Q$1000");MATCH(1;($G6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$B$3:$B$1000"))*($H6=IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$C$3:$C$1000"));0);MATCH(J$4;IMPORTRANGE(Einstellungen!$F$${foundRowNumber};"In_Development_8_weeks_Report_F!$E$1:$Q$1");0)));0));"")); 1; 1)`;

  const templateFormulaCell = targetSheet.getRange("J6");
  templateFormulaCell.setFormula(newFormulaForJ6);
  const fullTemplateRange = targetSheet.getRange("J6:Q6");
  templateFormulaCell.autoFill(fullTemplateRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

