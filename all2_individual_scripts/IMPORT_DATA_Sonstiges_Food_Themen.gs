
// /**
//  * 🔎 Imports raw data for 'Sonstiges Food u Akt. Themen'.
//  * V4: "Universal WG Fix" - Scans Rows 1-3 for 'WG', cleans trailing zeros.
//  */
// function importRawOtherFoodTopics() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const ui = SpreadsheetApp.getUi();
//   const settingsSheet = ss.getSheetByName("Einstellungen");
//   const targetSheet = ss.getSheetByName("Sonstiges Food u Akt. Themen");

//   Logger.log("--- Starting importRawOtherFoodTopics (V4 - WG Fix) ---");

//   if (!settingsSheet || !targetSheet) {
//     ui.alert('Error', 'Required sheets missing.', ui.ButtonSet.OK);
//     return;
//   }

//   const settingsValues = settingsSheet.getRange("A1:F" + settingsSheet.getLastRow()).getValues();
//   const searchFileName = "8WS_Food_DE_Sonstiges und aktuelle Themen";
//   let sourceFileId = null;
//   let foundRowNumber = -1;

//   for (let i = 0; i < settingsValues.length; i++) {
//     const cellA = settingsValues[i][0].toString().trim();
//     const cellB = settingsValues[i][1].toString().trim();
//     const excelFileName = searchFileName + ".xlsx";
//     const options = { sensitivity: 'base' }; 
//     if ((cellA && cellA.localeCompare(searchFileName, undefined, options) === 0) ||
//         (cellB && cellB.localeCompare(excelFileName, undefined, options) === 0)) {
//       sourceFileId = settingsValues[i][4]; 
//       foundRowNumber = i + 1;
//       break;
//     }
//   }

//   if (!sourceFileId) {
//     ui.alert('Configuration Error', `File ID for "${searchFileName}" not found.`, ui.ButtonSet.OK);
//     return;
//   }

//   const lastRow = targetSheet.getRange("E" + targetSheet.getMaxRows()).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
//   if (lastRow >= 7) {
//     targetSheet.getRange(7, 1, lastRow - 6, 18).clearContent(); 
//   }

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

//   // --- DYNAMIC HEADER SCAN ---
//   const potentialHeaders = sourceDataSheet.getRange(1, 1, 3, 10).getValues();
//   let wgColIndex = -1;

//   outerLoop:
//   for (let r = 0; r < potentialHeaders.length; r++) {
//     for (let c = 0; c < potentialHeaders[r].length; c++) {
//       if (String(potentialHeaders[r][c]).trim().toUpperCase() === "WG") {
//         wgColIndex = c;
//         break outerLoop;
//       }
//     }
//   }

//   // --- Read Data ---
//   const sourceRange = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 1); 
//   const sourceValues = sourceRange.getDisplayValues();

//   // --- CLEANING LOOP ---
//   for (let r = 0; r < sourceValues.length; r++) {
//     let c = 0; 
//     let val = sourceValues[r][c];

//     // WG Handling
//     if (c === wgColIndex || (wgColIndex === -1 && c === 0)) {
//       if (!val) {
//         sourceValues[r][c] = "";
//       } else {
//         let strVal = String(val).trim();
//         while (strVal.length > 7 && strVal.endsWith('0')) {
//             strVal = strVal.slice(0, -1);
//         }
//         sourceValues[r][c] = strVal; 
//       }
//     } else {
//       sourceValues[r][c] = String(val).replace(/"/g, "").replace(/\u00A0/g, " ").trim();
//     }
//   }

//   sourceRange.setNumberFormat("@"); 
//   sourceRange.setValues(sourceValues);

//   const destinationRange = targetSheet.getRange(7, 5, sourceValues.length, 1);
//   destinationRange.setNumberFormat("@");
//   destinationRange.setValues(sourceValues);

//   targetSheet.getRange("E2").activate();
//   ss.toast('Import successful! ', 'Success', 5);
// }


/**
 * 🔎 Imports raw data for 'Sonstiges Food u Akt. Themen'.
 * Replicates VBA macro 'ImportRohdatenSonstigesFoodundAktuelleThemen' (Modul30).
 * This script only copies data; it does not update any template formulas.
 */
function importRawOtherFoodTopics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const targetSheet = ss.getSheetByName("Sonstiges Food u Akt. Themen");

  Logger.log("--- Starting importRawOtherFoodTopics ---");

  if (!settingsSheet || !targetSheet) {
    ui.alert('Error', 'Could not find "Einstellungen" or "Sonstiges Food u Akt. Themen" sheet.', ui.ButtonSet.OK);
    Logger.log('Error: A required sheet is missing.');
    return;
  }

  // --- 1. Find the Source File ID Row ---
  const searchFileName = "8WS_Food_DE_Sonstiges und aktuelle Themen"; // Source file name for this sheet
  const settingMatch = all2FindSettingsRow_(settingsSheet, searchFileName);
  const sourceFileId = settingMatch
    ? (settingMatch.fileId || all2ExtractDriveId_(settingMatch.fileUrl))
    : "";

  if (!sourceFileId) {
    ui.alert('Configuration Error', `Could not find the file ID for "${searchFileName}". Check the 'Einstellungen' sheet.`, ui.ButtonSet.OK);
    Logger.log(`Error: File ID for "${searchFileName}" not found.`);
    return;
  }

  // --- 2. Prepare Target Sheet (Clear A7:R<lastRow>) ---
  all2ClearContentFromRow_(targetSheet, 7, 18);

  // --- 3. Get Source Data ---
  try {
    const sourceSs = SpreadsheetApp.openById(sourceFileId);
    // The VBA macro (Modul30) uses worksheet 1, which we assume is 'In_Development_8_weeks_Report_F'
    const sourceDataSheet = sourceSs.getSheetByName("In_Development_8_weeks_Report_F"); 
    if (!sourceDataSheet) {
      ui.alert('Sheet Error', 'Could not find "In_Development_8_weeks_Report_F" in source.', ui.ButtonSet.OK);
      return;
    }

    const sourceLastRow = sourceDataSheet.getLastRow();
    if (sourceLastRow < 3) {
      ui.alert('No Data', 'Source file has no data to import (less than 3 rows).', ui.ButtonSet.OK);
      return;
    }

    // --- 4. Copy data from source and paste to target ---
    // VBA: wb.Worksheets(1).Range("A3:A" & lr).Copy
    // VBA: Workbooks("...").Worksheets("Sonstiges Food u Akt. Themen").Range("E7").PasteSpecial xlPasteValues
    const sourceRange = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 1); // Column A, from row 3
    const sourceValues = all2SanitizeDisplayValues_(sourceRange.getDisplayValues());
    
    const destinationRange = targetSheet.getRange(7, 5, sourceValues.length, 1); // Column E, from row 7

    destinationRange.setNumberFormat("@");
    destinationRange.setValues(sourceValues);
    Logger.log(`Successfully imported ${sourceValues.length} rows of data into Column E.`);
    
    // No formula update function is called in this VBA macro.

    targetSheet.getRange("E2").activate(); // Activate E2 as per VBA
    ss.toast('Die Daten wurden erfolgreich importiert! Bitte fülle Spalten P & Q manuell aus und klicke dann auf "DATEN GENERIEREN".', 'Erfolg!', 15);
  
  } catch (e) {
    const errorMessage = `An error occurred: ${e.message}\n\nStack:\n${e.stack}`;
    ui.alert('Import Error', errorMessage, ui.ButtonSet.OK);
    Logger.log(`ERROR: ${e.stack}`);
  }
  Logger.log("--- Finished importRawOtherFoodTopics ---");
}

