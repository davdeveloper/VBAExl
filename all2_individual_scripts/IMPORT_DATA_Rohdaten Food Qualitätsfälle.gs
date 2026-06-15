
/**
 * Shared helpers for the migrated Apps Script modules in this file.
 * Names are prefixed to avoid collisions with the main project.
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

// /**
//  * 🔎 Imports raw data for 'Rohdaten Food Qualitätsfälle'.
//  * V4: "Universal WG Fix" - Dynamic Header Scan (Rows 1-3), WG Sanitization.
//  */
// function importRawFoodQualityCases() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const ui = SpreadsheetApp.getUi();
//   const settingsSheet = ss.getSheetByName("Einstellungen");
//   const targetSheet = ss.getSheetByName("Rohdaten Food Qualitätsfälle");

//   Logger.log("--- Starting importRawFoodQualityCases (V4 - WG Fix) ---");

//   if (!settingsSheet || !targetSheet) {
//     ui.alert('Error', 'Required sheets missing.', ui.ButtonSet.OK);
//     return;
//   }

//   const settingsValues = settingsSheet.getRange("A1:F" + settingsSheet.getLastRow()).getValues();
//   const searchFileName = "8WS_Food_DE_Qualitätsfälle";
//   let sourceFileId = null;

//   for (let i = 0; i < settingsValues.length; i++) {
//     const cellA = settingsValues[i][0].toString().trim();
//     const cellB = settingsValues[i][1].toString().trim();
//     const excelFileName = searchFileName + ".xlsx";
//     const options = { sensitivity: 'base' }; 
//     if ((cellA && cellA.localeCompare(searchFileName, undefined, options) === 0) ||
//         (cellB && cellB.localeCompare(excelFileName, undefined, options) === 0)) {
//       sourceFileId = settingsValues[i][4]; 
//       break;
//     }
//   }

//   if (!sourceFileId) {
//     ui.alert('Configuration Error', `File ID for "${searchFileName}" not found.`, ui.ButtonSet.OK);
//     return;
//   }

//   const lastRowK = targetSheet.getRange("K" + targetSheet.getMaxRows()).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
//   const lastRow = lastRowK < 8 ? 8 : lastRowK;
//   if (lastRow >= 8) {
//     targetSheet.getRange(8, 1, lastRow - 7, 25).clearContent();
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
//   if (sourceLastRow < 3) return;

//   // --- DYNAMIC HEADER SCAN ---
//   const potentialHeaders = sourceDataSheet.getRange(1, 1, 3, 10).getValues();
//   let wgColIndex = -1;

//   outerLoop:
//   for (let r = 0; r < potentialHeaders.length; r++) {
//     for (let c = 0; c < potentialHeaders[r].length; c++) {
//       if (String(potentialHeaders[r][c]).trim().toUpperCase() === "WG") {
//         wgColIndex = c;
//         Logger.log(`Found 'WG' header at Row ${r + 1}, Column ${c + 1}`);
//         break outerLoop;
//       }
//     }
//   }

//   // --- Read Data ---
//   const sourceRange = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 8); 
//   const sourceValues = sourceRange.getDisplayValues();

//   // --- CLEANING LOOP ---
//   for (let r = 0; r < sourceValues.length; r++) {
//     for (let c = 0; c < sourceValues[r].length; c++) {
//       let val = sourceValues[r][c];

//       // WG Handling
//       if (c === wgColIndex || (wgColIndex === -1 && c === 0)) {
//         if (!val) {
//           sourceValues[r][c] = "";
//         } else {
//           let strVal = String(val).trim();
//           while (strVal.length > 5 && strVal.endsWith('0')) {
//              strVal = strVal.slice(0, -1);
//           }
//           sourceValues[r][c] = strVal; 
//         }
//       } else {
//         sourceValues[r][c] = String(val).replace(/"/g, "").replace(/\u00A0/g, " ").trim();
//       }
//     }
//   }

//   sourceRange.setNumberFormat("@"); 
//   sourceRange.setValues(sourceValues);

//   const destinationRange = targetSheet.getRange(8, 7, sourceValues.length, sourceValues[0].length);
//   destinationRange.setNumberFormat("@");
//   destinationRange.setValues(sourceValues);
  
//   targetSheet.getRange("E2").activate();
//   ss.toast('Import successful! WG column sanitized.', 'Success', 5);
// }

/**
 * 🔎 Imports raw data for 'Rohdaten Food Qualitätsfälle'.
 * FIX: Plain Text method + Paste Values Only (Preserves existing borders/colors).
 * Removes destructive .toFixed(3) logic that was altering WG codes.
 */
function importRawFoodQualityCases() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const targetSheet = ss.getSheetByName("Rohdaten Food Qualitätsfälle");

  Logger.log("--- Starting importRawFoodQualityCases (Strict Plain Text) ---");
  if (!settingsSheet || !targetSheet) {
    ui.alert('Error', 'Required sheets missing.', ui.ButtonSet.OK);
    return;
  }

  const searchFileName = "8WS_Food_DE_Qualitätsfälle";
  const settingMatch = all2FindSettingsRow_(settingsSheet, searchFileName);
  const sourceFileId = settingMatch ? settingMatch.fileId : "";

  if (!sourceFileId) {
    ui.alert('Configuration Error', `File ID for "${searchFileName}" not found.`, ui.ButtonSet.OK);
    return;
  }

// --- Prepare Target Sheet (Clear CONTENT only to preserve borders/colors) ---
  all2ClearContentFromRow_(targetSheet, 8, 25);

  let sourceSs;
  try {
    sourceSs = SpreadsheetApp.openById(sourceFileId);
  } catch (e) {
    ui.alert('File Access Error', `Could not open source file.`, ui.ButtonSet.OK);
    return;
  }
  
  const sourceDataSheet = sourceSs.getSheetByName("In_Development_8_weeks_Report_F");
  if (!sourceDataSheet) {
    ui.alert('Sheet Error', 'Tab not found.', ui.ButtonSet.OK);
    return;
  }
  const sourceLastRow = sourceDataSheet.getLastRow();
  if (sourceLastRow < 3) return;

  // --- Read Data (using getDisplayValues to capture exact text) ---
  const sourceRange = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 8);
  const sourceValues = all2SanitizeDisplayValues_(sourceRange.getDisplayValues());

  // --- Write to Target Sheet (Column G / 7) as Plain Text (Values Only) ---
  const destinationRange = targetSheet.getRange(8, 7, sourceValues.length, sourceValues[0].length); 
  destinationRange.setNumberFormat("@");
  destinationRange.setValues(sourceValues); // setValues inherently acts like "Paste Values Only"
  
  targetSheet.getRange("E2").activate();
  ss.toast('Import successful! Data formatted perfectly as Plain Text.', 'Success', 5);
}
