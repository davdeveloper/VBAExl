/**
 * 🔎 Imports raw data for 'Rohdaten O+G Qualitätsf. o. LFN'.
 * Replicates VBA macro 'ImportRohdatenOGQualitätsfälleohneLFN' (Modul42).
 * V1: Fixes bug where sourceFileId was never assigned.
 */
function importRawOGQualityCasesNoLFN() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const targetSheet = ss.getSheetByName("Rohdaten O+G Qualitätsf. o. LFN");

  Logger.log("--- Starting importRawOGQualityCasesNoLFN (V1 - ID Fix) ---");

  if (!settingsSheet || !targetSheet) {
    ui.alert('Error', 'Could not find "Einstellungen" or "Rohdaten O+G Qualitätsf. o. LFN" sheet.', ui.ButtonSet.OK);
    Logger.log('Error: A required sheet is missing.');
    return;
  }

  // --- 1. Find the Source File ID Row ---
  const searchFileName = "8WS_O+G_DE_Qualitätsfälle (ohne LFN)"; // Source file name for this sheet
  const settingMatch = all2FindSettingsRow_(settingsSheet, searchFileName);
  const sourceFileId = settingMatch
    ? (settingMatch.fileId || all2ExtractDriveId_(settingMatch.fileUrl))
    : "";

  // This check will now pass
  if (!sourceFileId) {
    ui.alert('Configuration Error', `Could not find the file ID for "${searchFileName}". Check the 'Einstellungen' sheet.`, ui.ButtonSet.OK);
    Logger.log(`Error: File ID for "${searchFileName}" not found.`);
    return;
  }

  all2ClearContentFromRow_(targetSheet, 8, 17);

  // --- 3. Get Source Data ---
  try {
    const sourceSs = SpreadsheetApp.openById(sourceFileId);
    // The VBA macro (Modul42) uses worksheet 1
    const sourceDataSheet = sourceSs.getSheets()[0]; // Use the first sheet
    if (!sourceDataSheet) {
      ui.alert('Sheet Error', 'Could not find the first sheet in the source file.', ui.ButtonSet.OK);
      return;
    }
    Logger.log(`Using source sheet: ${sourceDataSheet.getName()}`);

    const sourceLastRow = sourceDataSheet.getLastRow();
    if (sourceLastRow < 3) {
      ui.alert('No Data', 'Source file has no data to import (less than 3 rows).', ui.ButtonSet.OK);
      return;
    }

    // --- 4. Copy data from source and paste to target ---
    // VBA: wb.Worksheets(1).Range("A3:F" & lr).Copy
    // VBA: ...Worksheets("Rohdaten O+G Qualitätsf. o. LFN").Range("B8").PasteSpecial xlPasteValues
    const sourceRange = sourceDataSheet.getRange(3, 1, sourceLastRow - 2, 6); // A3:F<sourceLastRow>
    const sourceValues = all2SanitizeDisplayValues_(sourceRange.getDisplayValues());
    
    // Paste into B8 (row 8, column 2)
    const destinationRange = targetSheet.getRange(8, 2, sourceValues.length, sourceValues[0].length); // B8:G<lastRow>

    // --- !! ADDED LINE TO OVERWRITE DESTINATION FORMATS !! ---
    // Set destination to Plain Text to ignore any pre-existing formats
    destinationRange.setNumberFormat("@"); 
    // --- !! END OF ADDED LINE !! ---

    destinationRange.setValues(sourceValues);
    Logger.log(`Successfully imported ${sourceValues.length} rows of data into B8:G${7 + sourceValues.length - 1}.`);
    
    targetSheet.getRange("E2").activate(); // Activate E2 as per VBA
    ss.toast('Die Daten wurden erfolgreich importiert! Klicke jetzt auf "DATEN GENERIEREN".', 'Erfolg!', 10);
  
  } catch (e) {
    const errorMessage = `An error occurred: ${e.message}\n\nStack:\n${e.stack}`;
    ui.alert('Import Error', errorMessage, ui.ButtonSet.OK);
    Logger.log(`ERROR: ${e.stack}`);
  }
  Logger.log("--- Finished importRawOGQualityCasesNoLFN ---");
}

