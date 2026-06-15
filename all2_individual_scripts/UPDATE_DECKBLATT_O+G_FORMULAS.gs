
/**
 * 🔄 Dynamically Updates Formulas on the 'Deckblatt O+G' Sheet.
 * This version correctly creates formulas that REFERENCE the cell containing the URL
 * in the 'Einstellungen' sheet, rather than embedding the URL string itself.
 * This is run from the runMisImportSmarter() function on the IMPORT_MIS_DATEN code page.
 */
function updateDeckblatt_O_G_Formulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Einstellungen");
  const deckblattSheet = ss.getSheetByName("Deckblatt O+G");
  const ui = SpreadsheetApp.getUi();

  // --- 1. Validate that the necessary sheets exist ---
  if (!settingsSheet) {
    ui.alert('Error', 'Sheet "Einstellungen" could not be found. The script cannot continue.', ui.ButtonSet.OK);
    return;
  }
  if (!deckblattSheet) {
    ui.alert('Error', 'Sheet "Deckblatt O+G" could not be found. The script cannot continue.', ui.ButtonSet.OK);
    return;
  }

  // --- 2. Find the correct row number ---
  const searchText = "Deckblatt O+G";
  const foundRowNumber = all2FindRowContainingAny_(settingsSheet, [searchText, "Deckblatt OG"]);

  // --- 3. Proceed only if the row was found ---
  if (foundRowNumber > -1) {
    try {
      // --- 4. Construct the cell reference and new formulas ---
      const sheetName = "X8_weeks_Report_Deckblatt_DE";
      // This creates the dynamic, absolute reference string, e.g., "Einstellungen!$F$7"
      const urlCellReference = `Einstellungen!$F$${foundRowNumber}`;
      
      const newFormulaM4 = `=IF(AND(J4<ISOWEEKNUM(TODAY()-56);YEAR(TODAY()-56)=YEAR(TODAY()));R4;XLOOKUP(L4;IMPORTRANGE(${urlCellReference};"'${sheetName}'!A:A");IMPORTRANGE(${urlCellReference};"'${sheetName}'!B:B");"";0;1))`;
      const newFormulaN4 = `=IF(AND(J4<ISOWEEKNUM(TODAY()-56);YEAR(TODAY()-56)=YEAR(TODAY()));S4;XLOOKUP(L4;IMPORTRANGE(${urlCellReference};"'${sheetName}'!A:A");IMPORTRANGE(${urlCellReference};"'${sheetName}'!C:C");"";0;1))`;
      const newFormulaO4 = `=IF(AND(J4<ISOWEEKNUM(TODAY()-56);YEAR(TODAY()-56)=YEAR(TODAY()));T4;XLOOKUP(L4;IMPORTRANGE(${urlCellReference};"'${sheetName}'!A:A");IMPORTRANGE(${urlCellReference};"'${sheetName}'!D:D");"";0;1))`;

      // --- 5. Apply formulas to the top row AND auto-fill down to row 57 ---
      deckblattSheet.getRange("M4").setFormula(newFormulaM4);
      deckblattSheet.getRange("N4").setFormula(newFormulaN4);
      deckblattSheet.getRange("O4").setFormula(newFormulaO4);

      const sourceRange = deckblattSheet.getRange("M4:O4");
      const destinationRange = deckblattSheet.getRange("M4:O57");
      
      sourceRange.autoFill(destinationRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
      
      //ui.alert('Success!', `The formulas in M4:O57 on "Deckblatt O+G" have been updated to reference cell ${urlCellReference}.`, ui.ButtonSet.OK);
      SpreadsheetApp.getActive().toast(`The formulas in M4:O57 on "Deckblatt O+G" have been updated to reference cell ${urlCellReference}.`);

    } catch (e) {
      ui.alert('Formula Error', `Could not update the formulas. Error: ${e.message}`, ui.ButtonSet.OK);
    }
  } else {
    // Show an error if the specific row wasn't found
    ui.alert('Configuration Error', `Could not find a row containing "${searchText}" in the "Einstellungen" sheet. Please check the sheet configuration.`, ui.ButtonSet.OK);
  }
}


