/**
 * Deletes the current import/generation data areas across the report.
 *
 * Button function recommendation: importLoeschenAlles
 *
 * This preserves header rows, template formulas, formatting, borders and deckblatt formulas.
 * It only clears the data rows that are normally filled by the import/generate buttons.
 */
function importLoeschenAlles() {
  clearAllImportsForRestart();
}

function clearAllReportImports() {
  clearAllImportsForRestart();
}

function clearAllImportsForRestart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    "Importe löschen?",
    "Es werden die aktuellen Import-/Generierungsbereiche in allen Report-Blättern geleert. Formatierungen, Kopfzeilen und Template-Formeln bleiben erhalten.",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ss.toast("Löschen abgebrochen.", "Abbruch", 5);
    return;
  }

  const configs = [
    { sheet: "MIS Final", startRow: 6, columns: 7 },

    { sheet: "Rohdaten Food Sortimentswünsche", startRow: 8, columns: 25 },
    { sheet: "Food Sortimentswünsche", startRow: 6, columns: 18 },

    { sheet: "Rohdaten Änderungswünsche Sort.", startRow: 8, columns: 25 },
    { sheet: "Änderungswünsche Sortierung", startRow: 6, columns: 18 },

    { sheet: "Rohdaten Allg. Sortimentsw.", startRow: 8, columns: 11 },
    { sheet: "Allgemeine Sortimentswünsche", startRow: 8, columns: 11 },

    { sheet: "Sonstiges Food u Akt. Themen", startRow: 7, columns: 18 },

    { sheet: "Rohdaten Food Qualitätsfälle", startRow: 8, columns: 25 },
    { sheet: "Food Qualitätsfälle", startRow: 6, columns: 22 },

    { sheet: "O+G Sortimentswünsche", startRow: 7, columns: 11 },
    { sheet: "O+G Sonstiges", startRow: 7, columns: 11 },

    { sheet: "Rohdaten O+G Qualitätsfälle", startRow: 8, columns: 17 },
    { sheet: "O+G Qualitätsfälle", startRow: 6, columns: 15 },

    { sheet: "Rohdaten O+G Qualitätsf. o. LFN", startRow: 8, columns: 17 },
    { sheet: "O+G Qualitätsf. o. LFN", startRow: 6, columns: 15 }
  ];

  const cleared = [];
  const missing = [];

  for (const config of configs) {
    const sheet = ss.getSheetByName(config.sheet);
    if (!sheet) {
      missing.push(config.sheet);
      continue;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < config.startRow) continue;

    const columnCount = Math.min(config.columns, sheet.getMaxColumns());
    sheet.getRange(config.startRow, 1, lastRow - config.startRow + 1, columnCount).clearContent();
    cleared.push(config.sheet);
  }

  resetDeleteTempSheets_();

  const messageParts = [`${cleared.length} Bereiche geleert.`];
  if (missing.length) {
    messageParts.push(`Nicht gefunden: ${missing.join(", ")}`);
  }

  ss.toast(messageParts.join(" "), "Importe gelöscht", 10);
  ui.alert("Fertig", messageParts.join("\n"), ui.ButtonSet.OK);
}

function resetDeleteTempSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  for (const sheet of sheets) {
    if (sheet.getName().toLowerCase().startsWith("temp_") && ss.getSheets().length > 1) {
      try {
        ss.deleteSheet(sheet);
      } catch (e) {
        Logger.log(`Temp-Blatt konnte nicht gelöscht werden: ${sheet.getName()} - ${e.message}`);
      }
    }
  }
}
