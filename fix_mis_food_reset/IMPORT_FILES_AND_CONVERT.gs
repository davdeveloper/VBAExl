/**
 * Converts German umlauts and ss to simple ASCII equivalents for robust file matching.
 */
function normalizeGermanUmlauts(str) {
  if (typeof str !== "string") return str;
  return str.replace(/Ä/g, "A").replace(/ä/g, "a")
            .replace(/Ö/g, "O").replace(/ö/g, "o")
            .replace(/Ü/g, "U").replace(/ü/g, "u")
            .replace(/ß/g, "s");
}

function convertNormalizeText_(value) {
  return String(value == null ? "" : value).normalize("NFC").trim();
}

function convertExtractDriveFolderId_(value) {
  const text = convertNormalizeText_(value);
  const patterns = [
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

function convertCleanDisplayText_(value) {
  return convertNormalizeText_(value).replace(/"/g, "").replace(/\u00A0/g, " ").trim();
}

function convertCellForPlainText_(rawValue, displayValue) {
  if (Object.prototype.toString.call(rawValue) === "[object Date]") {
    const year = rawValue.getFullYear();
    const month = String(rawValue.getMonth() + 1).padStart(2, "0");
    return `${year}-${month}`;
  }

  return convertCleanDisplayText_(displayValue);
}

/**
 * Converts listed Excel files and updates/creates matching Google Sheets.
 *
 * Important fix:
 * - Reads getDisplayValues() from the converted temp sheet instead of getValues().
 *   This preserves values exactly as users see them in Excel/Sheets, e.g.
 *   190.164 does not become the raw unrounded value 190.163.788.
 *
 * Button function: convertExcelFilesToSheets
 */
function convertExcelFilesToSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const fileNamesRange = sheet.getRange("B4:B17");
    const fileNamesValues = fileNamesRange.getValues();
    const namesMap = new Map();

    fileNamesValues.flat().forEach(name => {
      const trimmedName = convertNormalizeText_(name);
      if (trimmedName.length < 3) return;

      let baseName = trimmedName;
      if (trimmedName.toLowerCase().endsWith(".xlsx")) baseName = trimmedName.slice(0, -5);
      else if (trimmedName.toLowerCase().endsWith(".xls")) baseName = trimmedName.slice(0, -4);

      const robustName = normalizeGermanUmlauts(baseName);
      if (robustName.length >= 3) {
        namesMap.set(robustName, trimmedName);
      }
    });

    if (namesMap.size === 0) {
      ui.alert("Keine Dateien", "In B4:B17 wurden keine gültigen Dateinamen gefunden.", ui.ButtonSet.OK);
      return;
    }

    const folderId = convertExtractDriveFolderId_(sheet.getRange("B19").getValue());
    if (!folderId) {
      ui.alert("Ordner-Link fehlt", "Bitte den vollständigen Google-Drive-Ordnerlink oder die Ordner-ID in B19 eintragen.", ui.ButtonSet.OK);
      return;
    }

    let processedCount = 0;
    const finalFileIds = new Map();
    const allFilesInFolder = [];
    let pageToken;

    do {
      const listResult = Drive.Files.list({
        q: `'${folderId}' in parents and trashed = false`,
        corpora: "allDrives",
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        pageToken: pageToken,
        pageSize: 1000
      });

      if (listResult.items) allFilesInFolder.push(...listResult.items);
      pageToken = listResult.nextPageToken;
    } while (pageToken);

    const excelFilesToProcess = allFilesInFolder.filter(file => {
      const isExcel = file.title.toLowerCase().endsWith(".xlsx") || file.title.toLowerCase().endsWith(".xls");
      if (!isExcel) return false;

      const baseNameRaw = file.title.replace(/\.xlsx$/i, "").replace(/\.xls$/i, "");
      return namesMap.has(normalizeGermanUmlauts(baseNameRaw));
    });

    for (const file of excelFilesToProcess) {
      const fileName = file.title;
      const baseNameRaw = fileName.replace(/\.xlsx$/i, "").replace(/\.xls$/i, "");
      const robustBaseName = normalizeGermanUmlauts(baseNameRaw);
      let tempFile = null;

      try {
        SpreadsheetApp.getActive().toast(`Konvertiere: ${fileName}`);

        const destinationSheetFile = allFilesInFolder.find(item =>
          item.mimeType === MimeType.GOOGLE_SHEETS &&
          normalizeGermanUmlauts(item.title) === robustBaseName
        );

        let attempts = 0;
        while (attempts < 7) {
          try {
            attempts++;
            const fileBlob = DriveApp.getFileById(file.id).getBlob();
            const resource = {
              title: `temp_${new Date().getTime()}`,
              mimeType: MimeType.GOOGLE_SHEETS
            };

            tempFile = Drive.Files.insert(resource, fileBlob, {
              supportsAllDrives: true,
              convert: true
            });

            if (tempFile) break;
          } catch (err) {
            Logger.log(`Versuch ${attempts} fehlgeschlagen für ${fileName}: ${err.message}`);
            if (attempts >= 7) throw err;
            Utilities.sleep(5000);
          }
        }

        const tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
        const tempSourceSheet = tempSpreadsheet.getSheets()[0];
        const dataRange = tempSourceSheet.getDataRange();
        const rawValues = dataRange.getValues();
        const displayValues = dataRange.getDisplayValues();

        const sanitizedData = rawValues.map((row, rowIndex) => {
          return row.map((cell, columnIndex) => {
            return convertCellForPlainText_(cell, displayValues[rowIndex][columnIndex]);
          });
        });

        if (!destinationSheetFile) {
          tempSourceSheet.clear();

          if (sanitizedData.length > 0 && sanitizedData[0].length > 0) {
            const targetRange = tempSourceSheet.getRange(1, 1, sanitizedData.length, sanitizedData[0].length);
            targetRange.setNumberFormat("@");
            targetRange.setValues(sanitizedData);
          }

          const tempFileParents = Drive.Files.get(tempFile.id, { supportsAllDrives: true }).parents;
          const parentIdsToRemove = tempFileParents.map(parent => parent.id).join(",");

          Drive.Files.update({ title: baseNameRaw }, tempFile.id, null, {
            addParents: folderId,
            removeParents: parentIdsToRemove,
            supportsAllDrives: true
          });

          finalFileIds.set(robustBaseName, tempFile.id);
          tempFile = null;
        } else {
          const destSpreadsheet = SpreadsheetApp.openById(destinationSheetFile.id);
          const destSheet = destSpreadsheet.getSheets()[0];

          destSheet.clear();

          if (sanitizedData.length > 0 && sanitizedData[0].length > 0) {
            const targetRange = destSheet.getRange(1, 1, sanitizedData.length, sanitizedData[0].length);
            targetRange.setNumberFormat("@");
            targetRange.setValues(sanitizedData);
          }

          finalFileIds.set(robustBaseName, destinationSheetFile.id);
        }

        processedCount++;
      } catch (e) {
        Logger.log(`Fehler bei ${fileName}: ${e.message}\n${e.stack}`);
      } finally {
        if (tempFile) {
          try {
            Drive.Files.trash(tempFile.id, { supportsAllDrives: true });
          } catch (e) {
            Logger.log(`Temp-Datei konnte nicht gelöscht werden: ${e.message}`);
          }
        }
      }
    }

    const outputIdsArray = fileNamesValues.map(row => {
      const name = convertNormalizeText_(row);
      let baseName = name;
      if (name.toLowerCase().endsWith(".xlsx")) baseName = name.slice(0, -5);
      else if (name.toLowerCase().endsWith(".xls")) baseName = name.slice(0, -4);

      return [finalFileIds.get(normalizeGermanUmlauts(baseName)) || ""];
    });

    sheet.getRange("E4:E17").setValues(outputIdsArray);
    SpreadsheetApp.getActive().toast(`${processedCount} Dateien verarbeitet.`, "Fertig", 5);
  } catch (e) {
    ui.alert("Fehler", `Details: ${e.message}`, ui.ButtonSet.OK);
  }
}
