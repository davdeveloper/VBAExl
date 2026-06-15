
// /**
//  * EXPORT SORTIMENTSWÜNSCHE FOOD (Clone Strategy)
//  * - Replicates 'ExportSortimentswünscheFoodGesamt'
//  * - Generates 3 separate PDFs (Food Sortimentswünsche, Änderungswünsche, Allgemeine)
//  * - Footer Gap: 40 Rows.
//  * - Headers: Auto-Expanded.
//  * - Layout: Sidebar + Data + Footer.
//  */

// function export_sortimentswuensche_new() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const ui = SpreadsheetApp.getUi();

//   // --- 1. CONFIGURATION ---
  
//   // SHARED SETTINGS
//   // The filename KW comes from 'Allgemeine Sortimentswünsche'!J7 in the VBA.
//   const filenameSheet = ss.getSheetByName("Allgemeine Sortimentswünsche");
//   const kwValue = filenameSheet ? filenameSheet.getRange("J7").getValue() : "XX";

//   const folderUrl = ss.getSheetByName("EXPORT").getRange("D4").getValue();
//   let folder = getFolder(folderUrl, ss);

//   // SHEET 1: Food Sortimentswünsche
//   // VBA sorts A:R. We keep cols 1-18.
//   const configFood = {
//     name: "Food Sortimentswünsche",
//     keepStartCol: 1, keepEndCol: 18, 
//     week53Row: null, // No specific week 53 hiding logic found in this sub, usually data driven.
//     sidebarText: "", // Left blank as per PDF sample, but column exists for layout consistency.
//     filename: `02_Statistik_Food_Sortimentswünsche_KW${kwValue}_DE_Part1`
//   };

//   // SHEET 2: Änderungswünsche Sortierung
//   // VBA sorts A:R. Keep cols 1-18.
//   const configAenderung = {
//     name: "Änderungswünsche Sortierung",
//     keepStartCol: 1, keepEndCol: 18,
//     week53Row: null,
//     sidebarText: "",
//     filename: `02_Statistik_Food_Sortimentswünsche_KW${kwValue}_DE_Part2`
//   };

//   // SHEET 3: Allgemeine Sortimentswünsche
//   // VBA range A:K. Keep cols 1-11.
//   const configAllg = {
//     name: "Allgemeine Sortimentswünsche",
//     keepStartCol: 1, keepEndCol: 11,
//     week53Row: null,
//     sidebarText: "",
//     filename: `02_Statistik_Food_Sortimentswünsche_KW${kwValue}_DE_Part3`
//   };

//   // --- EXECUTE ---
  
//   const sheets = [configFood, configAenderung, configAllg];
  
//   // Ask 53 week question only if we need to implement logic (kept for consistency)
//   // The VBA for this specific sub didn't explicitly hide row 53/56, but we'll keep the object structure.
  
//   sheets.forEach(config => {
//     const sheet = ss.getSheetByName(config.name);
//     if (!sheet) {
//       console.warn(`Sheet ${config.name} not found. Skipping.`);
//       return;
//     }
    
//     // Pass 'kwValue' directly to footer logic
//     processSheetClone(ss, sheet, folder, config, kwValue);
//   });

//   ss.getSheetByName("EXPORT").activate();
//   ui.alert("Erfolg!", "Die Dateien wurden erfolgreich erstellt!", ui.ButtonSet.OK);
// }

// /**
//  * CORE LOGIC: CLONE & TRIM
//  */
// function processSheetClone(ss, sourceSheet, folder, config, kwValue) {
  
//   // 1. Clone the Sheet
//   const tempSheet = sourceSheet.copyTo(ss);
//   tempSheet.setName("Temp_" + new Date().getTime());
  
//   try {
//     // 2. Remove Drawings
//     const drawings = tempSheet.getDrawings();
//     drawings.forEach(d => d.remove());

//     // 3. Trim Columns (Right then Left)
//     const maxCols = tempSheet.getMaxColumns();
//     if (maxCols > config.keepEndCol) {
//       tempSheet.deleteColumns(config.keepEndCol + 1, maxCols - config.keepEndCol);
//     }
//     if (config.keepStartCol > 1) {
//       tempSheet.deleteColumns(1, config.keepStartCol - 1);
//     }

//     // 4. Insert Sidebar (New Col 1)
//     tempSheet.insertColumnBefore(1);
//     tempSheet.setColumnWidth(1, 100);

//     // 5. Sidebar Text (If any)
//     if (config.sidebarText) {
//       const sidebarRange = tempSheet.getRange(5, 1, 30, 1);
//       sidebarRange.merge();
//       sidebarRange.setValue(config.sidebarText);
//       sidebarRange.setVerticalAlignment("top").setHorizontalAlignment("left").setFontSize(8).setWrap(true);
//     }

//     // 6. Footer Logic
//     // Find last actual data row. Since we cloned, we scan up from bottom or assume a safe buffer.
//     // For these reports, data is dynamic. We'll use getLastRow() on the TEMP sheet.
//     const lastDataRow = tempSheet.getLastRow(); 
//     const footerRow = lastDataRow + 40; // 40 Rows gap

//     const dataCols = config.keepEndCol - config.keepStartCol + 1; 

//     // Calculate Footer Positions
//     const midIndex = Math.ceil(dataCols / 2);
    
//     const colErsteller = 2; // Always first data column (Col B)
//     const colKW = 1 + midIndex;

//     // Set Text
//     const cellErsteller = tempSheet.getRange(footerRow, colErsteller);
//     cellErsteller.setValue("Ersteller: Kundenservice Einkauf");
//     cellErsteller.setHorizontalAlignment("left");
//     cellErsteller.setFontFamily("Arial").setFontSize(10).setFontWeight("normal");

//     const cellKW = tempSheet.getRange(footerRow, colKW);
//     cellKW.setValue("Statistik KW " + kwValue);
//     cellKW.setHorizontalAlignment("center");
//     cellKW.setFontFamily("Arial").setFontSize(10).setFontWeight("normal");

//     // 7. Auto-Expand Headers (Simulate Double Click)
//     // Runs on rows 1-10
//     const sheetId = tempSheet.getSheetId();
//     const requests = [
//       {
//         "autoResizeDimensions": {
//           "dimensions": { 
//             "sheetId": sheetId, 
//             "dimension": "ROWS", 
//             "startIndex": 0, 
//             "endIndex": 10 
//           }
//         }
//       }
//     ];
//     Sheets.Spreadsheets.batchUpdate({ requests: requests }, ss.getId());
//     SpreadsheetApp.flush();

//     // 8. Export
//     const filename = `${config.filename}.pdf`;
    
//     // Explicit Print Range to include footer
//     // Range: A1 : (LastCol)(FooterRow)
//     const totalCols = dataCols + 1; // Data + Sidebar
//     const printRange = "r1=0&c1=0&r2=" + footerRow + "&c2=" + totalCols;
    
//     exportSheetToPdf(ss, tempSheet, folder, filename, printRange);

//   } catch (e) {
//     console.error(e);
//     SpreadsheetApp.getUi().alert("Error: " + e.message);
//   } finally {
//     ss.deleteSheet(tempSheet);
//   }
// }

// function exportSheetToPdf(ss, sheet, folder, filename, rangeParam) {
//   const ssId = ss.getId();
//   const sheetId = sheet.getSheetId();
  
//   const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
//     "?format=pdf" +
//     "&size=A4" +
//     "&portrait=false" +      // Landscape often better for wide sorting sheets (A-R), change to true if Portrait required
//     "&fitw=true" +
//     "&gridlines=false" +
//     "&printtitle=false" +
//     "&sheetnames=false" +
//     "&pagenumbers=false" +
//     "&gid=" + sheetId + 
//     "&" + rangeParam; 

//   const token = ScriptApp.getOAuthToken();
//   const blob = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } }).getBlob().setName(filename);
//   folder.createFile(blob);
// }

// function getFolder(url, ss) {
//   try {
//     if (url && String(url).includes("drive.google.com")) {
//       const match = url.match(/[-\w]{25,}/);
//       if (match) return DriveApp.getFolderById(match[0]);
//     }
//   } catch (e) {}
//   const parents = DriveApp.getFileById(ss.getId()).getParents();
//   return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
// }

// /**
//  * SERVER SIDE: Generate 3 Parts & Send to Client for Merging
//  */

// function openSortimentswunscheDialog() {
//   const html = HtmlService.createHtmlOutputFromFile('sortimentswunche_pdf_gen')
//       .setWidth(500)
//       .setHeight(400);
//   SpreadsheetApp.getUi().showModalDialog(html, 'Generating & Merging PDFs...');
// }

// /**
//  * Generates all 3 PDFs as Base64 and returns them to the client
//  */
// function getPdfDataForMerging() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const year = new Date().getFullYear();

//   // --- CONFIGURATION ---
//   const filenameSheet = ss.getSheetByName("Allgemeine Sortimentswünsche");
//   const kwValue = filenameSheet ? filenameSheet.getRange("J7").getValue() : "XX";
  
//   const folderUrl = ss.getSheetByName("EXPORT").getRange("D4").getValue();
//   let folder = getFolder(folderUrl, ss);
//   const folderId = folder.getId();

//   // 1. Food (Part 1)
//   const configFood = {
//     name: "Food Sortimentswünsche",
//     keepStartCol: 1, keepEndCol: 18, 
//     titleText: `Mehrfach angeforderte Sortimentswünsche Food ${year}`,
//     filename: `temp_part1`
//   };

//   // 2. Änderungswünsche (Part 2)
//   const configAenderung = {
//     name: "Änderungswünsche Sortierung",
//     keepStartCol: 1, keepEndCol: 18,
//     titleText: `Änderungswünsche Sortierung Food ${year}`,
//     filename: `temp_part2`
//   };

//   // 3. Allgemeine (Part 3)
//   const configAllg = {
//     name: "Allgemeine Sortimentswünsche",
//     keepStartCol: 1, keepEndCol: 11,
//     titleText: `Allgemeine Sortimentswünsche Food ${year}`,
//     filename: `temp_part3`
//   };

//   // --- PROCESS ALL 3 PARTS ---
//   const blob1 = processSheetClone(ss, configFood, kwValue);
//   const blob2 = processSheetClone(ss, configAenderung, kwValue);
//   const blob3 = processSheetClone(ss, configAllg, kwValue);

//   if (!blob1 || !blob2 || !blob3) {
//     throw new Error("Could not generate one or more sheets. Check sheet names.");
//   }

//   return {
//     part1: Utilities.base64Encode(blob1.getBytes()),
//     part2: Utilities.base64Encode(blob2.getBytes()),
//     part3: Utilities.base64Encode(blob3.getBytes()),
//     filename: `02_Statistik_Food_Sortimentswünsche_KW${kwValue}_DE.pdf`,
//     folderId: folderId
//   };
// }

// function saveMergedPdfToDrive(base64Data, filename, folderId) {
//   const folder = DriveApp.getFolderById(folderId);
//   const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), MimeType.PDF, filename);
//   folder.createFile(blob);
//   return "Done";
// }

// // --- SHARED HELPER: CLONE & FORMAT ---

// function processSheetClone(ss, config, kwValue) {
//   const sourceSheet = ss.getSheetByName(config.name);
//   if (!sourceSheet) return null;

//   const tempSheet = sourceSheet.copyTo(ss);
//   tempSheet.setName("Temp_" + new Date().getTime());
  
//   try {
//     // 1. Unfreeze
//     tempSheet.setFrozenRows(0);
//     tempSheet.setFrozenColumns(0);

//     // 2. Remove Drawings
//     const drawings = tempSheet.getDrawings();
//     drawings.forEach(d => d.remove());

//     // 3. Trim Columns
//     const maxCols = tempSheet.getMaxColumns();
//     if (maxCols > config.keepEndCol) {
//       tempSheet.deleteColumns(config.keepEndCol + 1, maxCols - config.keepEndCol);
//     }
//     if (config.keepStartCol > 1) {
//       tempSheet.deleteColumns(1, config.keepStartCol - 1);
//     }

//     // 4. Insert Sidebar
//     tempSheet.insertColumnBefore(1);
//     tempSheet.setColumnWidth(1, 20);

//     // 5. Insert Title
//     tempSheet.insertRowsBefore(1, 4);
//     const dataCols = config.keepEndCol - config.keepStartCol + 1;
    
//     // Robust Title Merge (Full Width)
//     const titleRange = tempSheet.getRange(2, 2, 1, dataCols);
//     titleRange.merge();
//     titleRange.setValue(config.titleText);
//     titleRange.setFontWeight("bold");
//     titleRange.setFontSize(14);
//     titleRange.setFontFamily("Arial");
//     titleRange.setHorizontalAlignment("center");
//     titleRange.setVerticalAlignment("middle");
//     titleRange.setBackground(null); 
//     tempSheet.setRowHeight(2, 40); 

//     // 6. SMART FOOTER LOGIC (Bottom of Page Calculation)
//     const ROWS_PER_PAGE = 38; // Adjusted for Title insertion
//     const lastDataRow = tempSheet.getLastRow();
    
//     // Calculate which page number the data ends on (1, 2, etc)
//     const pageNum = Math.ceil(lastDataRow / ROWS_PER_PAGE);
    
//     // Target the bottom of that page
//     let targetFooterRow = (pageNum * ROWS_PER_PAGE) - 2;

//     // Safety: If data ends too close to the bottom (within 5 rows), push to next page
//     if ((targetFooterRow - lastDataRow) < 5) {
//       targetFooterRow += ROWS_PER_PAGE;
//     }

//     // Insert Footer
//     const footerRow = targetFooterRow;
//     const midIndex = Math.ceil(dataCols / 2);
//     const colErsteller = 2; 
//     const colKW = 1 + midIndex;

//     const cellErsteller = tempSheet.getRange(footerRow, colErsteller);
//     cellErsteller.setValue("Ersteller: Kundenservice Einkauf");
//     cellErsteller.setHorizontalAlignment("left");
//     cellErsteller.setFontFamily("Arial").setFontSize(10).setFontWeight("normal");

//     const cellKW = tempSheet.getRange(footerRow, colKW);
//     cellKW.setValue("Statistik KW " + kwValue);
//     cellKW.setHorizontalAlignment("center");
//     cellKW.setFontFamily("Arial").setFontSize(10).setFontWeight("normal");

//     // Clean Footer
//     const footerCleanRange = tempSheet.getRange(footerRow, 2, 1, dataCols);
//     footerCleanRange.setBackground(null);
//     footerCleanRange.setBorder(false, false, false, false, false, false);

//     SpreadsheetApp.flush();

//     // 7. Auto-Expand Headers
//     const sheetId = tempSheet.getSheetId();
//     const requests = [{
//       "autoResizeDimensions": {
//         "dimensions": { "sheetId": sheetId, "dimension": "ROWS", "startIndex": 0, "endIndex": 20 }
//       }
//     }];
//     Sheets.Spreadsheets.batchUpdate({ requests: requests }, ss.getId());
//     SpreadsheetApp.flush();

//     // 8. Generate Blob
//     const totalCols = dataCols + 1; 
//     const printRange = "r1=0&c1=0&r2=" + footerRow + "&c2=" + totalCols;
    
//     return generatePdfBlob(ss, tempSheet, config.filename, printRange);

//   } catch (e) {
//     console.error(e);
//     return null;
//   } finally {
//     ss.deleteSheet(tempSheet);
//   }
// }

// function generatePdfBlob(ss, sheet, filename, rangeParam) {
//   const ssId = ss.getId();
//   const sheetId = sheet.getSheetId();
//   const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
//     "?format=pdf&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=false&gid=" + sheetId + "&" + rangeParam; 
//   const token = ScriptApp.getOAuthToken();
//   return UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } }).getBlob().setName(filename + ".pdf");
// }

// function getFolder(url, ss) {
//   try {
//     if (url && String(url).includes("drive.google.com")) {
//       const match = url.match(/[-\w]{25,}/);
//       if (match) return DriveApp.getFolderById(match[0]);
//     }
//   } catch (e) {}
//   const parents = DriveApp.getFileById(ss.getId()).getParents();
//   return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
// }


/**
 * 🖨️ EXPORT SORTIMENTSWÜNSCHE GESAMT (Merged PDF)
 * Creates a single, merged landscape PDF from up to 3 sheets.
 * Dynamically skips sheets that contain no data.
 */

function export_sortimentswuensche_gesamt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetExport = ss.getSheetByName("EXPORT");

  if (!sheetExport) {
    ui.alert("Fehler", "Tabellenblatt 'EXPORT' nicht gefunden.", ui.ButtonSet.OK);
    return;
  }

  // --- 1. RESOLVE FOLDER ---
  const folderUrl = sheetExport.getRange("D4").getValue().toString().trim();
  const folderId = all2ExtractDriveId_(folderUrl);
  try {
    DriveApp.getFolderById(folderId); 
  } catch (e) {
    ui.alert("Ordner Fehler", "Der Zielordner konnte nicht gefunden werden.", ui.ButtonSet.OK);
    return;
  }

  // Get KW explicitly from Allgemeine Sortimentswünsche as per VBA logic
  const allgSheet = ss.getSheetByName("Allgemeine Sortimentswünsche");
  const kwValue = allgSheet ? allgSheet.getRange("J7").getValue() : "XX";

  ss.toast("Prüfe Tabellenblätter... Bitte warten.", "Schritt 1/3", 10);

  // --- 2. FETCH DASHBOARD DATA (B31:D144) ---
  const dashboardData = sheetExport.getRange("B31:D144").getValues();
  function getDashboardText(key) {
    const row = dashboardData.find(r => r[0] === key);
    return row && row[2] ? row[2].toString().trim() : "";
  }

  // --- 3. CONFIGURATIONS ---
  const configs = [
    {
      name: "Food Sortimentswünsche",
      dataRowStart: 6, // Rows below 5 are data
      keepCols: 18,    // Keep A-R
      keys: {
        htl: "Food Sortimentswünsche header left", htc: "Food Sortimentswünsche header center", htr: "Food Sortimentswünsche header right",
        fbl: "Food Sortimentswünsche footer left", fbc: "Food Sortimentswünsche footer center", fbr: "Food Sortimentswünsche footer right"
      }
    },
    {
      name: "Änderungswünsche Sortierung",
      dataRowStart: 6, // Rows below 5 are data
      keepCols: 18,    // Keep A-R
      keys: {
        htl: "Änderungswünsche Sortierung header left", htc: "Änderungswünsche Sortierung header center", htr: "Änderungswünsche Sortierung header right",
        fbl: "Änderungswünsche Sortierung footer left", fbc: "Änderungswünsche Sortierung footer center", fbr: "Änderungswünsche Sortierung footer right"
      }
    },
    {
      name: "Allgemeine Sortimentswünsche",
      dataRowStart: 8, // Rows below 7 are data
      keepCols: 11,    // Keep A-K
      keys: {
        htl: "Allgemeine Sortimentswünsche header left", htc: "Allgemeine Sortimentswünsche header center", htr: "Allgemeine Sortimentswünsche header right",
        fbl: "Allgemeine Sortimentswünsche footer left", fbc: "Allgemeine Sortimentswünsche footer center", fbr: "Allgemeine Sortimentswünsche footer right"
      }
    }
  ];

  const token = ScriptApp.getOAuthToken();
  const pdfParts = [];
  const tempSheetIds = [];

  // --- 4. PREPARE BASE PDFs (Only for sheets with data) ---
  for (const config of configs) {
    const liveSheet = ss.getSheetByName(config.name);
    if (!liveSheet) continue;

    // Check if sheet actually has data (mimicking VBA's check)
    const lastRow = liveSheet.getLastRow();
    if (lastRow < config.dataRowStart) {
      console.log(`Skipping ${config.name} - No Data`);
      continue; 
    }

    // Clone & Trim
    const uniqueSuffix = Math.floor(Math.random() * 1000000);
    const tempSheet = liveSheet.copyTo(ss);
    tempSheet.setName("Temp_Merge_" + uniqueSuffix);
    const sheetId = tempSheet.getSheetId();
    tempSheetIds.push({ id: sheetId, name: tempSheet.getName() });

    try {
      // Clear UI elements
      const drawings = tempSheet.getDrawings();
      if (drawings) drawings.forEach(d => d.remove());

      // Trim extra columns
      const maxCols = tempSheet.getMaxColumns();
      if (maxCols > config.keepCols) tempSheet.deleteColumns(config.keepCols + 1, maxCols - config.keepCols);

      SpreadsheetApp.flush(); 
      Utilities.sleep(1500); 

      // Fetch Base PDF (LANDSCAPE mode via portrait=false)
      const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
        `format=pdf&gid=${sheetId}` +
        `&size=A4&portrait=false&fitw=true&gridlines=false` +
        `&printtitle=false&sheetnames=false&pagenum=UNDEFINED` +
        `&horizontal_alignment=CENTER` + 
        `&top_margin=0.787&bottom_margin=0.709&left_margin=0.4&right_margin=0.4`;

      const responseAPI = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });

      if (responseAPI.getResponseCode() === 200) {
        const base64Pdf = Utilities.base64Encode(responseAPI.getBlob().getBytes());
        
        pdfParts.push({
          sheetName: config.name,
          base64: base64Pdf,
          layout: {
            htl: getDashboardText(config.keys.htl),
            htc: getDashboardText(config.keys.htc),
            htr: getDashboardText(config.keys.htr),
            fbl: getDashboardText(config.keys.fbl),
            fbc: getDashboardText(config.keys.fbc),
            fbr: getDashboardText(config.keys.fbr)
          }
        });
      } else {
        throw new Error(`PDF-Export für "${config.name}" fehlgeschlagen: HTTP ${responseAPI.getResponseCode()}`);
      }
    } catch (e) {
      console.error(e);
    }
  }

  // --- 5. OPEN HTML FRONTEND ---
  if (pdfParts.length > 0) {
    const finalFilename = `02_Statistik_Food_Sortimentswünsche_KW${kwValue}_DE.pdf`;
    
    const htmlTemplate = HtmlService.createTemplateFromFile('multi-page-support-pdf-lib');
    htmlTemplate.payload = JSON.stringify({ 
      parts: pdfParts, 
      folderId: folderId, 
      filename: finalFilename,
      tempSheets: tempSheetIds 
    });
    
    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(500)
      .setHeight(300)
      .setTitle('Merge & PDF Renderer');
      
    ui.showModalDialog(htmlOutput, 'Generiere zusammengeführtes PDF...');
  } else {
    ui.alert("Hinweis", "Keine der ausgewählten Tabellen enthält Daten.", ui.ButtonSet.OK);
    // Cleanup any empty temp sheets just in case
    for (const ts of tempSheetIds) {
      const sheet = ss.getSheetByName(ts.name);
      if (sheet && ss.getSheets().length > 1) ss.deleteSheet(sheet);
    }
  }
}

/**
 * BACKEND CALLBACK: Saves the single merged PDF and cleans up.
 */
function saveMergedPdfToDrive_new(base64Data, filename, folderId, tempSheets) {
  const folder = DriveApp.getFolderById(folderId);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Save single PDF
    const decodedBytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedBytes, 'application/pdf', filename);
    folder.createFile(blob);

    return true;
  } finally {
    // Clean up Temp Sheets even if saving fails.
    for (const ts of tempSheets || []) {
      const sheet = ss.getSheetByName(ts.name);
      if (sheet && ss.getSheets().length > 1) ss.deleteSheet(sheet);
    }

    all2CleanUpTempSheets_();
  }
}

