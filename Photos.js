/**
 * Inserts =IMAGE(...) formulas into specified tabs: "Failures" and "Overall % upto 4th Sem".
 * Automatically detects the "TMI ID" column header (or falls back to Column B).
 * Writes explicit =IMAGE(...) formulas so IMPORTRANGE works across sheets.
 */
function processSpecificSheets() {
  // Folder ID containing student photos
  var FOLDER_ID = "1f-wGlnXvnXua-g5oxUunZTU3qk4NZHzN"; 
  var TARGET_SHEETS = ["Failures", "Overall % upto 4th Sem"];
  var HEADER_ROW = 2;   // Header row
  var START_ROW = 3;    // Data start row

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Step 1: Index all image files across folder and subfolders
  Logger.log("Indexing all image files from Google Drive folder ID: " + FOLDER_ID);
  var fileList = getAllImagesInFolderHierarchy(FOLDER_ID);
  Logger.log("Total image files indexed: " + fileList.length);

  // Step 2: Loop through target sheets
  for (var k = 0; k < TARGET_SHEETS.length; k++) {
    var sheetName = TARGET_SHEETS[k];
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log("Sheet not found: '" + sheetName + "'. Check tab name for exact spelling/spaces.");
      continue;
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < START_ROW) {
      Logger.log("No data starting from row " + START_ROW + " in sheet: " + sheetName);
      continue;
    }

    // Dynamically locate the "TMI ID" column, defaulting to Column B (col 2) if not found by header name
    var tmiCol = findTmiIdColumn(sheet, HEADER_ROW);
    Logger.log("[" + sheetName + "] Using Column " + tmiCol + " for TMI ID search.");

    // Locate or create the "Photos" column
    var photoCol = findOrCreatePhotosColumn(sheet, HEADER_ROW, lastRow);

    // Read TMI IDs from detected column (Row 3 onwards)
    var tmiIds = sheet.getRange(START_ROW, tmiCol, lastRow - START_ROW + 1, 1).getValues();

    // Read current formulas in the Photos column to avoid re-processing
    var currentPhotoRange = sheet.getRange(START_ROW, photoCol, lastRow - START_ROW + 1, 1);
    var currentFormulas = currentPhotoRange.getFormulas();

    for (var i = 0; i < tmiIds.length; i++) {
      var rawTmiId = String(tmiIds[i][0]).trim();
      var targetRow = START_ROW + i;

      if (!rawTmiId) continue;

      var currentFormula = String(currentFormulas[i][0]).trim();

      // Skip if valid formula already present
      if (currentFormula.toUpperCase().startsWith("=IMAGE")) {
        Logger.log("[" + sheetName + "] Row " + targetRow + " -> Already has =IMAGE formula. Skipping.");
        continue;
      }

      // Clean TMI ID for search matching
      var cleanTmiId = rawTmiId.replace(/[^0-9a-zA-Z]/g, '');
      var searchKey = cleanTmiId || rawTmiId;

      // Search matching file from indexed folder array
      var matchedFile = fileList.find(function(file) {
        var cleanFileName = file.name.replace(/[^0-9a-zA-Z]/g, '').toLowerCase();
        var cleanKey = searchKey.toLowerCase();
        return cleanFileName.indexOf(cleanKey) !== -1;
      });

      // Direct fallback Drive search
      if (!matchedFile && searchKey.length >= 3) {
        matchedFile = searchDirectDriveFile(searchKey);
      }

      if (matchedFile) {
        // Construct standard =IMAGE() formula using Google's direct web thumbnail URL
        var imageUrl = "https://drive.google.com/thumbnail?id=" + matchedFile.id + "&sz=w500";
        var imageFormula = '=IMAGE("' + imageUrl + '", 1)';

        sheet.getRange(targetRow, photoCol).setFormula(imageFormula);
        Logger.log("[" + sheetName + "] Row " + targetRow + " -> Inserted =IMAGE() formula for TMI ID: " + searchKey);
      } else {
        sheet.getRange(targetRow, photoCol).setValue("Image Not Found");
        Logger.log("[" + sheetName + "] Row " + targetRow + " -> No image found for TMI ID: " + searchKey);
      }
    }
  }
}

/**
 * Dynamically finds the column index containing "TMI ID" in the header row.
 * Defaults to Column 2 (Column B) if no explicit header matches.
 */
function findTmiIdColumn(sheet, headerRow) {
  var lastColumn = sheet.getLastColumn();
  if (lastColumn >= 1) {
    var headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
    for (var col = 0; col < headers.length; col++) {
      var headerText = String(headers[col]).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
      if (headerText === "tmiid" || headerText === "tmi" || headerText === "imuid") {
        return col + 1; // 1-based index
      }
    }
  }
  return 2; // Default to Column B
}

/**
 * Finds existing "Photos" column or creates a new one in the first empty column
 */
function findOrCreatePhotosColumn(sheet, headerRow, lastRow) {
  var lastColumn = sheet.getLastColumn();

  if (lastColumn >= 3) {
    var headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
    for (var col = 1; col < headers.length; col++) {
      if (String(headers[col]).trim().toLowerCase() === "photos") {
        return col + 1;
      }
    }
  }

  var newCol = getNextEmptyColumnForSheet(sheet, headerRow, lastRow);
  var headerCell = sheet.getRange(headerRow, newCol);
  headerCell.setValue("Photos");
  headerCell.setFontWeight("bold");
  return newCol;
}

/**
 * Direct fallback search for missing individual files in Drive
 */
function searchDirectDriveFile(searchKey) {
  try {
    var query = "title contains '" + searchKey + "' and mimeType contains 'image/' and trashed = false";
    var files = DriveApp.searchFiles(query);
    if (files.hasNext()) {
      var file = files.next();
      return { name: file.getName(), id: file.getId() };
    }
  } catch (e) {
    Logger.log("Direct search error: " + e.message);
  }
  return null;
}

/**
 * Recursively fetches all image files across subfolders
 */
function getAllImagesInFolderHierarchy(rootFolderId) {
  var fileList = [];
  var folderIds = [rootFolderId];
  collectSubfolderIds(rootFolderId, folderIds);

  for (var i = 0; i < folderIds.length; i++) {
    var query = "'" + folderIds[i] + "' in parents and mimeType contains 'image/' and trashed = false";
    var files = DriveApp.searchFiles(query);
    while (files.hasNext()) {
      var file = files.next();
      fileList.push({
        name: file.getName(),
        id: file.getId()
      });
    }
  }

  return fileList;
}

/**
 * Collects subfolder IDs recursively
 */
function collectSubfolderIds(parentFolderId, folderIds) {
  try {
    var query = "'" + parentFolderId + "' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false";
    var subfolders = DriveApp.searchFiles(query);
    while (subfolders.hasNext()) {
      var subfolder = subfolders.next();
      var subfolderId = subfolder.getId();
      folderIds.push(subfolderId);
      collectSubfolderIds(subfolderId, folderIds);
    }
  } catch (e) {
    Logger.log("Error querying subfolders: " + e.message);
  }
}

/**
 * Finds next available empty column in the sheet
 */
function getNextEmptyColumnForSheet(sheet, startRow, endRow) {
  var lastColumn = sheet.getLastColumn();
  if (lastColumn < 2) return 3;

  for (var col = 3; col <= lastColumn; col++) {
    var rangeValues = sheet.getRange(startRow, col, endRow - startRow + 1, 1).getValues();
    var isColumnEmpty = rangeValues.every(function(row) {
      return row[0] === "" || row[0] === null || row[0] === undefined;
    });

    if (isColumnEmpty) {
      return col;
    }
  }

  return lastColumn + 1;
}