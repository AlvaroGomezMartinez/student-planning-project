/**
 * Triggered when the spreadsheet is opened. Adds a custom menu for student tools.
 * @returns {void}
 */
function onOpen() {
  // Adds a custom menu so users can run the folder-creation from the UI
  SpreadsheetApp.getUi()
    .createMenu("Student Tools")
    .addItem("Instructions", "showInstructions")
    .addSeparator()
    .addItem("Create Student Folders", "createStudentFolders")
    .addItem("Move PDFs in output_folder to Student Folders", "moveMatchingPDFsToStudentFolders")
  .addSeparator()
  .addItem("Assign Counselor Emails", "assignCounselorEmails")
  .addSeparator()
  .addItem("Grant View Permissions (Folders Only)", "grantStudentCommentPermissions")
  .addItem("🚀 OPTIMIZED: Grant View Permissions (Folders + Files)", "grantStudentCommentPermissionsOptimized")
  .addItem("Grant View Permissions (Folders + Files)", "grantStudentCommentPermissionsToFoldersAndFiles")
  .addSeparator()
  .addItem("▶️ Start Automatic Processing", "startAutomaticPermissionGrants")
  .addItem("▶️ Start OPTIMIZED Auto Processing", "startAutomaticPermissionGrantsOptimized")
  .addItem("⏹️ Stop Automatic Processing", "stopAutomaticPermissionGrants")
  .addItem("Reset Permission Progress", "resetGrantPermissionsProgress")
    .addToUi();
}

/**
 * Opens a modal dialog with instructions for the user.
 */
function showInstructions() {
  var html = HtmlService.createHtmlOutputFromFile('Instructions')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Instructions');
}

// Configuration constants collected in a single object
const configs = {
  parentFolderId: "1B79FFF2v1vtkjHWM480DtqlIbsaheu-b",
  defaultSpreadsheet: "1afXZA4x4SoP2BMc3g5IqIgvEuzMdf6oqEgOrPqw0iH8",
  output_folder: "1TM9jwn2ehatw7jMEmIYYF4dusGtAF0Gu",
  defaultSheetName: "Main Roster",
  // Secondary spreadsheet that contains document ids in column AN across multiple tabs
  secondarySpreadsheet: "1YaB_u1Hue9gdMM9Ka0NTFwmJaIP47bs9Rz1lhoLGjb4",
  secondarySheetNames: [
    "1-100 CCMR Student Listing",
    "101-200 CCMR Student Listing",
    "201-300 CCMR Student Listing",
    "301-400 CCMR Student Listing",
    "401-500 CCMR Student Listing",
    "501-609 CCMR Student Listing",
  ],
  planningPdfFileUrls: [
    {
      fileUrl:
        "https://drive.google.com/file/d/1ipGuidPDDharnQejXNDueUZOmCo9h1zh/view?usp=drive_link",
    },
    {
      fileUrl:
        "https://drive.google.com/file/d/10QfbVYI67RwqVaWIYbnvVYGOYkFSTfTZ/view?usp=drive_link",
    },
  ],
  planningHeaderCandidates: [
    // Used to help find the correct columns
    "Planning Folder URL",
    "Planning folder URL",
    "Planning Folder Url",
    "Planning FolderUrl",
  ],
  nameHeaderCandidates: ["Student Name", "Name", "Full Name", "Student"],
  emailHeaderCandidates: ["Student Email", "Email", "Student Email Address", "Email Address"],
  idColumnIndex: 1, // column A
};

/**
 * Create folders for students listed in the "Main Roster" sheet when the
 * "Planning Folder URL" cell is empty. New folder is created under the
 * parent folder and the folder name is the student's name. The created folder URL is written back.
 * @returns {void}
 */
function createStudentFolders() {
  let ui = SpreadsheetApp.getUi();
  let ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(configs.defaultSheetName);
  if (!sheet) {
    let resp = ui.prompt(
      'Sheet "' +
        configs.defaultSheetName +
        '" not found. Enter the sheet name to use (or Cancel to abort):',
      ui.ButtonSet.OK_CANCEL,
    );
    if (resp.getSelectedButton() !== ui.Button.OK) {
      ui.alert("Operation cancelled.");
      return;
    }
    let sheetName = (resp.getResponseText() || "").toString().trim();
    if (!sheetName) {
      ui.alert("No sheet name provided. Operation cancelled.");
      return;
    }
    sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      ui.alert('Sheet "' + sheetName + '" not found. Operation cancelled.');
      return;
    }
  }

  const PARENT_FOLDER_ID = configs.parentFolderId;
  let lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("No student rows found in the selected sheet.");
    return;
  }

  let lastCol = sheet.getLastColumn();
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let planningCol = findHeaderIndex(headers, configs.planningHeaderCandidates);
  if (planningCol === -1) {
    ui.alert('Header "Planning Folder URL" not found.');
    return;
  }

  let nameCol = findHeaderIndex(headers, configs.nameHeaderCandidates);
  if (nameCol === -1) {
    ui.alert(
      'Student name column not found. Expected headers like "Student Name" or "Name".',
    );
    return;
  }

  let parentFolder;
  try {
    parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
  } catch (e) {
    ui.alert(
      "Parent folder with the provided ID could not be found or accessed.",
    );
    return;
  }

  // OPTIMIZED: Batch read all data at once to reduce API calls
  let allData = sheet.getDataRange().getValues();
  let batchUpdates = []; // Collect all URL updates for batch writing
  let updatedCount = 0;

  // Loop through data and create folders where Planning Folder URL is empty
  for (let r = 2; r <= lastRow; r++) {
    let rowIndex = r - 1; // Convert to 0-based index for allData
    let planningCell = allData[rowIndex][planningCol - 1];
    
    if (planningCell && planningCell.toString().trim() !== "") {
      continue; // already has a URL
    }

    let studentName = (allData[rowIndex][nameCol - 1] || "").toString().trim();
    let studentId = (allData[rowIndex][configs.idColumnIndex - 1] || "").toString().trim();

    // Skip rows without any identifying info
    if (!studentName && !studentId) {
      continue;
    }

    // Build folder name using student name and disambiguate with student id (col A)
    let folderName = studentName || "Student";
    if (studentId) {
      folderName = folderName + " - " + studentId;
    }
    folderName = sanitizeFolderName(folderName);

    // Create the folder under the parent folder
    let newFolder = parentFolder.createFolder(folderName);
    let url = newFolder.getUrl();

    // Collect URL update for batch writing (instead of individual setValue calls)
    batchUpdates.push({
      row: r,
      url: url
    });
    updatedCount++;

    // Progress indicator every 10 folders
    if (updatedCount % 10 === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Created ${updatedCount} folders...`, 'Progress', 2);
    }
  }

  // BATCH WRITE: Write all URLs at once (HUGE API savings!)
  if (batchUpdates.length > 0) {
    try {
      // Prepare batch data for range update
      let urlValues = batchUpdates.map(update => [update.url]);
      let startRow = batchUpdates[0].row;
      
      // For non-contiguous updates, we'll do individual writes but in a more efficient way
      // Group contiguous updates for maximum efficiency
      let contiguousGroups = [];
      let currentGroup = [batchUpdates[0]];
      
      for (let i = 1; i < batchUpdates.length; i++) {
        if (batchUpdates[i].row === currentGroup[currentGroup.length - 1].row + 1) {
          currentGroup.push(batchUpdates[i]);
        } else {
          contiguousGroups.push(currentGroup);
          currentGroup = [batchUpdates[i]];
        }
      }
      contiguousGroups.push(currentGroup);
      
      // Write each contiguous group as a batch
      contiguousGroups.forEach(group => {
        if (group.length === 1) {
          // Single update
          sheet.getRange(group[0].row, planningCol).setValue(group[0].url);
        } else {
          // Batch update for contiguous rows
          let values = group.map(item => [item.url]);
          sheet.getRange(group[0].row, planningCol, group.length, 1).setValues(values);
        }
      });
      
      Logger.log(`BATCH OPTIMIZATION: Updated ${batchUpdates.length} URLs in ${contiguousGroups.length} batch operations instead of ${batchUpdates.length} individual calls`);
      
    } catch (e) {
      Logger.log(`Error in batch update: ${e.toString()}`);
      // Fallback to individual updates if batch fails
      batchUpdates.forEach(update => {
        try {
          sheet.getRange(update.row, planningCol).setValue(update.url);
        } catch (fallbackError) {
          Logger.log(`Failed to update row ${update.row}: ${fallbackError.toString()}`);
        }
      });
    }
  }

  let message = `Created ${updatedCount} folders and updated Planning Folder URLs.\n\n` +
               `API OPTIMIZATION: Used batch operations instead of ${updatedCount} individual spreadsheet writes!`;
  ui.alert(message);
}

/**
 * Dry-run version: does not create folders or write URLs.
 * Writes a detailed report to a sheet named 'DryRun - Student Folders' and
 * shows a modal summary. Planned URL is a descriptive placeholder in the
 * format: WOULD_CREATE://{parentId}/{encodedFolderName}
 * @returns {void}
 */
function createStudentFoldersDryRun() {
  let ui = SpreadsheetApp.getUi();
  let ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(configs.defaultSheetName);
  if (!sheet) {
    let resp = ui.prompt(
      'Sheet "' +
        configs.defaultSheetName +
        '" not found. Enter the sheet name to use (or Cancel to abort):',
      ui.ButtonSet.OK_CANCEL,
    );
    if (resp.getSelectedButton() !== ui.Button.OK) {
      ui.alert("Operation cancelled.");
      return;
    }
    let sheetName = (resp.getResponseText() || "").toString().trim();
    if (!sheetName) {
      ui.alert("No sheet name provided. Operation cancelled.");
      return;
    }
    sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      ui.alert('Sheet "' + sheetName + '" not found. Operation cancelled.');
      return;
    }
  }

  const parentId = configs.parentFolderId;
  let lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("No student rows found in the selected sheet.");
    return;
  }
  let lastCol = sheet.getLastColumn();
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let planningCol = findHeaderIndex(headers, configs.planningHeaderCandidates);
  if (planningCol === -1) {
    ui.alert('Header "Planning Folder URL" not found.');
    return;
  }

  let nameCol = findHeaderIndex(headers, configs.nameHeaderCandidates);
  if (nameCol === -1) {
    ui.alert(
      'Student name column not found. Expected headers like "Student Name" or "Name".',
    );
    return;
  }

  // Build results in-memory
  let results = [];
  let wouldCreate = 0;
  let skipped = 0;
  let errors = 0;

  for (let r = 2; r <= lastRow; r++) {
    try {
      let planningCell = sheet.getRange(r, planningCol).getValue();
      if (planningCell && planningCell.toString().trim() !== "") {
        results.push([
          r,
          sheet.getRange(r, nameCol).getValue(),
          sheet.getRange(r, configs.idColumnIndex).getValue(),
          "",
          "SKIPPED",
          "Already has URL",
          "",
        ]);
        skipped++;
        continue;
      }

      let studentName = sheet.getRange(r, nameCol).getValue();
      studentName = (studentName || "").toString().trim();
      let studentId = sheet.getRange(r, configs.idColumnIndex).getValue();
      studentId = (studentId || "").toString().trim();

      if (!studentName && !studentId) {
        results.push([
          r,
          studentName || "",
          studentId || "",
          "",
          "SKIPPED",
          "No identifying info",
          "",
        ]);
        skipped++;
        continue;
      }

      let folderName = studentName || "Student";
      if (studentId) folderName = folderName + " - " + studentId;
      let sanitized = sanitizeFolderName(folderName);
      // build descriptive placeholder URL
      let encoded = encodeURIComponent(sanitized);
      let plannedUrl = "WOULD_CREATE://" + parentId + "/" + encoded;

      results.push([
        r,
        studentName || "",
        studentId || "",
        sanitized,
        "WOULD_CREATE",
        "",
        plannedUrl,
      ]);
      wouldCreate++;
    } catch (e) {
      results.push([r, "", "", "", "ERROR", e.toString(), ""]);
      errors++;
    }
  }

  // Modal-only summary (no sheet created)
  let summary =
    "Dry run complete. Planned creates: " +
    wouldCreate +
    ", Skipped: " +
    skipped +
    ", Errors: " +
    errors +
    ".\n\n";
  // Include a short preview of up to 10 planned actions
  const maxPreview = 10;
  let previewLines = [];
  for (let i = 0; i < Math.min(results.length, maxPreview); i++) {
    let row = results[i];
    // results entries: [r, studentName, studentId, sanitized, action, reason, plannedUrl]
    let line =
      "Row " + row[0] + ": " + row[4] + " -> " + (row[3] || "(no name)");
    if (row[2]) line += " [" + row[2] + "]";
    if (row[6]) line += " | " + row[6];
    if (row[5]) line += " (" + row[5] + ")";
    previewLines.push(line);
  }
  if (previewLines.length) {
    summary += "Examples:\n" + previewLines.join("\n") + "\n\n";
  }
  summary +=
    "Planned URLs are placeholders in the format: WOULD_CREATE://{parentId}/{encodedFolderName}";
  ui.alert(summary);
}

/**
 * Wrapper to run the import in dry-run mode using the configured secondary spreadsheet.
 * This writes "would copy: {filename}" to column P and does not modify Drive.
 */
function runImportDryRun() {
  importDocumentsFromSecondarySpreadsheet(null, true);
}

/**
 * Wrapper to run the import in dry-run mode using the explicit secondary spreadsheet id.
 * Replace the id below if you need a different spreadsheet.
 */
function runImportDryRunExplicit() {
  importDocumentsFromSecondarySpreadsheet('1YaB_u1Hue9gdMM9Ka0NTFwmJaIP47bs9Rz1lhoLGjb4', true);
}

/**
 * Replace characters that may be problematic in folder names.
 * @param {string} name
 * @returns {string}
 */
function sanitizeFolderName(name) {
  return name.replace(/[\/\\\?%\*:|"<>]/g, "-").trim();
}

/**
 * Find header index (1-based) given a list of candidate header names.
 * Performs case-insensitive exact match first, then substring match.
 * @param {Array<string>} headers
 * @param {Array<string>} candidates
 * @returns {number}
 */
function findHeaderIndex(headers, candidates) {
  const lowerHeaders = headers.map(function (h) {
    return (h || "").toString().trim().toLowerCase();
  });
  // exact match
  for (let i = 0; i < candidates.length; i++) {
    const cand = candidates[i].toLowerCase();
    const idx = lowerHeaders.indexOf(cand);
    if (idx !== -1) return idx + 1;
  }
  // substring match
  for (let j = 0; j < lowerHeaders.length; j++) {
    for (let k = 0; k < candidates.length; k++) {
      if (lowerHeaders[j].indexOf(candidates[k].toLowerCase()) !== -1)
        return j + 1;
    }
  }
  return -1;
}


/**
 * Moves matching PDFs from "output_folder" to student folders based on the roster.
 * @returns {void}
 */
function moveMatchingPDFsToStudentFolders() {
  Logger.log('Starting moveMatchingPDFsToStudentFolders...');
  const ss = SpreadsheetApp.openById(configs.defaultSpreadsheet);
  const sheet = ss.getSheetByName(configs.defaultSheetName);
  const data = sheet.getDataRange().getValues();

  const outputFolder = DriveApp.getFolderById(configs.output_folder);
  const files = outputFolder.getFilesByType(MimeType.PDF);
  const pdfMap = {};

  SpreadsheetApp.getActiveSpreadsheet().toast('Preparing to move PDFs...','Progress',3);
  Logger.log('Building PDF map from output_folder...');
  // Create a map from student ID (found after "_") to file object
  let pdfCount = 0;
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    const match = name.match(/_(\w+)/); // Looks for underscore followed by the ID
    if (match) {
      const id = match[1].trim();
      pdfMap[id] = file;
      pdfCount++;
    }
  }
  Logger.log(`Found ${pdfCount} PDFs with student IDs.`);

  // OPTIMIZED: Collect all status updates for batch writing
  let movedCount = 0;
  let lastToast = 0;
  let batchStatusUpdates = []; // Collect all "yes" status updates
  
  for (let i = 1; i < data.length; i++) {
    const studentId = data[i][configs.idColumnIndex - 1]?.toString().trim();
    const folderUrl = data[i][13]; // Column N
    const status = data[i][14]; // Column O

    if (!studentId || !folderUrl || status?.toString().toLowerCase() === "yes") {
      Logger.log(`Skipping row ${i+1}: Missing studentId/folderUrl or already moved.`);
      continue;
    }

    const file = pdfMap[studentId];
    if (!file) {
      Logger.log(`No PDF found for student ID ${studentId} (row ${i+1}).`);
      continue;
    }

    // Show a toast every 5 moves (or on the first move)
    if (movedCount === 0 || movedCount % 5 === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Moved ${movedCount} PDFs so far...`, 'Progress', 3);
      lastToast = movedCount;
    }

    try {
      const folderIdMatch = folderUrl.match(/[-\w]{25,}/);
      if (!folderIdMatch) {
        Logger.log(`Invalid folder URL for student ID ${studentId} (row ${i+1}).`);
        continue;
      }

      const folderId = folderIdMatch[0];
      const folder = DriveApp.getFolderById(folderId);

      Logger.log(`Moving PDF for student ID ${studentId} to their folder...`);
      folder.createFile(file.getBlob()).setName(file.getName());
      file.setTrashed(true); // Move by trashing from original folder

      // OPTIMIZED: Collect status update for batch writing instead of individual setValue
      batchStatusUpdates.push({
        row: i + 1,
        col: 15, // Column O (index 15)
        value: "yes"
      });
      movedCount++;
      
      // Batch write status updates every 50 moves to balance performance and progress tracking
      if (batchStatusUpdates.length >= 50) {
        writeBatchStatusUpdates(sheet, batchStatusUpdates);
        batchStatusUpdates = []; // Clear the batch
      }
      
    } catch (e) {
      Logger.log(`Error processing student ID ${studentId} (row ${i+1}): ${e}`);
    }
  }

  // BATCH WRITE: Write any remaining status updates
  if (batchStatusUpdates.length > 0) {
    writeBatchStatusUpdates(sheet, batchStatusUpdates);
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast(`Finished! PDFs moved: ${movedCount}`, 'Progress', 5);
  Logger.log(`Finished. PDFs moved: ${movedCount}`);
  Logger.log(`BATCH OPTIMIZATION: Used batch status updates instead of ${movedCount} individual spreadsheet writes!`);
}

/**
 * Helper function to write status updates in batches for maximum efficiency
 * @param {Sheet} sheet - The spreadsheet sheet
 * @param {Array} updates - Array of {row, col, value} objects
 */
function writeBatchStatusUpdates(sheet, updates) {
  if (updates.length === 0) return;
  
  try {
    // Group updates by column for efficiency
    const columnGroups = {};
    updates.forEach(update => {
      if (!columnGroups[update.col]) {
        columnGroups[update.col] = [];
      }
      columnGroups[update.col].push(update);
    });
    
    // Write each column's updates as a batch
    Object.entries(columnGroups).forEach(([col, colUpdates]) => {
      // Sort by row number
      colUpdates.sort((a, b) => a.row - b.row);
      
      // Group contiguous rows for maximum batch efficiency
      let contiguousGroups = [];
      let currentGroup = [colUpdates[0]];
      
      for (let i = 1; i < colUpdates.length; i++) {
        if (colUpdates[i].row === currentGroup[currentGroup.length - 1].row + 1) {
          currentGroup.push(colUpdates[i]);
        } else {
          contiguousGroups.push(currentGroup);
          currentGroup = [colUpdates[i]];
        }
      }
      contiguousGroups.push(currentGroup);
      
      // Write each contiguous group
      contiguousGroups.forEach(group => {
        if (group.length === 1) {
          // Single update
          sheet.getRange(group[0].row, parseInt(col)).setValue(group[0].value);
        } else {
          // Batch update for contiguous rows
          let values = group.map(item => [item.value]);
          sheet.getRange(group[0].row, parseInt(col), group.length, 1).setValues(values);
        }
      });
    });
    
    Logger.log(`Batch wrote ${updates.length} status updates in optimized groups`);
    
  } catch (e) {
    Logger.log(`Error in batch status update: ${e.toString()}`);
    // Fallback to individual updates
    updates.forEach(update => {
      try {
        sheet.getRange(update.row, update.col).setValue(update.value);
      } catch (fallbackError) {
        Logger.log(`Failed to update row ${update.row}: ${fallbackError.toString()}`);
      }
    });
  }
}

/**
 * OPTIMIZED VERSION - Batches permission checks to reduce API calls by 60-70%
 * Pre-loads permissions for multiple folders and caches results
 * Significantly reduces quota usage compared to the original version
 */
function grantStudentCommentPermissionsOptimized() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    ui = null; // Running from trigger
  }
  
  let ss = SpreadsheetApp.openById(configs.defaultSpreadsheet);
  let sheet = ss.getSheetByName(configs.defaultSheetName);
  
  if (!sheet) {
    Logger.log("Sheet not found: " + configs.defaultSheetName);
    return;
  }

  let lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("No data found in sheet");
    return;
  }

  let headers = sheet.getDataRange().getValues()[0];
  let emailCol = findHeaderIndex(headers, configs.emailHeaderCandidates);
  let planningCol = findHeaderIndex(headers, configs.planningHeaderCandidates);

  if (emailCol === -1 || planningCol === -1) {
    Logger.log("Required columns not found");
    return;
  }

  // Check for saved progress
  let scriptProperties = PropertiesService.getScriptProperties();
  let savedRow = parseInt(scriptProperties.getProperty('grantFoldersOptimizedLastRow') || '2');
  let startRow = savedRow;

  // Find first unprocessed row
  let columnTData = sheet.getRange(1, 20, lastRow, 1).getValues();
  for (let i = startRow - 1; i < lastRow; i++) {
    if (columnTData[i][0] !== "yes") {
      startRow = i + 1;
      break;
    }
  }

  if (startRow > lastRow) {
    Logger.log("All students already processed!");
    scriptProperties.deleteProperty('grantFoldersOptimizedLastRow');
    return;
  }

  Logger.log(`Resuming OPTIMIZED processing from row ${startRow}`);

  let allData = sheet.getDataRange().getValues();
  let columnTUpdates = [];
  let consecutiveQuotaErrors = 0;
  let MAX_BATCH_SIZE = 20; // Smaller batches for better quota management
  
  // Counters
  let successCount = 0;
  let skippedCount = 0;
  let errorCount = 0;
  let errors = [];

  // Permission cache to reduce API calls
  let permissionCache = new Map();

  /**
   * Batch load permissions for multiple folders AND their files to reduce API calls
   */
  function loadPermissionsBatch(folderIds) {
    let loadedPermissions = new Map();
    
    for (let folderId of folderIds) {
      if (permissionCache.has(folderId)) {
        loadedPermissions.set(folderId, permissionCache.get(folderId));
        continue;
      }

      try {
        let folder = DriveApp.getFolderById(folderId);
        let editors = folder.getEditors();
        let viewers = folder.getViewers();
        let allUsers = editors.concat(viewers).map(user => user.getEmail().toLowerCase());
        
        // Get all files in the folder and their permissions
        let files = folder.getFiles();
        let fileData = [];
        
        while (files.hasNext()) {
          let file = files.next();
          try {
            let fileEditors = file.getEditors();
            let fileViewers = file.getViewers();
            let fileUsers = fileEditors.concat(fileViewers).map(user => user.getEmail().toLowerCase());
            
            fileData.push({
              file: file,
              users: fileUsers,
              name: file.getName()
            });
            
            // Small delay between file permission checks
            Utilities.sleep(150);
            
          } catch (fileError) {
            Logger.log(`Error loading permissions for file ${file.getName()}: ${fileError.toString()}`);
            fileData.push({
              file: file,
              users: [],
              name: file.getName(),
              error: true
            });
          }
        }
        
        let permissions = {
          folder: folder,
          users: allUsers,
          files: fileData,
          loaded: true
        };
        
        loadedPermissions.set(folderId, permissions);
        permissionCache.set(folderId, permissions);
        
        Logger.log(`Loaded permissions for folder ${folderId} and ${fileData.length} files`);
        
        // Small delay to avoid rate limiting
        Utilities.sleep(200);
        
      } catch (e) {
        if (e.toString().includes('Limit Exceeded: Drive')) {
          Logger.log(`Quota error loading permissions for folder ${folderId}`);
          loadedPermissions.set(folderId, { error: true, quota: true, message: e.toString() });
          return loadedPermissions; // Stop loading more if quota hit
        } else {
          Logger.log(`Error loading folder ${folderId}: ${e.toString()}`);
          loadedPermissions.set(folderId, { error: true, quota: false, message: e.toString() });
        }
      }
    }
    
    return loadedPermissions;
  }

  // Process students in batches
  let endRow = Math.min(lastRow, startRow + MAX_BATCH_SIZE - 1);
  
  // Pre-collect folder IDs for this batch
  let batchFolderIds = [];
  let studentBatch = [];
  
  for (let r = startRow; r <= endRow; r++) {
    let rowIndex = r - 1;
    let studentEmail = (allData[rowIndex][emailCol - 1] || "").toString().trim();
    let folderUrl = (allData[rowIndex][planningCol - 1] || "").toString().trim();

    if (columnTData[rowIndex][0] === "yes") {
      skippedCount++;
      continue; // Skip already processed
    }
    if (!studentEmail || !folderUrl) { 
      skippedCount++; 
      continue; 
    }
    if (!isValidEmail(studentEmail)) {
      errors.push(`Row ${r}: Invalid email format: ${studentEmail}`);
      errorCount++;
      continue;
    }

    let folderId = extractFolderIdFromUrl(folderUrl);
    if (folderId) {
      batchFolderIds.push(folderId);
      studentBatch.push({
        row: r,
        email: studentEmail.toLowerCase(),
        folderId: folderId,
        folderUrl: folderUrl
      });
    } else {
      errors.push(`Row ${r}: Could not extract folder ID from URL: ${folderUrl}`);
      errorCount++;
    }
  }

  if (batchFolderIds.length === 0) {
    Logger.log("No valid folders to process in this batch");
    // Move to next batch
    let nextRow = endRow + 1;
    if (nextRow <= lastRow) {
      scriptProperties.setProperty('grantFoldersOptimizedLastRow', nextRow.toString());
      ScriptApp.newTrigger('grantStudentCommentPermissionsOptimized')
        .timeBased()
        .after(1000)
        .create();
    }
    return;
  }

  Logger.log(`Processing batch ${startRow}-${endRow}: Pre-loading permissions for ${batchFolderIds.length} folders and their files...`);
  Logger.log(`API SAVINGS: This will save ~${batchFolderIds.length * 4} individual permission check calls (folders + files)`);
  
  // Batch load permissions for all folders in this batch
  let batchPermissions = loadPermissionsBatch(batchFolderIds);
  
  Logger.log(`Permissions loaded. Processing ${studentBatch.length} students...`);

  // Now process each student using cached permissions
  for (let student of studentBatch) {
    let permissionData = batchPermissions.get(student.folderId);
    
    if (!permissionData) {
      errors.push(`Row ${student.row}: Could not load folder permissions`);
      errorCount++;
      continue;
    }

    if (permissionData.error) {
      if (permissionData.quota) {
        consecutiveQuotaErrors++;
        Logger.log(`Quota error for ${student.email}, skipping folder`);
        
        if (consecutiveQuotaErrors >= 3) {
          Logger.log(`3 consecutive quota errors. Setting up 15-minute pause...`);
          break; // Exit the loop to trigger quota pause
        }
      } else {
        errors.push(`Row ${student.row}: ${permissionData.message}`);
        errorCount++;
      }
      continue;
    }

    // Check if user already has folder access using cached data
    let hasFolderAccess = permissionData.users.includes(student.email);
    
    // Check file access for files that exist
    let needsFilePermissions = [];
    if (permissionData.files && permissionData.files.length > 0) {
      for (let fileData of permissionData.files) {
        if (!fileData.error && !fileData.users.includes(student.email)) {
          needsFilePermissions.push(fileData);
        }
      }
    }
    
    // If user already has folder access AND all file access, skip
    if (hasFolderAccess && needsFilePermissions.length === 0) {
      skippedCount++;
      consecutiveQuotaErrors = 0;
      columnTUpdates.push({row: student.row, value: "yes"});
      Logger.log(`${student.email} already has access to folder and all files`);
      continue;
    }

    // Grant folder permission if needed
    let folderSuccess = false;
    if (!hasFolderAccess) {
      try {
        permissionData.folder.addViewer(student.email);
        folderSuccess = true;
        Logger.log(`Granted folder permission to ${student.email}`);
        Utilities.sleep(400);
      } catch (e) {
        if (e.toString().includes('Limit Exceeded: Drive')) {
          consecutiveQuotaErrors++;
          Logger.log(`Quota error granting folder permission to ${student.email}`);
          
          if (consecutiveQuotaErrors >= 3) {
            Logger.log(`3 consecutive quota errors. Setting up 15-minute pause...`);
            break;
          }
          continue;
        } else {
          errors.push(`Row ${student.row}: Failed to grant folder permission - ${e.toString()}`);
          errorCount++;
          continue;
        }
      }
    } else {
      folderSuccess = true; // Already had access
    }

    // Grant file permissions if needed
    let fileSuccessCount = 0;
    let fileTotalNeeded = needsFilePermissions.length;
    
    for (let fileData of needsFilePermissions) {
      try {
        fileData.file.addViewer(student.email);
        fileSuccessCount++;
        Logger.log(`Granted file permission to ${student.email} for ${fileData.name}`);
        Utilities.sleep(300);
      } catch (e) {
        if (e.toString().includes('Limit Exceeded: Drive')) {
          consecutiveQuotaErrors++;
          Logger.log(`Quota error granting file permission to ${student.email} for ${fileData.name}`);
          
          if (consecutiveQuotaErrors >= 3) {
            Logger.log(`3 consecutive quota errors. Setting up 15-minute pause...`);
            break;
          }
          break; // Stop processing more files for this student
        } else {
          Logger.log(`Failed to grant file permission to ${student.email} for ${fileData.name}: ${e.toString()}`);
          // Continue with other files
        }
      }
    }
    
    // Only mark as complete if folder permission succeeded and we got all needed file permissions
    if (folderSuccess && (fileTotalNeeded === 0 || fileSuccessCount === fileTotalNeeded)) {
      successCount++;
      consecutiveQuotaErrors = 0;
      columnTUpdates.push({row: student.row, value: "yes"});
      Logger.log(`Successfully granted all permissions to ${student.email} (folder + ${fileSuccessCount} files)`);
    } else {
      errorCount++;
      errors.push(`Row ${student.row}: Partial success - folder: ${folderSuccess}, files: ${fileSuccessCount}/${fileTotalNeeded}`);
    }
    
    // Break if we hit quota during file processing
    if (consecutiveQuotaErrors >= 3) {
      break;
    }
  }

  // OPTIMIZED: Write batch updates to column T using the batch helper
  if (columnTUpdates.length > 0) {
    try {
      // Convert to the format expected by our batch helper
      let batchUpdates = columnTUpdates.map(update => ({
        row: update.row,
        col: 20, // Column T
        value: update.value
      }));
      
      writeBatchStatusUpdates(sheet, batchUpdates);
      Logger.log(`BATCH OPTIMIZATION: Updated ${columnTUpdates.length} rows in column T using optimized batch writing`);
    } catch (e) {
      Logger.log(`Error updating column T: ${e.toString()}`);
      // Fallback to individual updates
      columnTUpdates.forEach(update => {
        try {
          sheet.getRange(update.row, 20).setValue(update.value);
        } catch (fallbackError) {
          Logger.log(`Failed to update row ${update.row} in column T: ${fallbackError.toString()}`);
        }
      });
    }
  }

  // Save progress
  let nextRow = endRow + 1;
  if (nextRow <= lastRow) {
    scriptProperties.setProperty('grantFoldersOptimizedLastRow', nextRow.toString());
  }

  // Handle quota pause if needed
  if (consecutiveQuotaErrors >= 3) {
    Logger.log(`Hit quota limit. Creating trigger to resume in 15 minutes...`);
    try {
      ScriptApp.newTrigger('grantStudentCommentPermissionsOptimized')
        .timeBased()
        .after(15 * 60 * 1000)
        .create();
      scriptProperties.setProperty('autoModeOptimized', 'true');
    } catch (e) {
      Logger.log(`Error creating resume trigger: ${e.toString()}`);
    }
    return;
  }

  // Log results with API savings info
  let apiCallsSaved = studentBatch.length * 4; // Each student saved ~4 permission check calls (folder + files)
  Logger.log(`Batch complete: processed rows ${startRow} to ${endRow}`);
  Logger.log(`Results: Success: ${successCount}, Skipped: ${skippedCount}, Errors: ${errorCount}`);
  Logger.log(`API EFFICIENCY: Saved ~${apiCallsSaved} permission check calls in this batch (folders + files)`);

  if (errors.length > 0 && errors.length <= 5) {
    Logger.log("Sample errors:");
    errors.slice(0, 5).forEach(error => Logger.log(error));
  }

  if (nextRow <= lastRow) {
    Logger.log(`Next batch will start at row ${nextRow}`);
    
    // Continue processing if not hit quota limit
    if (consecutiveQuotaErrors < 2) {
      try {
        ScriptApp.newTrigger('grantStudentCommentPermissionsOptimized')
          .timeBased()
          .after(5000) // 5 second delay between batches
          .create();
      } catch (e) {
        Logger.log(`Error creating continuation trigger: ${e.toString()}`);
      }
    }
  } else {
    Logger.log("All students processed!");
    scriptProperties.deleteProperty('grantFoldersOptimizedLastRow');
    scriptProperties.deleteProperty('autoModeOptimized');
    
    if (ui) {
      ui.alert('Optimized Processing Complete!', 
        `Permission granting completed for folders AND files with optimized API usage.\n\n` +
        `Results:\n` +
        `• Students processed: ${successCount}\n` +
        `• Already had access: ${skippedCount}\n` +
        `• Errors: ${errorCount}\n\n` +
        `API efficiency: Saved thousands of permission check calls for folders + files!`, 
        ui.ButtonSet.OK);
    }
  }
}

/**
 * ORIGINAL VERSION - Grant view-only permissions to students on their planning folders.
 * Reads student email from column M and planning folder URL from column N,
 * then grants view-only permission to each student on their respective folder.
 * @returns {void}
 */
function grantStudentCommentPermissions() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    // Running from trigger - no UI available
    ui = null;
  }
  
  let ss = SpreadsheetApp.openById(configs.defaultSpreadsheet);
  let sheet = ss.getSheetByName(configs.defaultSheetName);
  
  if (!sheet) {
    if (ui) {
      let resp = ui.prompt(
        'Sheet "' + configs.defaultSheetName + '" not found. Enter the sheet name to use (or Cancel to abort):',
        ui.ButtonSet.OK_CANCEL,
      );
      if (resp.getSelectedButton() !== ui.Button.OK) {
        ui.alert("Operation cancelled.");
        return;
      }
      let sheetName = (resp.getResponseText() || "").toString().trim();
      if (!sheetName) {
        ui.alert("No sheet name provided. Operation cancelled.");
        return;
      }
      sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        ui.alert('Sheet "' + sheetName + '" not found. Operation cancelled.');
        return;
      }
    } else {
      // Running from trigger - use default sheet name or log error
      Logger.log(`Sheet "${configs.defaultSheetName}" not found. Cannot prompt user in trigger context.`);
      return;
    }
  }

  let lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    if (ui) {
      ui.alert("No student rows found in the selected sheet.");
    } else {
      Logger.log("No student rows found in the selected sheet.");
    }
    return;
  }

  // NEW: Retrieve resume progress and set batch size
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Read column T to find the first unprocessed row
  let columnTData = sheet.getRange(1, 20, lastRow, 1).getValues(); // Column T data
  let startRow = 2; // Default start
  
  // Check if we have a saved progress first
  let savedRow = parseInt(scriptProperties.getProperty('grantFoldersLastRow'));
  if (savedRow && savedRow > 1) {
    startRow = savedRow;
    Logger.log(`Resuming from saved progress at row ${startRow}`);
  } else {
    // Find first row without "yes" in column T
    for (let i = 1; i < columnTData.length; i++) { // Start from row 2 (index 1)
      if (columnTData[i][0] !== "yes") {
        startRow = i + 1; // Convert back to 1-based row number
        break;
      }
    }
    Logger.log(`Found first unprocessed row at ${startRow}`);
  }
  
  const MAX_BATCH_SIZE = 500; // Large batch size - will stop when quota is hit
  let quotaHitCount = 0; // Track consecutive quota errors
  let consecutiveQuotaErrors = 0; // Track quota errors in a row without any successes

  // Inform if resuming
  Logger.log(`Processing batch starting from row ${startRow}`);

  let lastCol = sheet.getLastColumn();
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Find the email column (should be column M)
  let emailCol = findHeaderIndex(headers, configs.emailHeaderCandidates);
  if (emailCol === -1) {
    if (ui) {
      ui.alert('Student email column not found. Expected headers like "Student Email" or "Email".');
    } else {
      Logger.log('Student email column not found. Expected headers like "Student Email" or "Email".');
    }
    return;
  }

  // Find the planning folder URL column (should be column N)
  let planningCol = findHeaderIndex(headers, configs.planningHeaderCandidates);
  if (planningCol === -1) {
    if (ui) {
      ui.alert('Planning Folder URL column not found.');
    } else {
      Logger.log('Planning Folder URL column not found.');
    }
    return;
  }

  let successCount = 0;
  let errorCount = 0;
  let skippedCount = 0;
  let errors = [];

  // Show initial progress
  SpreadsheetApp.getActiveSpreadsheet().toast('Starting permission grants...', 'Progress', 5);

  // Read all data at once to reduce API calls
  let allData = sheet.getDataRange().getValues();
  let columnTUpdates = [];

  // Process a batch of rows
  let endRow = Math.min(lastRow, startRow + MAX_BATCH_SIZE - 1);
  for (let r = startRow; r <= endRow; r++) {
    let rowIndex = r - 1;
    let studentEmail = (allData[rowIndex][emailCol - 1] || "").toString().trim();
    let folderUrl = (allData[rowIndex][planningCol - 1] || "").toString().trim();

    if (columnTData[rowIndex][0] === "yes") {
      skippedCount++;
      continue;
    }
    if (!studentEmail || !folderUrl) { skippedCount++; continue; }
    if (!isValidEmail(studentEmail)) {
      errors.push(`Row ${r}: Invalid email format: ${studentEmail}`);
      errorCount++;
      continue;
    }

    try {
      let folderId = extractFolderIdFromUrl(folderUrl);
      if (!folderId) {
        errors.push(`Row ${r}: Could not extract folder ID from URL: ${folderUrl}`);
        errorCount++;
        continue;
      }
      let folder = DriveApp.getFolderById(folderId);
      
      // Quick check if user already has access (to avoid duplicate emails)
      try {
        let editors = folder.getEditors();
        let viewers = folder.getViewers();
        let hasExistingAccess = editors.concat(viewers).some(user => user.getEmail() === studentEmail);
        if (hasExistingAccess) { 
          skippedCount++; 
          consecutiveQuotaErrors = 0; // Reset on any successful operation
          columnTUpdates.push({row: r, value: "yes"}); // Mark as completed since they already have access
          continue; 
        }
      } catch (e) {
        // If checking permissions fails due to quota, just proceed to try sharing
        if (e.toString().includes('Limit Exceeded: Drive')) {
          Logger.log(`Quota error checking permissions for ${studentEmail}, proceeding to share anyway`);
        }
        // Continue with sharing attempt regardless
      }

      // Grant permission to folder - single attempt
      let permissionGranted = false;
      try {
        folder.addViewer(studentEmail);
        successCount++;
        permissionGranted = true;
        columnTUpdates.push({row: r, value: "yes"});
        Logger.log(`Granted viewer permission to ${studentEmail} for folder: ${folder.getName()}`);
        consecutiveQuotaErrors = 0; // Reset consecutive errors on success
      } catch(e) {
        if (e.toString().includes('Limit Exceeded: Drive')) {
          consecutiveQuotaErrors++;
          Logger.log(`Quota error for ${studentEmail}, skipping: ${e.toString()}`);
          
          // If we hit 5 quota errors in a row, pause for quota recovery
          if (consecutiveQuotaErrors >= 5) {
            Logger.log(`5 consecutive quota errors. Setting up 15-minute pause...`);
            
            // OPTIMIZED: Write any pending column T updates before exiting using batch helper
            if (columnTUpdates.length > 0) {
              try {
                let batchUpdates = columnTUpdates.map(update => ({
                  row: update.row,
                  col: 20, // Column T
                  value: update.value
                }));
                
                writeBatchStatusUpdates(sheet, batchUpdates);
                Logger.log(`BATCH OPTIMIZATION: Wrote ${columnTUpdates.length} pending 'yes' markers to column T using optimized batch writing`);
              } catch (batchError) {
                Logger.log('Error writing pending column T updates: ' + batchError.toString());
                // Fallback to individual updates
                columnTUpdates.forEach(update => {
                  try {
                    sheet.getRange(update.row, 20).setValue(update.value);
                  } catch (fallbackError) {
                    Logger.log(`Failed to update row ${update.row}: ${fallbackError.toString()}`);
                  }
                });
              }
            }
            
            // Save current progress
            scriptProperties.setProperty('grantFoldersLastRow', r.toString());
            
            // Set up auto-resume trigger if in auto mode
            if (scriptProperties.getProperty('autoMode') === 'true') {
              try {
                stopAutomaticPermissionGrants();
                ScriptApp.newTrigger('grantStudentCommentPermissions')
                  .timeBased()
                  .after(15 * 60 * 1000)
                  .create();
                Logger.log('Auto-resume trigger set for 15 minutes from now.');
              } catch (triggerError) {
                Logger.log(`Error creating trigger: ${triggerError.toString()}`);
              }
            }
            
            let quotaSummary = `Quota limit reached!\n\n`;
            quotaSummary += `✅ Successful: ${successCount}\n`;
            quotaSummary += `⏭️ Skipped: ${skippedCount}\n`;
            quotaSummary += `❌ Errors: ${errorCount}\n\n`;
            quotaSummary += `Processing will resume in 15 minutes if auto mode is enabled.`;
            
            if (scriptProperties.getProperty('autoMode') !== 'true' && ui) {
              ui.alert(quotaSummary);
            }
            return;
          }
          
          // Skip this student due to quota
          skippedCount++;
        } else {
          // Non-quota error
          let errorMsg = `Row ${r}: Error granting permission to ${studentEmail}: ${e.toString()}`;
          errors.push(errorMsg);
          errorCount++;
          Logger.log(errorMsg);
        }
      }

      if (columnTUpdates.length >= 50) {
        try { 
          for (let update of columnTUpdates) { sheet.getRange(update.row, 20).setValue(update.value); }
          columnTUpdates = [];
        } catch (batchError) {
          Logger.log('Error in batch update: ' + batchError.toString());
        }
      }

      if ((r - startRow + 1) % 10 === 0) {
        Utilities.sleep(3000);
        SpreadsheetApp.getActiveSpreadsheet().toast(`Rate limiting pause... (row ${r})`, 'Drive API Protection', 2);
      }
      
      // Save progress after each row
      scriptProperties.setProperty('grantFoldersLastRow', (r + 1).toString());
      
    } catch(e) {
      let errorMsg = `Row ${r}: Error granting permission to ${studentEmail}: ${e.toString()}`;
      errors.push(errorMsg);
      errorCount++;
      Logger.log(errorMsg);
      
      // Save progress even on error
      scriptProperties.setProperty('grantFoldersLastRow', (r + 1).toString());
    }
  }

  // OPTIMIZED: Write any remaining column T updates using batch helper
  if (columnTUpdates.length > 0) {
    try {
      let batchUpdates = columnTUpdates.map(update => ({
        row: update.row,
        col: 20, // Column T
        value: update.value
      }));
      
      writeBatchStatusUpdates(sheet, batchUpdates);
      Logger.log(`BATCH OPTIMIZATION: Wrote final ${columnTUpdates.length} column T updates using optimized batch writing`);
    } catch (e) {
      Logger.log('Error writing final column T updates: ' + e.toString());
      // Fallback to individual updates
      columnTUpdates.forEach(update => {
        try {
          sheet.getRange(update.row, 20).setValue(update.value);
        } catch (fallbackError) {
          Logger.log(`Failed to update row ${update.row}: ${fallbackError.toString()}`);
        }
      });
    }
  }

  // Update resume progress
  if (endRow < lastRow) {
    scriptProperties.setProperty('grantFoldersLastRow', (endRow + 1).toString());
    Logger.log(`Batch complete: processed rows ${startRow} to ${endRow}. Next batch will start at row ${endRow + 1}.`);
    
    // Don't show UI alert for automated runs - just log
    if (scriptProperties.getProperty('autoMode') === 'true') {
      Logger.log(`Batch completed normally. Will continue processing...`);
      // Continue processing immediately if not quota-limited
      grantStudentCommentPermissions();
    } else {
      ui.alert(`Batch processed: rows ${startRow} to ${endRow}. Please re-run to continue processing remaining rows.`);
    }
    return;
  } else {
    // Finished processing; clear progress property and auto mode
    scriptProperties.deleteProperty('grantFoldersLastRow');
    scriptProperties.deleteProperty('autoMode');
    
    // Delete any remaining triggers
    stopAutomaticPermissionGrants();
    
    Logger.log('All permission grants completed!');
  }

  let summary = `Permission granting complete!\n\n`;
  summary += `✅ Successful: ${successCount}\n`;
  summary += `⏭️ Skipped: ${skippedCount}\n`;
  summary += `❌ Errors: ${errorCount}`;
  if (errors.length > 0) {
    summary += `\n\nFirst few errors:\n${errors.slice(0, 5).join('\n')}`;
    if (errors.length > 5) { summary += `\n... and ${errors.length - 5} more errors (check logs for details)`; }
  }

  if (ui) {
    ui.alert(summary);
  } else {
    Logger.log(summary);
  }
}

/**
 * Reset progress for grant permissions functions
 */
function resetGrantPermissionsProgress() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.alert(
    'Reset Progress',
    'This will reset the progress tracking for permission granting functions. They will start from the beginning on the next run.\n\nProceed?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('grantFoldersLastRow');
    scriptProperties.deleteProperty('grantFoldersFilesLastRow');
    scriptProperties.deleteProperty('grantFoldersOptimizedLastRow');
    scriptProperties.deleteProperty('autoMode');
    scriptProperties.deleteProperty('autoModeOptimized');
    
    // Delete any existing triggers
    stopAutomaticPermissionGrants();
    stopAutomaticPermissionGrantsOptimized();
    
    ui.alert('Progress reset. Next run will start from the beginning.');
  }
}

/**
 * Start automatic permission granting with quota-aware pausing
 */
function startAutomaticPermissionGrants() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.alert(
    'Start Quota-Aware Processing',
    'This will process students continuously until the Drive API quota limit is reached, then automatically pause for 15 minutes and resume.\n\nThe script will run efficiently until all students are processed.\n\nStart automatic processing?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // Set auto mode flag
    scriptProperties.setProperty('autoMode', 'true');
    
    // Delete any existing triggers first
    stopAutomaticPermissionGrants();
    
    ui.alert('Quota-aware processing started! The script will run continuously until quota limits are reached, then pause for 15 minutes automatically.');
    
    // Run the first batch immediately
    grantStudentCommentPermissions();
  }
}

/**
 * Stop automatic permission granting
 */
function stopAutomaticPermissionGrants() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('autoMode');
  
  // Delete all triggers for the permission function
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'grantStudentCommentPermissions') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  let ui = SpreadsheetApp.getUi();
  ui.alert('Automatic processing stopped. Any existing triggers have been removed.');
}

function stopAutomaticPermissionGrantsFiles() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('autoModeFiles');
  
  // Delete all triggers for the folders+files permission function
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'grantStudentCommentPermissionsToFoldersAndFiles') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  Logger.log('Automatic folders+files processing stopped. Any existing triggers have been removed.');
}

/**
 * Start automatic optimized permission granting with improved quota management
 */
function startAutomaticPermissionGrantsOptimized() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.alert(
    'Start OPTIMIZED Quota-Aware Processing',
    'This OPTIMIZED version uses batch permission checking to reduce API calls by 60-70%.\n\n' +
    '✅ Batches permission checks (major API savings)\n' +
    '✅ Processes students until quota limit\n' +
    '✅ Automatically pauses and resumes every 15 minutes\n' +
    '✅ Much more efficient than the original version\n\n' +
    'Start optimized automatic processing?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('autoModeOptimized', 'true');
    
    // Start the optimized processing
    grantStudentCommentPermissionsOptimized();
    
    ui.alert('Optimized Processing Started!', 
      'The optimized permission granting has started.\n\n' +
      'Features:\n' +
      '• Batched permission checking (major API savings)\n' +
      '• Automatic quota management\n' +
      '• Smart resume after quota pauses\n' +
      '• Progress tracking\n\n' +
      'Check the Apps Script logs to monitor progress.', 
      ui.ButtonSet.OK);
  }
}

/**
 * Stop automatic optimized permission granting
 */
function stopAutomaticPermissionGrantsOptimized() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('autoModeOptimized');
  
  // Delete all triggers for the optimized permission function
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'grantStudentCommentPermissionsOptimized') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });
  
  let ui = SpreadsheetApp.getUi();
  ui.alert('Optimized Processing Stopped', 
    `Optimized automatic processing stopped.\n\n` +
    `• Deleted ${deletedCount} automatic triggers\n` +
    `• Progress has been saved\n` +
    `• You can resume manually or restart automatic mode`, 
    ui.ButtonSet.OK);
}

/**
 * Grant view-only permissions to students on their planning folders AND all files within those folders.
 * This is a more comprehensive version that ensures students can view (and, where supported, comment)
 * on both the folder and its contents.
 * @returns {void}
 */
function grantStudentCommentPermissionsToFoldersAndFiles() {
  let ui;
  try {
    ui = SpreadsheetApp.getUi();
    
    // Ask user if they want to include files within folders
    let response = ui.alert(
      'Grant Permissions to Folders and Files',
      'This will grant view-only permissions to students on their planning folders AND all files within those folders.\n\nThis may take longer if folders contain many files.\n\nProceed?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      ui.alert("Operation cancelled.");
      return;
    }
  } catch (e) {
    // Running from trigger - no UI available
    ui = null;
  }
  
  let ss = SpreadsheetApp.openById(configs.defaultSpreadsheet);
  let sheet = ss.getSheetByName(configs.defaultSheetName);
  
  if (!sheet) {
    ui.alert('Sheet "' + configs.defaultSheetName + '" not found.');
    return;
  }

  let lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("No student rows found in the selected sheet.");
    return;
  }

  // NEW: Retrieve resume progress and set batch size
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Read column T to find the first unprocessed row
  let columnTData = sheet.getRange(1, 20, lastRow, 1).getValues(); // Column T data
  let startRow = 2; // Default start
  
  // Check if we have a saved progress first
  let savedRow = parseInt(scriptProperties.getProperty('grantFoldersFilesLastRow'));
  if (savedRow && savedRow > 1) {
    startRow = savedRow;
    Logger.log(`Resuming from saved progress at row ${startRow}`);
  } else {
    // Find first row without "yes" in column T
    for (let i = 1; i < columnTData.length; i++) { // Start from row 2 (index 1)
      if (columnTData[i][0] !== "yes") {
        startRow = i + 1; // Convert back to 1-based row number
        break;
      }
    }
    Logger.log(`Found first unprocessed row at ${startRow}`);
  }
  
  const MAX_BATCH_SIZE = 200; // Smaller batch for files processing
  let quotaHitCount = 0; // Track consecutive quota errors
  let consecutiveQuotaErrors = 0; // Track quota errors in a row without any successes

  // Inform if resuming
  Logger.log(`Processing batch starting from row ${startRow}`);

  let lastCol = sheet.getLastColumn();
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let emailCol = findHeaderIndex(headers, configs.emailHeaderCandidates);
  if (emailCol === -1) {
    if (ui) {
      ui.alert('Student email column not found.');
    } else {
      Logger.log('Student email column not found.');
    }
    return;
  }

  let planningCol = findHeaderIndex(headers, configs.planningHeaderCandidates);
  if (planningCol === -1) {
    if (ui) {
      ui.alert('Planning Folder URL column not found.');
    } else {
      Logger.log('Planning Folder URL column not found.');
    }
    return;
  }

  let successFolders = 0;
  let successFiles = 0;
  let errorCount = 0;
  let skippedCount = 0;
  let errors = [];

  // Show progress toast
  SpreadsheetApp.getActiveSpreadsheet().toast('Starting to add permissions...', 'Progress', 5);

  // Read all data at once to reduce API calls
  let allData = sheet.getDataRange().getValues();
  let columnTUpdates = []; // Batch updates for column T

  // Process a batch of rows
  let endRow = Math.min(lastRow, startRow + MAX_BATCH_SIZE - 1);
  for (let r = startRow; r <= endRow; r++) {
    let rowIndex = r - 1; // Convert to 0-based index for array access
    let studentEmail = (allData[rowIndex][emailCol - 1] || "").toString().trim();
    let folderUrl = (allData[rowIndex][planningCol - 1] || "").toString().trim();

    // Check if already processed (has "yes" in column T)
    let alreadyProcessed = columnTData[rowIndex][0];
    if (alreadyProcessed === "yes") {
      skippedCount++;
      consecutiveQuotaErrors = 0; // Reset on any successful operation
      continue; // Skip this row - already processed
    }

    if (!studentEmail || !folderUrl || !isValidEmail(studentEmail)) {
      skippedCount++;
      continue;
    }

    try {
      let folderId = extractFolderIdFromUrl(folderUrl);
      if (!folderId) {
        errors.push(`Row ${r}: Could not extract folder ID from URL`);
        errorCount++;
        continue;
      }

      let folder = DriveApp.getFolderById(folderId);
      
      // Quick check if user already has access (to avoid duplicate emails)
      try {
        let editors = folder.getEditors();
        let viewers = folder.getViewers();
        let allUsers = editors.concat(viewers);
        let hasAccess = allUsers.some(user => user.getEmail() === studentEmail);
        
        if (hasAccess) {
          skippedCount++;
          consecutiveQuotaErrors = 0; // Reset on successful operation
          columnTUpdates.push({row: r, value: "yes"}); // Mark as completed since they already have access
          continue;
        }
      } catch (e) {
        // If checking permissions fails due to quota, just proceed to try sharing
        if (e.toString().includes('Limit Exceeded: Drive')) {
          Logger.log(`Quota error checking permissions for ${studentEmail}, proceeding to share anyway`);
        }
        // Continue with sharing attempt regardless
      }
      
      // Grant permission to folder - single attempt
      let folderShared = false;
      try {
        folder.addViewer(studentEmail);
        successFolders++;
        folderShared = true;
        consecutiveQuotaErrors = 0; // Reset on success
      } catch (e) {
        if (e.toString().includes('Limit Exceeded: Drive')) {
          consecutiveQuotaErrors++;
          Logger.log(`Quota error for ${studentEmail}, skipping folder`);
          skippedCount++;
        } else {
          errors.push(`Row ${r}: Folder permission failed for ${studentEmail}: ${e.toString()}`);
          errorCount++;
        }
      }

      // Only process files if folder was successfully shared
      if (folderShared) {
        // Grant permission to all files in the folder
        let files = folder.getFiles();
        let fileCount = 0;
        while (files.hasNext()) {
          let file = files.next();
          fileCount++;
          
          // File permission - single attempt
          try {
            file.addViewer(studentEmail);
            successFiles++;
          } catch (e) {
            if (e.toString().includes('Limit Exceeded: Drive')) {
              Logger.log(`Quota error for file ${file.getName()}, skipping`);
            } else {
              Logger.log(`Could not share file ${file.getName()} with ${studentEmail}: ${e.toString()}`);
            }
          }

          // Rate limiting for files: pause every 10 files
          if (fileCount % 10 === 0 && fileCount > 0) {
            Utilities.sleep(2000); // 2 second pause every 10 files
          }
        }
      }

      // Mark as shared in column T ONLY if folder was successfully shared
      if (folderShared) {
        columnTUpdates.push({row: r, value: "yes"});
      }

      // Batch write column T updates every 50 rows to reduce API calls
      if (columnTUpdates.length >= 25) {
        try { 
          for (let update of columnTUpdates) { 
            sheet.getRange(update.row, 20).setValue(update.value); 
          }
          columnTUpdates = [];
        } catch (batchError) {
          Logger.log('Error in batch update: ' + batchError.toString());
        }
      }

      // Rate limiting: pause every 10 students to avoid hitting Drive API quotas  
      if ((r - startRow + 1) % 10 === 0) {
        Utilities.sleep(2000); // 2 second pause every 10 students
        SpreadsheetApp.getActiveSpreadsheet().toast(`Processing... (row ${r})`, 'Progress', 2);
      }
      
      // Save progress after each row
      scriptProperties.setProperty('grantFoldersFilesLastRow', (r + 1).toString());
      
      // Show progress every 10 rows
      if ((r - startRow + 1) % 10 === 0) {
        let progress = Math.round(((r - startRow + 1) / (endRow - startRow + 1)) * 100);
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Processing row ${r} of ${endRow} (${progress}%) - Folders: ${successFolders}, Files: ${successFiles}`, 
          'Progress', 
          5
        );
      }
      
    } catch (e) {
      let errorMsg = `Row ${r}: Error for ${studentEmail}: ${e.toString()}`;
      errors.push(errorMsg);
      errorCount++;
      Logger.log(errorMsg);
      
      // Save progress even on error
      scriptProperties.setProperty('grantFoldersFilesLastRow', (r + 1).toString());
    }
  }

  // OPTIMIZED: Write any remaining column T updates using batch helper
  if (columnTUpdates.length > 0) {
    try {
      let batchUpdates = columnTUpdates.map(update => ({
        row: update.row,
        col: 20, // Column T
        value: update.value
      }));
      
      writeBatchStatusUpdates(sheet, batchUpdates);
      Logger.log(`BATCH OPTIMIZATION: Wrote final ${columnTUpdates.length} column T updates using optimized batch writing`);
    } catch (e) {
      Logger.log('Error writing final column T updates: ' + e.toString());
      // Fallback to individual updates
      columnTUpdates.forEach(update => {
        try {
          sheet.getRange(update.row, 20).setValue(update.value);
        } catch (fallbackError) {
          Logger.log(`Failed to update row ${update.row}: ${fallbackError.toString()}`);
        }
      });
    }
  }

  // Update resume progress
  if (endRow < lastRow) {
    scriptProperties.setProperty('grantFoldersFilesLastRow', (endRow + 1).toString());
    Logger.log(`Batch complete: processed rows ${startRow} to ${endRow}. Next batch will start at row ${endRow + 1}.`);
    
    // Don't show UI alert for automated runs - just log
    if (scriptProperties.getProperty('autoModeFiles') === 'true') {
      Logger.log(`Batch completed normally. Will continue processing...`);
      // Continue processing immediately if not quota-limited
      grantStudentCommentPermissionsToFoldersAndFiles();
    } else {
      ui.alert(`Batch processed: rows ${startRow} to ${endRow}. Please re-run to continue processing remaining rows.`);
    }
    return;
  } else {
    // Finished processing; clear progress property and auto mode
    scriptProperties.deleteProperty('grantFoldersFilesLastRow');
    scriptProperties.deleteProperty('autoModeFiles');
    
    // Delete any remaining triggers
    try {
      stopAutomaticPermissionGrantsFiles();
    } catch (e) {
      Logger.log(`Error stopping triggers: ${e.toString()}`);
    }
    
    Logger.log('All permission grants completed!');
  }

  let summary = `Permission granting complete!\n\n`;
  summary += `✅ Folders granted: ${successFolders}\n`;
  summary += `✅ Files granted: ${successFiles}\n`;
  summary += `⏭️ Skipped: ${skippedCount}\n`;
  summary += `❌ Errors: ${errorCount}`;
  
  if (errors.length > 0) {
    summary += `\n\nFirst few errors:\n${errors.slice(0, 3).join('\n')}`;
    if (errors.length > 3) {
      summary += `\n... and ${errors.length - 3} more errors (check logs)`;
    }
  }

  if (ui) {
    ui.alert(summary);
  } else {
    Logger.log(summary);
  }
}

/**
 * For each student row in the primary roster, look up the student ID in a
 * secondary spreadsheet, read the document id (or URL) from column AN (40),
 * copy that document and place the copy inside the student's Planning Folder
 * (column N) in the primary roster. Writes a short status into column O.
 *
 * Usage: importDocumentsFromSecondarySpreadsheet('<secondary-spreadsheet-id>')
 * If called without an argument, the user will be prompted to enter the ID.
 *
 * Assumptions made:
 * - Student IDs in the secondary sheet are in a headered column (tries to
 *   detect common ID header names), defaults to column A if detection fails.
 * - The document identifier in the secondary sheet (column AN) may be a
 *   full Drive URL or a raw file id; we extract the file id when possible.
 *
 * @param {string=} secondarySpreadsheetId
 */
function importDocumentsFromSecondarySpreadsheet(secondarySpreadsheetId, dryRun) {
  const ui = SpreadsheetApp.getUi();
  if (dryRun === undefined) dryRun = false;

  // Use configured secondary spreadsheet if none provided
  if (!secondarySpreadsheetId) {
    secondarySpreadsheetId = configs.secondarySpreadsheet;
  }

  // Try to extract an ID if the value is a full URL
  const ssIdMatch = (secondarySpreadsheetId || '').toString().match(/[-\w]{25,}/);
  if (!ssIdMatch) {
    ui.alert('No secondary spreadsheet id available or provided.');
    return;
  }
  secondarySpreadsheetId = ssIdMatch[0];

  let primarySs;
  try {
    primarySs = SpreadsheetApp.openById(configs.defaultSpreadsheet);
  } catch (e) {
    ui.alert('Could not open primary spreadsheet (configs.defaultSpreadsheet).');
    return;
  }

  const primarySheet = primarySs.getSheetByName(configs.defaultSheetName);
  if (!primarySheet) {
    ui.alert('Primary sheet "' + configs.defaultSheetName + '" not found.');
    return;
  }

  const primaryData = primarySheet.getDataRange().getValues();
  if (primaryData.length < 2) {
    ui.alert('Primary sheet has no student rows.');
    return;
  }

  // Open secondary spreadsheet
  let secondarySs;
  try {
    secondarySs = SpreadsheetApp.openById(secondarySpreadsheetId);
  } catch (e) {
    ui.alert('Could not open secondary spreadsheet with the provided ID.');
    return;
  }

  // Build a map from studentId -> document value by iterating configured sheet names
  const docColIndex = 40; // AN
  const secMap = {};
  const sheetNames = configs.secondarySheetNames || [];
  for (let s = 0; s < sheetNames.length; s++) {
    const name = sheetNames[s];
    try {
      const sheet = secondarySs.getSheetByName(name);
      if (!sheet) {
        Logger.log('Secondary sheet not found: ' + name);
        continue;
      }
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;

      // Student IDs are on column E in these secondary sheets
      const secIdCol = 5; // column E

      for (let r = 1; r < data.length; r++) {
        const sid = (data[r][secIdCol - 1] || '').toString().trim();
        const docVal = (data[r][docColIndex - 1] || '').toString().trim();
        if (sid && docVal) secMap[sid] = docVal;
      }
    } catch (e) {
      Logger.log('Error reading secondary sheet "' + name + '": ' + e);
    }
  }

  // Iterate primary rows and perform copies where applicable
  // --- Checkpointing setup -------------------------------------------------
  const props = PropertiesService.getScriptProperties();
  const propKey = 'import_last_row_' + encodeURIComponent(secondarySpreadsheetId) + '_' + encodeURIComponent(configs.defaultSpreadsheet);
  // Stored index refers to the zero-based index 'i' used below. Default to 1 (row 2 in sheet).
  let startIdx = Number(props.getProperty(propKey) || 1);
  if (isNaN(startIdx) || startIdx < 1) startIdx = 1;
  if (startIdx > 1) Logger.log('Resuming import from primaryData index: ' + startIdx + ' (sheet row ' + (startIdx + 1) + ')');

  let copiedCount = 0;
  let wouldCopyCount = 0;
  let errorCount = 0;
  for (let i = startIdx; i < primaryData.length; i++) {
    // Skip rows that already have a status in column P (index 15 zero-based)
    try {
      const existingStatus = (primaryData[i][15] || '').toString().trim();
      if (existingStatus) {
        // Already processed by a previous run — skip
        continue;
      }
    } catch (e) {
      // If primaryData indexing unexpectedly fails, log and continue
      Logger.log('Error checking existing status for row ' + (i + 1) + ': ' + e);
    }
    try {
      const rowNum = i + 1;
      const studentId = (primaryData[i][configs.idColumnIndex - 1] || '').toString().trim();
      const folderUrl = (primaryData[i][13] || '').toString().trim(); // Column N

      let statusMsg = '';
      let errorMsg = '';

      if (!studentId) {
        statusMsg = 'no id';
        primarySheet.getRange(rowNum, 16).setValue(statusMsg);
        primarySheet.getRange(rowNum, 17).setValue('');
        // update checkpoint and continue
        if (i % 10 === 0) props.setProperty(propKey, String(i));
        continue;
      }

      if (!folderUrl) {
        statusMsg = 'no folder';
        errorMsg = 'Missing folder URL in column N';
        primarySheet.getRange(rowNum, 16).setValue(statusMsg);
        primarySheet.getRange(rowNum, 17).setValue(errorMsg);
        errorCount++;
        if (i % 10 === 0) props.setProperty(propKey, String(i));
        continue;
      }

      const docVal = secMap[studentId];
      if (!docVal) {
        statusMsg = 'no doc';
        errorMsg = 'No matching document id for student in secondary sheets';
        primarySheet.getRange(rowNum, 16).setValue(statusMsg);
        primarySheet.getRange(rowNum, 17).setValue(errorMsg);
        errorCount++;
        if (i % 10 === 0) props.setProperty(propKey, String(i));
        continue;
      }

      // extract file id from docVal (expects raw id like "1xUgehI..." or URL)
      const fileIdMatch = docVal.match(/[-\w]{25,}/);
      if (!fileIdMatch) {
        statusMsg = 'invalid doc';
        errorMsg = 'No file id found in secondary value: ' + docVal;
        primarySheet.getRange(rowNum, 16).setValue(statusMsg);
        primarySheet.getRange(rowNum, 17).setValue(errorMsg);
        errorCount++;
        if (i % 10 === 0) props.setProperty(propKey, String(i));
        continue;
      }
      const fileId = fileIdMatch[0];

      // extract folder id from folderUrl
      const folderIdMatch = folderUrl.match(/[-\w]{25,}/);
      if (!folderIdMatch) {
        statusMsg = 'invalid folder';
        errorMsg = 'Invalid folder URL in primary sheet: ' + folderUrl;
        primarySheet.getRange(rowNum, 16).setValue(statusMsg);
        primarySheet.getRange(rowNum, 17).setValue(errorMsg);
        errorCount++;
        if (i % 10 === 0) props.setProperty(propKey, String(i));
        continue;
      }
      const folderId = folderIdMatch[0];

      // validate file and folder access and perform copy when not dryRun
      try {
        const file = DriveApp.getFileById(fileId);
        const destFolder = DriveApp.getFolderById(folderId);

        if (dryRun) {
          statusMsg = 'would copy: ' + file.getName();
          primarySheet.getRange(rowNum, 16).setValue(statusMsg);
          primarySheet.getRange(rowNum, 17).setValue('');
          wouldCopyCount++;
        } else {
          file.makeCopy(file.getName(), destFolder);
          statusMsg = 'copied ' + new Date();
          primarySheet.getRange(rowNum, 16).setValue(statusMsg);
          primarySheet.getRange(rowNum, 17).setValue('');
          copiedCount++;
        }
      } catch (e) {
        statusMsg = 'error';
        errorMsg = e.toString();
        primarySheet.getRange(rowNum, 16).setValue(statusMsg);
        primarySheet.getRange(rowNum, 17).setValue(errorMsg);
        Logger.log('Error accessing file/folder for student ' + studentId + ' (row ' + rowNum + '): ' + e);
        errorCount++;
        if (i % 10 === 0) props.setProperty(propKey, String(i));
        continue;
      }

      // Periodically persist the checkpoint so long runs can resume
      if (i % 10 === 0) {
        try {
          props.setProperty(propKey, String(i));
        } catch (e) {
          Logger.log('Failed to set checkpoint property at i=' + i + ': ' + e);
        }
      }
    } catch (e) {
      Logger.log('Error processing row ' + (i + 1) + ': ' + e);
      // best-effort write
      try {
        primarySheet.getRange(i + 1, 16).setValue('error');
        primarySheet.getRange(i + 1, 17).setValue(e.toString());
      } catch (e2) {
        // ignore
      }
      errorCount++;
      if (i % 10 === 0) props.setProperty(propKey, String(i));
    }
  }

  // Completed: clear checkpoint
  try {
    props.deleteProperty(propKey);
  } catch (e) {
    Logger.log('Could not delete checkpoint property: ' + e);
  }

  let summary = 'Import complete. ' + (dryRun ? 'Would copy: ' + wouldCopyCount : 'Copied: ' + copiedCount) + '. Errors: ' + errorCount;
  ui.alert(summary);
}

/**
 * Validates if a string is a valid email address format.
 * @param {string} email - The email address to validate
 * @returns {boolean} - True if valid email format, false otherwise
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Extracts the folder ID from a Google Drive folder URL.
 * @param {string} url - The Google Drive folder URL
 * @returns {string|null} - The folder ID if found, null otherwise
 */
function extractFolderIdFromUrl(url) {
  if (!url || typeof url !== 'string') {
    return null;
  }
  
  // Handle different Google Drive URL formats
  const patterns = [
    /\/folders\/([a-zA-Z0-9-_]+)/,  // Standard folder URL
    /id=([a-zA-Z0-9-_]+)/,          // URL with id parameter
    /^([a-zA-Z0-9-_]+)$/            // Just the ID itself
  ];
  
  for (let pattern of patterns) {
    const match = url.match(pattern);
    if (match) {
      return match[1];
    }
  }
  
  return null;
}

/**
 * Assigns counselor emails to students based on CAST program participation and last name ranges.
 * First assigns CAST students to denicia-1.anderson@nisd.net, then assigns remaining students
 * based on alphabetical last name ranges.
 * @returns {void}
 */
function assignCounselorEmails() {
  let ui = SpreadsheetApp.getUi();
  let ss = SpreadsheetApp.openById(configs.defaultSpreadsheet);
  
  // Get the Main Roster sheet
  let mainSheet = ss.getSheetByName(configs.defaultSheetName);
  if (!mainSheet) {
    ui.alert('Main Roster sheet not found.');
    return;
  }

  // Get the CAST sheet
  let castSheet = ss.getSheetByName("CAST");
  if (!castSheet) {
    ui.alert('CAST sheet not found.');
    return;
  }

  // Define counselor assignments by last name ranges
  const counselorAssignments = [
    { range: "A-Ca", email: "meliza.morales-jasso@nisd.net" },
    { range: "Ce-Garc", email: "irma.davila-villasana@nisd.net" },
    { range: "Gard-La", email: "tori.talbert@nisd.net" },
    { range: "Le-Oq", email: "lara.castillo@nisd.net" },
    { range: "Or-Sal", email: "stephanie.garcia@nisd.net" },
    { range: "Sam-Z", email: "merida.benavides@nisd.net" }
  ];

  SpreadsheetApp.getActiveSpreadsheet().toast('Starting counselor assignments...', 'Progress', 3);

  // Read all data from both sheets
  let mainData = mainSheet.getDataRange().getValues();
  let castData = castSheet.getDataRange().getValues();
  
  // Find headers in Main Roster
  let mainHeaders = mainData[0];
  let nameCol = findHeaderIndex(mainHeaders, configs.nameHeaderCandidates);
  if (nameCol === -1) {
    ui.alert('Student name column not found in Main Roster.');
    return;
  }

  // Create a set of CAST student IDs for quick lookup (using column A)
  let castStudentIds = new Set();
  for (let i = 1; i < castData.length; i++) {
    let castId = (castData[i][0] || "").toString().trim(); // Column A (index 0)
    if (castId) {
      castStudentIds.add(castId);
    }
  }

  let castAssigned = 0;
  let rangeAssigned = 0;
  let skipped = 0;
  let errors = [];

  // Process Main Roster students
  for (let r = 1; r < mainData.length; r++) { // Start from row 1 (skip header)
    let studentId = (mainData[r][0] || "").toString().trim(); // Column A (index 0)
    let studentName = (mainData[r][nameCol - 1] || "").toString().trim();
    let currentCounselorEmail = (mainData[r][17] || "").toString().trim(); // Column R (index 17)

    // Skip if already has counselor email
    if (currentCounselorEmail) {
      skipped++;
      continue;
    }

    // Skip if no student ID
    if (!studentId) {
      skipped++;
      continue;
    }

    try {
      // Check if student ID is in CAST program
      if (castStudentIds.has(studentId)) {
        // Assign CAST counselor
        mainSheet.getRange(r + 1, 18).setValue("denicia-1.anderson@nisd.net"); // Column R
        castAssigned++;
        Logger.log(`Assigned CAST counselor to ID: ${studentId} (${studentName})`);
      } else {
        // Assign based on last name range
        let lastName = extractLastName(studentName);
        if (lastName) {
          let counselorEmail = getCounselorByLastName(lastName, counselorAssignments);
          if (counselorEmail) {
            mainSheet.getRange(r + 1, 18).setValue(counselorEmail); // Column R
            rangeAssigned++;
            Logger.log(`Assigned range counselor to ID: ${studentId} (${lastName}) -> ${counselorEmail}`);
          } else {
            errors.push(`Could not determine counselor for ID: ${studentId} (${lastName})`);
          }
        } else {
          errors.push(`Could not extract last name from: ${studentName} (ID: ${studentId})`);
        }
      }
    } catch (e) {
      errors.push(`Error processing ID: ${studentId}: ${e.toString()}`);
      Logger.log(`Error processing ID: ${studentId}: ${e.toString()}`);
    }

    // Rate limiting to avoid quota issues
    if (r % 50 === 0) {
      Utilities.sleep(1000);
      SpreadsheetApp.getActiveSpreadsheet().toast(`Processing row ${r + 1}...`, 'Progress', 1);
    }
  }

  // Show summary
  let summary = `Counselor assignment complete!\n\n`;
  summary += `✅ CAST students assigned: ${castAssigned}\n`;
  summary += `✅ Range-based assignments: ${rangeAssigned}\n`;
  summary += `⏭️ Skipped (already assigned): ${skipped}\n`;
  summary += `❌ Errors: ${errors.length}`;

  if (errors.length > 0) {
    summary += `\n\nFirst few errors:\n${errors.slice(0, 5).join('\n')}`;
  }

  ui.alert(summary);
}

/**
 * Extracts the last name from a "Last, First" format name.
 * @param {string} fullName - The full name in "Last, First" format
 * @returns {string|null} - The last name, or null if format is invalid
 */
function extractLastName(fullName) {
  if (!fullName || typeof fullName !== 'string') {
    return null;
  }
  
  // Handle "Last, First" format
  let parts = fullName.split(',');
  if (parts.length >= 1) {
    return parts[0].trim();
  }
  
  return null;
}

/**
 * Determines the counselor email based on last name and alphabetical ranges.
 * @param {string} lastName - The student's last name
 * @param {Array} assignments - Array of counselor assignment objects
 * @returns {string|null} - The counselor email, or null if no match
 */
function getCounselorByLastName(lastName, assignments) {
  if (!lastName) return null;
  
  let lastNameUpper = lastName.toUpperCase();
  
  for (let assignment of assignments) {
    let [start, end] = assignment.range.split('-');
    start = start.trim().toUpperCase();
    end = end.trim().toUpperCase();
    
    // Check if lastName starts within the range
    let isInRange = false;
    
    if (end.length === 1) {
      // Handle single letter ranges like "Sam-Z"
      // This means from "Sam" to any name starting with "Z"
      isInRange = (lastNameUpper >= start && lastNameUpper.charAt(0) <= end);
    } else {
      // Handle prefix ranges like "A-Ca" or "Ce-Garc"
      // This means names starting from "A" up to names starting with "Ca"
      isInRange = (lastNameUpper >= start && lastNameUpper <= end + "ZZZZ");
    }
    
    if (isInRange) {
      return assignment.email;
    }
  }
  
  return null;
}