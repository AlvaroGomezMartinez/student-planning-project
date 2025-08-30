/**
 * Triggered when the spreadsheet is opened. Adds a custom menu for student tools.
 * @returns {void}
 */
function onOpen() {
  // Adds a custom menu so users can run the folder-creation from the UI
  SpreadsheetApp.getUi()
    .createMenu("Student Tools")
    .addItem("Instructions", "showInstructions")
    .addItem("Create Student Folders", "createStudentFolders")
    .addItem("Move PDFs in output_folder to Student Folders", "moveMatchingPDFsToStudentFolders")
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

  let updatedCount = 0;

  // Loop rows 2..lastRow and create folders where Planning Folder URL is empty
  for (let r = 2; r <= lastRow; r++) {
    let planningCell = sheet.getRange(r, planningCol).getValue();
    if (planningCell && planningCell.toString().trim() !== "") {
      continue; // already has a URL
    }

    let studentName = sheet.getRange(r, nameCol).getValue();
    studentName = (studentName || "").toString().trim();
    let studentId = sheet.getRange(r, configs.idColumnIndex).getValue(); // column A
    studentId = (studentId || "").toString().trim();

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

    // Write URL back into Planning Folder URL column
    sheet.getRange(r, planningCol).setValue(url);
    updatedCount++;
  }

  ui.alert(
    "Created " + updatedCount + " folders and updated Planning Folder URLs.",
  );
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

  let movedCount = 0;
  let lastToast = 0;
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

      sheet.getRange(i + 1, 15).setValue("yes"); // Column O (index 15)
      movedCount++;
    } catch (e) {
      Logger.log(`Error processing student ID ${studentId} (row ${i+1}): ${e}`);
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(`Finished! PDFs moved: ${movedCount}`,'Progress',5);
  Logger.log(`Finished. PDFs moved: ${movedCount}`);
}

// Todo: Add a function that shares the student folders with each student and other people who need access perhaps their counselor.

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