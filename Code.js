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
    .addItem("Create Student Folders (Dry Run)", "createStudentFoldersDryRun")
    .addItem("Move PDFs in output_folder to Student Folders", "moveMatchingPDFsToStudentFolders")
  // .addItem("Import Planning PDF", "importPlanningPdf")
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
  planningHeaderCandidates: [ // Used to help find the correct columns
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