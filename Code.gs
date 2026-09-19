/**
 * ==============================================================================
 * CloudOS Core Engine - Google Apps Script
 * ==============================================================================
 * Acts as the centralized storage, authentication, and Google Drive manager.
 * No user-side OAuth keys required.
 * ==============================================================================
 */

const DB_SCHEMA = {
  Users: ['UserID', 'Username', 'Email', 'PasswordHash', 'DriveFolderId', 'CreatedAt', 'LastLogin'],
  Contacts: ['ID', 'Username', 'Name', 'Phone', 'Email', 'Notes', 'UpdatedAt'],
  Reminders: ['ID', 'Username', 'Title', 'DateTime', 'Status', 'UpdatedAt'],
  Calendar: ['ID', 'Username', 'Title', 'Date', 'Time', 'Description', 'UpdatedAt'],
  Tasks: ['ID', 'Username', 'Task', 'Completed', 'DueDate', 'UpdatedAt'],
  Bookmarks: ['ID', 'Username', 'Title', 'URL', 'Icon', 'UpdatedAt'],
  Apps: ['AppID', 'Username', 'Name', 'Icon', 'Type', 'CodeOrUrl', 'CreatedAt']
};

/**
 * Ensures all database sheets exist, sets headers, styling, and frozen rows.
 */
function initDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(DB_SCHEMA).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    const headers = DB_SCHEMA[sheetName];
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(String);
    if (existing.length === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      styleHeader(sheet, headers.length);
    } else {
      const missing = headers.filter(h => !existing.includes(h));
      if (missing.length > 0) {
        sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
        styleHeader(sheet, existing.length + missing.length);
      }
    }
  });
  return { status: "success", message: "Database schema verified and active." };
}

function styleHeader(sheet, colCount) {
  const header = sheet.getRange(1, 1, 1, colCount);
  header.setBackground("#1a73e8"); // Google Blue
  header.setFontColor("#ffffff");
  header.setFontWeight("bold");
  header.setFontFamily("Roboto");
  sheet.setFrozenRows(1);
}

function doGet(e) {
  initDatabase();
  return jsonResponse({
    status: "online",
    name: "CloudOS Unified Backend",
    version: "2.5",
    time: new Date().toISOString()
  });
}

function doPost(e) {
  initDatabase();
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ status: "error", message: "Malformed JSON payload: " + err.message });
  }

  const action = payload.action;
  try {
    switch (action) {
      case "register": return registerUser(payload);
      case "login": return loginUser(payload);
      case "forgot_password": return forgotPassword(payload);
      case "update_password": return updatePassword(payload);
      case "sync_sheet": return syncSheet(payload);
      case "get_sheet": return getSheet(payload);
      case "drive_upload": return uploadToDrive(payload);
      case "drive_list": return listDriveFiles(payload);
      case "drive_delete": return deleteDriveFile(payload);
      case "save_app": return saveInstalledApp(payload);
      case "get_apps": return getInstalledApps(payload);
      case "delete_app": return deleteInstalledApp(payload);
      default: return jsonResponse({ status: "error", message: "Invalid action specified: " + action });
    }
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * Register User & initialize dedicated Google Drive folder under CloudOS/<Username>
 */
function registerUser(data) {
  const { username, email, password } = data;
  if (!username || !email || !password) {
    return jsonResponse({ status: "error", message: "All registration fields are required." });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().toLowerCase() === username.toLowerCase()) {
      return jsonResponse({ status: "error", message: "Username already taken." });
    }
    if (rows[i][2].toString().toLowerCase() === email.toLowerCase()) {
      return jsonResponse({ status: "error", message: "Email is already registered." });
    }
  }

  // Create isolated user folder inside Google Drive
  let rootFolder;
  const rootIterator = DriveApp.getFoldersByName("CloudOS");
  if (rootIterator.hasNext()) {
    rootFolder = rootIterator.next();
  } else {
    rootFolder = DriveApp.createFolder("CloudOS");
  }

  let userFolder;
  const userIterator = rootFolder.getFoldersByName(username);
  if (userIterator.hasNext()) {
    userFolder = userIterator.next();
  } else {
    userFolder = rootFolder.createFolder(username);
    // Create subfolders
    userFolder.createFolder("Documents");
    userFolder.createFolder("Scans");
    userFolder.createFolder("Code");
    userFolder.createFolder("Media");
  }

  const userId = "USR_" + Utilities.getUuid().slice(0, 8);
  const now = new Date().toISOString();

  sheet.appendRow([userId, username, email, password, userFolder.getId(), now, now]);

  return jsonResponse({
    status: "success",
    message: "Registration successful!",
    user: {
      userId: userId,
      username: username,
      email: email,
      driveFolderId: userFolder.getId()
    }
  });
}

/**
 * Authenticate User
 */
function loginUser(data) {
  const { username, password } = data;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[1].toString().toLowerCase() === username.toLowerCase()) {
      if (row[3].toString() === password) {
        sheet.getRange(i + 1, 7).setValue(new Date().toISOString()); // Update LastLogin
        return jsonResponse({
          status: "success",
          message: "Login successful!",
          user: {
            userId: row[0],
            username: row[1],
            email: row[2],
            driveFolderId: row[4]
          }
        });
      } else {
        return jsonResponse({ status: "error", message: "Incorrect password." });
      }
    }
  }
  return jsonResponse({ status: "error", message: "User does not exist." });
}

/**
 * Forgot Password: Dispatches email directly to user's registered inbox
 */
function forgotPassword(data) {
  const { email } = data;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[2].toString().toLowerCase() === email.toLowerCase()) {
      const username = row[1];
      const password = row[3];
      const subject = "CloudOS Account Password Recovery";
      const body = "Hello " + username + ",\n\nYour CloudOS password is: " + password + 
                   "\n\nYou can sign in and change your password in the Control Panel.\n\n— CloudOS System";
      MailApp.sendEmail(email, subject, body);
      return jsonResponse({ status: "success", message: "Your password has been emailed to you." });
    }
  }
  return jsonResponse({ status: "error", message: "Email address not found." });
}

/**
 * Update user password
 */
function updatePassword(data) {
  const { username, oldPassword, newPassword } = data;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().toLowerCase() === username.toLowerCase()) {
      if (rows[i][3].toString() === oldPassword) {
        sheet.getRange(i + 1, 4).setValue(newPassword);
        return jsonResponse({ status: "success", message: "Password updated successfully." });
      } else {
        return jsonResponse({ status: "error", message: "Old password does not match." });
      }
    }
  }
  return jsonResponse({ status: "error", message: "User not found." });
}

/**
 * Upload base64 or text file directly to user's Google Drive folder
 */
function uploadToDrive(data) {
  const { username, folderId, fileName, mimeType, base64Content } = data;
  let targetFolder;

  try {
    targetFolder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  } catch (e) {
    targetFolder = DriveApp.getRootFolder();
  }

  const decoded = Utilities.base64Decode(base64Content);
  const blob = Utilities.newBlob(decoded, mimeType || 'application/octet-stream', fileName);
  const file = targetFolder.createFile(blob);

  return jsonResponse({
    status: "success",
    message: "File stored in Google Drive.",
    file: {
      id: file.getId(),
      name: file.getName(),
      size: file.getSize(),
      url: file.getUrl(),
      mimeType: file.getMimeType()
    }
  });
}

/**
 * List files inside the user's Google Drive folder
 */
function listDriveFiles(data) {
  const { folderId } = data;
  let targetFolder;
  try {
    targetFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    return jsonResponse({ status: "error", message: "Folder not found: " + e.message });
  }

  const files = [];
  const fileIterator = targetFolder.getFiles();
  while (fileIterator.hasNext()) {
    const f = fileIterator.next();
    files.push({
      id: f.getId(),
      name: f.getName(),
      size: (f.getSize() / 1024).toFixed(1) + " KB",
      mimeType: f.getMimeType(),
      url: f.getUrl(),
      lastUpdated: f.getLastUpdated().toISOString()
    });
  }

  const subfolders = [];
  const folderIterator = targetFolder.getFolders();
  while (folderIterator.hasNext()) {
    const sf = folderIterator.next();
    subfolders.push({
      id: sf.getId(),
      name: sf.getName()
    });
  }

  return jsonResponse({
    status: "success",
    files: files,
    folders: subfolders
  });
}

/**
 * Delete a file in Google Drive
 */
function deleteDriveFile(data) {
  const { fileId } = data;
  try {
    const file = DriveApp.getFileById(fileId);
    file.setTrashed(true);
    return jsonResponse({ status: "success", message: "File moved to trash." });
  } catch (e) {
    return jsonResponse({ status: "error", message: e.message });
  }
}

/**
 * Sync Google Sheet tab data (Contacts, Reminders, Calendar, Tasks, Bookmarks)
 */
function syncSheet(data) {
  const { username, sheetName, items } = data;
  if (!DB_SCHEMA[sheetName]) {
    return jsonResponse({ status: "error", message: "Invalid sheet target: " + sheetName });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];

  // Purge existing records for user
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1].toString().toLowerCase() === username.toLowerCase()) {
      sheet.deleteRow(i + 1);
    }
  }

  // Insert updated batch
  if (Array.isArray(items) && items.length > 0) {
    const rowsToAdd = items.map(item => {
      return headers.map(h => {
        if (h === 'Username') return username;
        if (h === 'UpdatedAt') return new Date().toISOString();
        return item[h] !== undefined ? item[h] : (item[h.toLowerCase()] || "");
      });
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, headers.length).setValues(rowsToAdd);
  }

  return jsonResponse({ status: "success", message: sheetName + " updated." });
}

/**
 * Retrieve user's records from Sheet
 */
function getSheet(data) {
  const { username, sheetName } = data;
  if (!DB_SCHEMA[sheetName]) {
    return jsonResponse({ status: "error", message: "Invalid sheet target." });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const rows = sheet.getDataRange().getValues();

  if (rows.length <= 1) return jsonResponse({ status: "success", data: [] });

  const headers = rows[0];
  const userRows = [];

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().toLowerCase() === username.toLowerCase()) {
      const obj = {};
      headers.forEach((h, idx) => { obj[h] = rows[i][idx]; });
      userRows.push(obj);
    }
  }

  return jsonResponse({ status: "success", data: userRows });
}

/**
 * Save custom installed application (code or URL)
 */
function saveInstalledApp(data) {
  const { username, name, icon, type, codeOrUrl } = data;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Apps");
  const appId = "APP_" + Utilities.getUuid().slice(0, 8);
  const now = new Date().toISOString();

  sheet.appendRow([appId, username, name, icon || "box", type, codeOrUrl, now]);
  return jsonResponse({ status: "success", message: "App installed successfully.", appId: appId });
}

/**
 * Fetch all installed applications for a user
 */
function getInstalledApps(data) {
  const { username } = data;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Apps");
  const rows = sheet.getDataRange().getValues();
  const apps = [];

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().toLowerCase() === username.toLowerCase()) {
      apps.push({
        appId: rows[i][0],
        name: rows[i][2],
        icon: rows[i][3],
        type: rows[i][4],
        codeOrUrl: rows[i][5]
      });
    }
  }
  return jsonResponse({ status: "success", apps: apps });
}

/**
 * Delete an installed application
 */
function deleteInstalledApp(data) {
  const { appId, username } = data;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Apps");
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === appId && rows[i][1].toString().toLowerCase() === username.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ status: "success", message: "Application uninstalled." });
    }
  }
  return jsonResponse({ status: "error", message: "Application not found." });
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}