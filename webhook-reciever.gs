function doGet(e) {
  const issueKey = e.parameter.key;
  const jiraStatus = (e.parameter.status || "").toUpperCase();

  let finalStatus = null;

  // All statuses that should map to "In Progress"
  const toInProgress = [
    "BLOCKED / REJECTED",
    "TESTING",
    "IN-PROGRESS",
    "STORYBOOK",
    "PLANNING WEB",
    "PLANNING APP"
  ];

  if (toInProgress.includes(jiraStatus)) {
    finalStatus = "In Progress";
  } else if (jiraStatus === "UAT") {
    finalStatus = "Done";
  } else {
    return ContentService.createTextOutput("IGNORED");
  }

  // Check both sheets
  const sheetNames = ["Core", "Pattern"];
  const spreadsheet = SpreadsheetApp.getActive();

  for (const sheetName of sheetNames) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) continue;

    const result = updateSheetStatus(sheet, issueKey, finalStatus);
    if (result === "UPDATED") {
      return ContentService.createTextOutput("UPDATED");
    }
  }

  return ContentService.createTextOutput("KEY_NOT_FOUND");
}

function updateSheetStatus(sheet, issueKey, finalStatus) {
  // Find column indices by header names
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const findColumnIndex = (headerName) => {
    const index = headerRow.findIndex(h => h && h.toString().trim() === headerName);
    return index >= 0 ? index + 1 : null;
  };

  const webTaskCol = findColumnIndex("🌐 Task");
  const webStatusCol = findColumnIndex("🌐 Status");
  const mobileTaskCol = findColumnIndex("📱 Task");
  const mobileStatusCol = findColumnIndex("📱 Status");

  if (!webTaskCol || !webStatusCol || !mobileTaskCol || !mobileStatusCol) {
    return "COLUMNS_NOT_FOUND";
  }

  const lastRow = sheet.getLastRow();
  const webTaskData = sheet.getRange(1, webTaskCol, lastRow, 1).getValues();
  const mobileTaskData = sheet.getRange(1, mobileTaskCol, lastRow, 1).getValues();

  for (let i = 0; i < lastRow; i++) {
    if (webTaskData[i][0] === issueKey) {
      sheet.getRange(i + 1, webStatusCol).setValue(finalStatus);
      return "UPDATED";
    }
    if (mobileTaskData[i][0] === issueKey) {
      sheet.getRange(i + 1, mobileStatusCol).setValue(finalStatus);
      return "UPDATED";
    }
  }

  return "KEY_NOT_FOUND";
}