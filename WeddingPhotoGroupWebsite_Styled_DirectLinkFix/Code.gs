const SHEET_NAME = "PhotoTracker";

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = params.action || "status";
  const callback = params.callback || "";

  let result;

  if (action === "complete") {
    result = completeGroup(params.group);
  } else if (action === "setCurrent") {
    result = setCurrentGroup(params.group);
  } else if (action === "reset") {
    result = resetTracker();
  } else {
    result = getStatusObject();
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(result) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getStatusObject() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    return { success: false, message: "Sheet tab not found: " + SHEET_NAME, groups: [] };
  }

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return { success: false, message: "PhotoTracker sheet has no group rows.", groups: [] };
  }

  data.shift();

  const groups = data
    .filter(row => row[0] !== "" && row[0] !== null && row[0] !== undefined)
    .map(row => ({
      group: String(row[0]).trim(),
      status: String(row[1] || "Pending").trim()
    }));

  return { success: true, groups };
}

function completeGroup(groupNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    return { success: false, message: "Sheet tab not found: " + SHEET_NAME, groups: [] };
  }

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  const completedGroup = String(groupNumber || "").trim();
  const completedIndex = rows.findIndex(row => String(row[0]).trim() === completedGroup);

  if (completedIndex === -1) {
    return { success: false, message: "Group not found: " + completedGroup, groups: getStatusObject().groups };
  }

  rows.forEach((row, i) => {
    let status = "Pending";
    if (i <= completedIndex) status = "Done";
    if (i === completedIndex + 1) status = "Current";
    sheet.getRange(i + 2, 2).setValue(status);
  });

  SpreadsheetApp.flush();
  return getStatusObject();
}

function setCurrentGroup(groupNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    return { success: false, message: "Sheet tab not found: " + SHEET_NAME, groups: [] };
  }

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  const currentGroup = String(groupNumber || "").trim();
  const currentIndex = rows.findIndex(row => String(row[0]).trim() === currentGroup);

  if (currentIndex === -1) {
    return { success: false, message: "Group not found: " + currentGroup, groups: getStatusObject().groups };
  }

  rows.forEach((row, i) => {
    let status = "Pending";
    if (i < currentIndex) status = "Done";
    if (i === currentIndex) status = "Current";
    sheet.getRange(i + 2, 2).setValue(status);
  });

  SpreadsheetApp.flush();
  return getStatusObject();
}

function resetTracker() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    return { success: false, message: "Sheet tab not found: " + SHEET_NAME, groups: [] };
  }

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  rows.forEach((row, i) => {
    sheet.getRange(i + 2, 2).setValue(i === 0 ? "Current" : "Pending");
  });

  SpreadsheetApp.flush();
  return getStatusObject();
}
