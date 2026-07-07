const SHEET_NAME = "PhotoTracker";

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = params.action || "status";

  if (action === "complete") {
    return completeGroup(params.group);
  }

  if (action === "setCurrent") {
    return setCurrentGroup(params.group);
  }

  if (action === "reset") {
    return resetTracker();
  }

  return getStatus();
}

function getStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  data.shift();

  const groups = data.map(row => ({
    group: String(row[0]),
    status: String(row[1] || "Pending")
  }));

  return jsonResponse({ success: true, groups });
}

function completeGroup(groupNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  const completedGroup = String(groupNumber);
  const completedIndex = rows.findIndex(row => String(row[0]) === completedGroup);

  if (completedIndex === -1) {
    return jsonResponse({ success: false, message: "Group not found: " + completedGroup });
  }

  rows.forEach((row, i) => {
    let status = "Pending";
    if (i <= completedIndex) status = "Done";
    if (i === completedIndex + 1) status = "Current";
    sheet.getRange(i + 2, 2).setValue(status);
  });

  return getStatus();
}

function setCurrentGroup(groupNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  const currentGroup = String(groupNumber);
  const currentIndex = rows.findIndex(row => String(row[0]) === currentGroup);

  if (currentIndex === -1) {
    return jsonResponse({ success: false, message: "Group not found: " + currentGroup });
  }

  rows.forEach((row, i) => {
    let status = "Pending";
    if (i < currentIndex) status = "Done";
    if (i === currentIndex) status = "Current";
    sheet.getRange(i + 2, 2).setValue(status);
  });

  return getStatus();
}

function resetTracker() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  rows.forEach((row, i) => {
    sheet.getRange(i + 2, 2).setValue(i === 0 ? "Current" : "Pending");
  });

  return getStatus();
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/*
TEST LINKS AFTER DEPLOYING:
1. Status:
   YOUR_WEB_APP_URL?action=status

2. Complete Group 1:
   YOUR_WEB_APP_URL?action=complete&group=1

3. Set Group 5 Current:
   YOUR_WEB_APP_URL?action=setCurrent&group=5

4. Reset:
   YOUR_WEB_APP_URL?action=reset
*/
