const SHEET_NAME = "PhotoTracker";

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  data.shift();

  const groups = data.map(row => ({
    group: String(row[0]),
    status: String(row[1] || "Pending")
  }));

  return jsonResponse({ groups });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents || "{}");
  const action = body.action || "complete";

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  if (action === "reset") {
    rows.forEach((row, i) => {
      sheet.getRange(i + 2, 2).setValue(i === 0 ? "Current" : "Pending");
    });
    return jsonResponse({ success: true });
  }

  if (action === "setCurrent") {
    const currentGroup = String(body.currentGroup);
    const currentIndex = rows.findIndex(row => String(row[0]) === currentGroup);

    if (currentIndex === -1) {
      return jsonResponse({ success: false, message: "Group not found" });
    }

    rows.forEach((row, i) => {
      let status = "Pending";
      if (i < currentIndex) status = "Done";
      if (i === currentIndex) status = "Current";
      sheet.getRange(i + 2, 2).setValue(status);
    });

    return jsonResponse({ success: true });
  }

  if (action === "complete") {
    const completedGroup = String(body.completedGroup);
    const completedIndex = rows.findIndex(row => String(row[0]) === completedGroup);

    if (completedIndex === -1) {
      return jsonResponse({ success: false, message: "Group not found" });
    }

    rows.forEach((row, i) => {
      let status = "Pending";
      if (i <= completedIndex) status = "Done";
      if (i === completedIndex + 1) status = "Current";
      sheet.getRange(i + 2, 2).setValue(status);
    });

    return jsonResponse({ success: true });
  }

  return jsonResponse({ success: false, message: "Unknown action" });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
