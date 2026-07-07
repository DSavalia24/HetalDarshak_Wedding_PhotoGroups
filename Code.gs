const SHEET_NAME = "PhotoTracker";

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = params.action || "status";
  const callback = params.callback || "";

  if (action === "guestStatus") {
    return guestStatusHtml(params.groups || "");
  }

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

function guestStatusHtml(groupText) {
  const status = getGuestStatus(groupText);
  const bg = status.className === "now" ? "#e5ffed" :
             status.className === "next" ? "#fff8e2" :
             status.className === "done" ? "#efefe8" : "#fff4eb";
  const border = status.className === "now" ? "rgba(47,125,78,.35)" :
                 status.className === "next" ? "rgba(188,140,55,.38)" :
                 status.className === "done" ? "rgba(123,123,99,.35)" : "rgba(201,147,98,.45)";
  const titleColor = status.className === "now" ? "#2f7d4e" :
                     status.className === "next" ? "#8b6427" :
                     status.className === "done" ? "#66664f" : "#4a2d22";

  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta http-equiv="refresh" content="5">
  <style>
    html, body {
      margin: 0;
      padding: 0;
      background: transparent;
      font-family: Georgia, "Times New Roman", serif;
      color: #3b2a24;
      overflow: hidden;
    }
    .box {
      box-sizing: border-box;
      width: 100%;
      min-height: 106px;
      padding: 15px 14px;
      border-radius: 18px;
      border: 1px solid ${border};
      background: ${bg};
      text-align: center;
    }
    .kicker {
      margin: 0 0 5px;
      font-size: 12px;
      letter-spacing: 1.3px;
      text-transform: uppercase;
      color: #8c5a3a;
    }
    h3 {
      margin: 0 0 5px;
      font-size: 23px;
      color: ${titleColor};
      line-height: 1.15;
    }
    p {
      margin: 0;
      font-size: 15px;
      line-height: 1.35;
      color: #5e4840;
    }
  </style>
</head>
<body>
  <div class="box">
    <p class="kicker">${escapeHtml(status.highlight)}</p>
    <h3>${escapeHtml(status.title)}</h3>
    <p>${escapeHtml(status.message)}</p>
  </div>
</body>
</html>`;

  return HtmlService
    .createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getGuestStatus(groupText) {
  const statusObj = getStatusObject();

  if (!statusObj.success || !statusObj.groups.length) {
    return {
      className: "pending",
      highlight: "",
      title: "Photo tracker loading",
      message: "Your group numbers are listed above. Please listen for the coordinator."
    };
  }

  const trackerGroups = statusObj.groups;
  const current = trackerGroups.find(g => String(g.status).toLowerCase() === "current");

  if (!current) {
    return {
      className: "done",
      highlight: "All Done",
      title: "Photos are complete",
      message: "All photo groups appear to be finished."
    };
  }

  const currentGroup = Number(current.group);
  const guestGroups = String(groupText || "")
    .split(",")
    .map(g => Number(String(g).trim()))
    .filter(n => !isNaN(n));

  if (guestGroups.includes(currentGroup)) {
    return {
      className: "now",
      highlight: "Now Taking Group " + currentGroup,
      title: "You are up now!",
      message: "Please head to the photo area now."
    };
  }

  const alreadyDone = guestGroups.length > 0 && guestGroups.every(n => n < currentGroup);
  if (alreadyDone) {
    return {
      className: "done",
      highlight: "Completed",
      title: "Your group photos are done",
      message: "Thank you! Your listed photo groups have already been completed."
    };
  }

  const upcoming = guestGroups.filter(n => n > currentGroup).sort((a, b) => a - b);
  const nextGuestGroup = upcoming[0];

  if (nextGuestGroup) {
    const groupsAway = nextGuestGroup - currentGroup;

    if (groupsAway === 1) {
      return {
        className: "next",
        highlight: "Currently: Group " + currentGroup + " • Your next group: " + nextGuestGroup,
        title: "You are next",
        message: "Please stay close to the photo area."
      };
    }

    if (groupsAway <= 3) {
      return {
        className: "pending",
        highlight: "Currently: Group " + currentGroup + " • Your next group: " + nextGuestGroup,
        title: "You are coming up soon",
        message: "Please stay nearby."
      };
    }

    return {
      className: "pending",
      highlight: "Currently: Group " + currentGroup + " • Your next group: " + nextGuestGroup,
      title: "You are not up yet",
      message: "Please relax for now. We will call your group soon."
    };
  }

  return {
    className: "pending",
    highlight: "Current Group " + currentGroup,
    title: "Please check with the coordinator",
    message: "Your groups could not be compared to the live tracker."
  };
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
