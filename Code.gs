/**
 * SPARROW FARMS - SECURE BACKEND WITH TASK MANAGEMENT
 */

const SECRETS = {
  NOTIFY_EMAILS: "sparrowfarmhay@gmail.com,tina@sparrowfarms.ca,dave@sparrowfarms.ca"
};

// ===== SHARED REQUEST HANDLER =====
// Both doGet (legacy JSONP) and doPost (new fetch) route through here.

function handleRequest(params) {
  let result = { error: "Initial State" };

  try {
    const pin = String(params.pin || "").trim();
    const action = params.action;
    const userSheet = findSheetByPin(pin);

    if (!userSheet) {
      result = { error: "PIN not found" };
    } else {
      const isAdmin = (String(userSheet.getRange("H3").getValue()).toUpperCase() === "ADMIN");

      if (action === "getWeeklyHistory") {
        result = getWeeklyHistory(pin);
      } else if (action === "handleClockAction") {
        result = handleClockAction(pin, params.type, params.lunchMinutes);
      } else if (action === "getAdminLogs") {
        result = isAdmin ? getAdminLogs() : { error: "Unauthorized" };
      } else if (action === "recalculate") {
        result = isAdmin ? recalculateAllSheets() : { error: "Unauthorized" };
      } else if (action === "sendLowStockAlert") {
        sendLowStockAlert(pin, params.item);
        result = { success: true };
      } else if (action === "sendEquipmentIssue") {
        sendEquipmentIssue(pin, params.machine, params.issue);
        result = { success: true };
      } else if (action === "submitShiftReport") {
        submitShiftReport(pin, params.work);
        result = { success: true };
      } else if (action === "assignTasks") {
        result = isAdmin ? assignTasks(params.targetPin, params.tasks) : { error: "Unauthorized" };
      } else if (action === "getTasks") {
        result = getTasks(pin);
      } else if (action === "getStaffList") {
        result = isAdmin ? getStaffList() : { error: "Unauthorized" };
      } else if (action === "getAllTasks") {
        result = isAdmin ? getAllTasks() : { error: "Unauthorized" };
      } else if (action === "deleteTask") {
        result = isAdmin ? deleteTask(params.rowIndex) : { error: "Unauthorized" };
      } else if (action === "editTask") {
        result = isAdmin ? editTask(params.rowIndex, params.newTask) : { error: "Unauthorized" };
      } else if (action === "reassignTask") {
        result = isAdmin ? reassignTask(params.rowIndex, params.newPin) : { error: "Unauthorized" };
      } else if (action === "addStaffMember") {
        result = isAdmin ? addStaffMember(params.name, params.phone, params.email, params.newPin, params.rate) : { error: "Unauthorized" };
      }
    }
  } catch (err) {
    result = { error: "Server Error: " + err.toString() };
  }

  return result;
}

// ===== doGet — legacy JSONP support (kept for transition) =====

function doGet(e) {
  const callback = (e.parameter && e.parameter.callback) ? e.parameter.callback : "callback";
  const result = handleRequest(e.parameter || {});
  const output = callback + "(" + JSON.stringify(result) + ")";
  return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// ===== doPost — new fetch() support =====

function doPost(e) {
  let params = {};
  try {
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: "Invalid JSON in request body" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const result = handleRequest(params);
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== CORE FUNCTIONS =====

function recalculateAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const pin = sheet.getRange("D3").getValue();
    if (!pin || sheet.getName().includes("Log") || sheet.getName().includes("Tasks")) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 7) return;

    const rate = parseFloat(sheet.getRange("E3").getValue()) || 0;
    const data = sheet.getRange(7, 1, lastRow - 6, 9).getValues();

    data.forEach((row, index) => {
      const rowIndex = index + 7;
      const timeIn = row[5];  // Col F
      const lunchStr = String(row[6] || "0"); // Col G
      const timeOut = row[8]; // Col I

      if (timeIn instanceof Date && timeOut instanceof Date) {
        const lunchMins = parseInt(lunchStr.replace(/[^0-9]/g, "")) || 0;
        let diffMs = timeOut - timeIn;
        let netMs = diffMs - (lunchMins * 60 * 1000);
        const hours = Math.max(0, netMs / (1000 * 60 * 60));
        const earned = hours * rate;
        sheet.getRange(rowIndex, 4).setValue(hours.toFixed(2));  // Col D
        sheet.getRange(rowIndex, 5).setValue(earned.toFixed(2)); // Col E
      }
    });
  });

  return { success: true, msg: "Recalculation complete." };
}

function getWeeklyHistory(pin) {
  const userSheet = findSheetByPin(pin);
  if (!userSheet) return { error: "Not found" };

  const h3Value = String(userSheet.getRange("H3").getValue() || "").trim().toUpperCase();
  const isAdmin = (h3Value === "ADMIN");

  let status = "Clocked Out";
  const lastRow = userSheet.getLastRow();
  if (lastRow >= 7) {
    const timeIn = userSheet.getRange(lastRow, 6).getValue();
    const timeOut = userSheet.getRange(lastRow, 9).getValue();
    if (timeIn && !timeOut) status = "Clocked In";
  }

  const rate = userSheet.getRange("E3").getValue();

  // Fetch cols A-K (11 columns) to include Pay (col E, index 4) and Pay Date (col K, index 10)
  const data = userSheet.getLastRow() < 7 ? [] : userSheet.getRange(7, 1, userSheet.getLastRow() - 6, 11).getValues();

  let html = "<table style='width:100%; font-size: 13px; border-collapse: collapse;'>";
  html += "<tr style='background:#eee; font-size:11px;'><th style='padding:5px; text-align:left;'>Date</th><th style='padding:5px; text-align:right;'>Hrs</th><th style='padding:5px; text-align:right;'>Pay</th><th style='padding:5px; text-align:center;'>Date Paid</th></tr>";

  data.slice(-10).reverse().forEach(row => {
    const dateStr = row[0] ? new Date(row[0]).toLocaleDateString() : '';
    const hours = parseFloat(row[3] || 0).toFixed(2);
    const pay = parseFloat(row[4] || 0).toFixed(2);
    const payDateVal = row[10];
    const isPending = String(payDateVal).toLowerCase() === "pending";
    const isPaid = payDateVal instanceof Date;
    const paidLabel = isPaid
      ? `<span style='color:#388e3c; font-size:11px;'>✓ ${new Date(payDateVal).toLocaleDateString()}</span>`
      : isPending
        ? `<span style='color:#f57c00; font-size:11px;'>Pending</span>`
        : `<span style='color:#bbb; font-size:11px;'>-</span>`;

    html += `<tr style='border-bottom: 1px solid #eee;'>
      <td style='padding:5px;'>${dateStr}</td>
      <td style='padding:5px; text-align:right;'>${hours}</td>
      <td style='padding:5px; text-align:right;'>$${pay}</td>
      <td style='padding:5px; text-align:center;'>${paidLabel}</td>
    </tr>`;
  });

  html += "</table>";

  // Compute unpaid balance (pending rows only)
  const pendingRows = data.filter(row => String(row[10]).toLowerCase() === "pending");
  const unpaidHours = pendingRows.reduce((sum, row) => sum + (parseFloat(row[3]) || 0), 0);
  const unpaidPay = pendingRows.reduce((sum, row) => sum + (parseFloat(row[4]) || 0), 0);

  const tasks = getTasks(pin);
  return {
    name: userSheet.getRange("A3").getValue(),
    currentStatus: status,
    pay: isAdmin ? "Administrator Access" : `Current Rate: $${rate}/hr`,
    history: html,
    unpaidSummary: isAdmin ? null : { hours: unpaidHours.toFixed(2), pay: unpaidPay.toFixed(2) },
    isAdmin: isAdmin,
    tasks: tasks.tasks || []
  };
}

function getAdminLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let html = "";

  // 1. PULL RECENT ALERTS (Last 3 Stock & Maint)
  html += "<div style='text-align:left;'><b>Recent Alerts:</b>";
  html += "<div style='background:#fffcf0; padding:10px; border-radius:8px; margin:10px 0; border:1px solid #ffd54f; font-size:12px;'>";

  const sLog = ss.getSheetByName("Low Stock Log");
  if (sLog && sLog.getLastRow() >= 3) {
    const sStart = Math.max(3, sLog.getLastRow() - 2);
    const sNumRows = (sLog.getLastRow() - sStart) + 1;
    const sData = sLog.getRange(sStart, 1, sNumRows, 3).getValues().reverse();
    sData.forEach(row => {
      if (row[0] && row[2] && row[2] !== "Item") {
        html += `🚩 <b>STOCK:</b> ${row[2]} (by ${row[1] || "Unknown"})<br>`;
      }
    });
  } else {
    html += "<i>No recent stock alerts.</i><br>";
  }

  html += "<hr style='border:0; border-top:1px solid #ffe082; margin:5px 0;'>";

  const mLog = ss.getSheetByName("Maintenance Log");
  if (mLog && mLog.getLastRow() >= 3) {
    const mStart = Math.max(3, mLog.getLastRow() - 2);
    const mNumRows = (mLog.getLastRow() - mStart) + 1;
    const mData = mLog.getRange(mStart, 1, mNumRows, 4).getValues().reverse();
    mData.forEach(row => {
      if (row[0] && row[2] && row[2] !== "Machine") {
        html += `🛠️ <b>MAINT:</b> ${row[2]} - ${row[3]} (by ${row[1] || "Unknown"})<br>`;
      }
    });
  } else {
    html += "<i>No maintenance issues.</i>";
  }

  html += "</div></div>";

  // 2. PULL STAFF ACTIVITY (Last 7 Days)
  let allLogs = [];
  const oneWeekAgo = new Date();
  oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);

  sheets.forEach(sheet => {
    const pin = sheet.getRange("D3").getValue();
    if (!pin || sheet.getName().includes("Log") || sheet.getName().includes("Tasks")) return;

    const name = sheet.getRange("A3").getValue();
    const lastRow = sheet.getLastRow();
    if (lastRow < 7) return;

    const data = sheet.getRange(7, 1, lastRow - 6, 9).getValues();
    data.forEach(row => {
      const dateVal = new Date(row[0]);
      if (row[0] instanceof Date && row[5] instanceof Date && dateVal >= oneWeekAgo) {
        allLogs.push({
          name: name,
          work: row[1] || "---",
          date: dateVal.toLocaleDateString(),
          in: row[5].toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
          out: row[8] instanceof Date ? row[8].toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : "Working..."
        });
      }
    });
  });

  allLogs.sort((a, b) => new Date(b.date) - new Date(a.date));

  html += "<div style='text-align:left;'><b>Staff Activity:</b></div>";
  html += "<table style='width:100%; font-size:10px; border-collapse:collapse; margin-top:10px;'>";
  html += "<tr style='background:#eee;'><th>Staff</th><th>Date</th><th>In/Out</th><th>Work Done</th></tr>";

  allLogs.slice(0, 20).forEach(log => {
    html += `<tr style='border-bottom:1px solid #ddd; vertical-align:top;'>`;
    html += `<td style='padding:5px;'><b>${log.name}</b></td>`;
    html += `<td style='padding:5px;'>${log.date}</td>`;
    html += `<td style='padding:5px;'>${log.in}<br>${log.out}</td>`;
    html += `<td style='padding:5px; font-style:italic; color:#444;'>${log.work}</td>`;
    html += `</tr>`;
  });

  return html + "</table>";
}

function findSheetByPin(pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const searchPin = String(pin || "").trim();
  if (!searchPin) return null;

  for (let i = 0; i < sheets.length; i++) {
    try {
      const cellValue = String(sheets[i].getRange("D3").getValue() || "").trim();
      if (cellValue === searchPin) return sheets[i];
    } catch(e) { continue; }
  }
  return null;
}

function handleClockAction(pin, type, lunchMinutes) {
  const userSheet = findSheetByPin(pin);
  if (!userSheet) return { error: "Invalid PIN" };

  const name = userSheet.getRange("A3").getValue();
  const isAdmin = (String(userSheet.getRange("H3").getValue()).toUpperCase() === "ADMIN");
  const now = new Date();
  const currentRate = userSheet.getRange("E3").getValue();

  if (type === 'clock-in') {
    userSheet.appendRow([now.toLocaleDateString(), "Pending...", "", "", "", now, "", "", "", "", ""]);
    return { name: name, msg: "Clocked in at " + now.toLocaleTimeString(), status: "Clocked In", isAdmin: isAdmin };
  }

  const lastRow = userSheet.getLastRow();
  const targetRow = Math.max(lastRow, 7);

  if (type === 'clock-out') {
    userSheet.getRange(targetRow, 9).setValue(now);
    userSheet.getRange(targetRow, 11).setValue("pending"); // Col K - awaiting payment

    const timeInValue = userSheet.getRange(targetRow, 6).getValue();
    const timeIn = new Date(timeInValue);
    let totalMs = now - timeIn;
    const lunchMs = (parseInt(lunchMinutes) || 0) * 60 * 1000;
    totalMs = totalMs - lunchMs;
    const hours = (totalMs / (1000 * 60 * 60)).toFixed(2);
    const earned = (hours * currentRate).toFixed(2);

    userSheet.getRange(targetRow, 7).setValue(lunchMinutes + " mins");
    userSheet.getRange(targetRow, 4).setValue(hours);
    userSheet.getRange(targetRow, 5).setValue(earned);

    const tasks = getTasks(pin);
    let tasksPrefill = "";
    if (tasks.tasks && tasks.tasks.length > 0) {
      tasksPrefill = tasks.tasks.map(t => `${t} (STATUS)`).join("\n");
    }

    return {
      name: name,
      status: "Clocked Out",
      showReport: true,
      summary: `Total: ${hours} hrs (Break: ${lunchMinutes}m) | Earned: $${earned}`,
      isAdmin: isAdmin,
      tasksPrefill: tasksPrefill,
      hasTasks: tasks.tasks && tasks.tasks.length > 0
    };
  }
}

function sendLowStockAlert(pin, item) {
  const userSheet = findSheetByPin(pin);
  const name = userSheet ? userSheet.getRange("A3").getValue() : "Unknown";
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Low Stock Log");
  if (log) log.appendRow([new Date(), name, item]);
  MailApp.sendEmail(SECRETS.NOTIFY_EMAILS, `Low Stock: ${item}`, `Staff: ${name}\nItem: ${item}`);
}

function sendEquipmentIssue(pin, machine, issue) {
  const userSheet = findSheetByPin(pin);
  const name = userSheet ? userSheet.getRange("A3").getValue() : "Unknown";
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Maintenance Log");
  if (log) log.appendRow([new Date(), name, machine, issue]);
  MailApp.sendEmail(SECRETS.NOTIFY_EMAILS, `Maintenance: ${machine}`, `Staff: ${name}\nMachine: ${machine}\nIssue: ${issue}`);
}

function submitShiftReport(pin, work) {
  const userSheet = findSheetByPin(pin);
  if (userSheet) {
    const name = userSheet.getRange("A3").getValue();
    userSheet.getRange(userSheet.getLastRow(), 2).setValue(work);
    userSheet.getRange(userSheet.getLastRow(), 10).setValue(new Date());
    MailApp.sendEmail(SECRETS.NOTIFY_EMAILS, `Shift Report: ${name}`, `Staff: ${name}\nWork:\n${work}`);
    clearCompletedTasks(pin);
  }
}

function sendWeeklySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  let emailBody = "<html><body style='font-family: Arial, sans-serif;'>";
  emailBody += "<h2 style='color: #2b3d2b;'>Weekly Payroll Summary - Sparrow Farms</h2>";
  emailBody += "<p>Here's the summary of hours to be paid this week:</p>";

  let totalOwed = 0;
  let hasData = false;
  const today = new Date();
  const thirteenDaysAgo = new Date();
  thirteenDaysAgo.setDate(today.getDate() - 13);

  sheets.forEach(sheet => {
    const pin = sheet.getRange("D3").getValue();
    if (!pin || sheet.getName().includes("Log") || sheet.getName().includes("Tasks")) return;

    const name = sheet.getRange("A3").getValue();
    const email = sheet.getRange("C3").getValue();
    const rate = parseFloat(sheet.getRange("E3").getValue()) || 0;
    const lastRow = sheet.getLastRow();
    if (lastRow < 7) return;

    // Fetch cols A-K (11 columns) to check Pay (col E) and Pay Date (col K)
    const data = sheet.getRange(7, 1, lastRow - 6, 11).getValues();

    // Pending rows = col K (index 10) contains the string "pending"
    const pendingRows = data.filter(row =>
      row[0] instanceof Date &&
      (parseFloat(row[3]) || 0) > 0 &&
      String(row[10]).toLowerCase() === "pending"
    );

    if (pendingRows.length === 0) return;

    const pendingHours = pendingRows.reduce((sum, row) => sum + (parseFloat(row[3]) || 0), 0);
    const pendingEarned = pendingRows.reduce((sum, row) => sum + (parseFloat(row[4]) || 0), 0);

    // Find the most recent actual pay date (a real Date value, not "pending", not blank)
    const lastPayDate = data
      .map(row => row[10])
      .filter(val => val instanceof Date)
      .sort((a, b) => b - a)[0] || null;

    // Force include if: hours < 3.0 AND there is a prior pay date AND it was > 13 days ago
    const forceInclude = pendingHours < 3.0 && lastPayDate && lastPayDate < thirteenDaysAgo;

    // Skip if: hours < 3.0 AND we are NOT in the force-include situation
    const skip = pendingHours < 3.0 && !forceInclude;

    if (skip) return;

    hasData = true;
    totalOwed += pendingEarned;

    emailBody += "<div style='background: #f9fbf9; padding: 15px; margin: 15px 0; border-left: 4px solid #388e3c; border-radius: 5px;'>";
    emailBody += `<h3 style='margin-top: 0; color: #2b3d2b;'>${name}`;

    if (forceInclude) {
      emailBody += ` <span style='background:#e65100; color:white; font-size:12px; padding:3px 8px; border-radius:10px; font-weight:bold; vertical-align:middle;'>⚠️ ROLLOVER - 2 weeks combined</span>`;
    }

    emailBody += `</h3>`;
    emailBody += `<p><strong>Email:</strong> ${email}<br>`;
    emailBody += `<strong>Rate:</strong> $${rate.toFixed(2)}/hr<br>`;
    emailBody += `<strong>Unpaid Hours:</strong> ${pendingHours.toFixed(2)}<br>`;
    emailBody += `<strong>Amount to Send:</strong> <span style='color: #388e3c; font-size: 1.2em;'>$${pendingEarned.toFixed(2)}</span></p>`;

    if (forceInclude) {
      emailBody += `<p style='background:#fff3e0; padding:8px 12px; border-radius:6px; font-size:13px; color:#e65100; margin-bottom:12px;'>
        Last payment was on ${lastPayDate.toLocaleDateString()} - more than 13 days ago.
        This payout covers all pending shifts regardless of total hours.
      </p>`;
    }

    emailBody += "<table style='width: 100%; border-collapse: collapse; font-size: 0.9em;'>";
    emailBody += "<tr style='background: #e8f5e9;'><th style='padding: 8px; text-align: left;'>Date</th><th style='padding: 8px; text-align: right;'>Hours</th><th style='padding: 8px; text-align: right;'>Earned</th><th style='padding: 8px; text-align: left;'>Work Done</th></tr>";

    pendingRows.forEach(row => {
      emailBody += `<tr style='border-bottom: 1px solid #ddd;'>`;
      emailBody += `<td style='padding: 8px;'>${new Date(row[0]).toLocaleDateString()}</td>`;
      emailBody += `<td style='padding: 8px; text-align: right;'>${(parseFloat(row[3]) || 0).toFixed(2)}</td>`;
      emailBody += `<td style='padding: 8px; text-align: right;'>$${(parseFloat(row[4]) || 0).toFixed(2)}</td>`;
      emailBody += `<td style='padding: 8px; font-style: italic; color: #666;'>${row[1] || "No description"}</td>`;
      emailBody += `</tr>`;
    });

    emailBody += "</table></div>";
  });

  if (!hasData) {
    emailBody += "<p><em>No payments to process this week - all staff either have no pending shifts, or have under 3 unpaid hours and are within their first eligible skip.</em></p>";
  } else {
    emailBody += `<div style='background: #2b3d2b; color: white; padding: 20px; margin: 20px 0; border-radius: 5px; text-align: center;'>`;
    emailBody += `<h3 style='margin: 0;'>Total to Send This Week</h3>`;
    emailBody += `<p style='font-size: 2em; margin: 10px 0;'>$${totalOwed.toFixed(2)}</p>`;
    emailBody += `</div>`;
  }

  emailBody += "<p style='color: #666; font-size: 0.9em;'>This is an automated weekly summary from your Sparrow Farms time tracking system.</p>";
  emailBody += "</body></html>";

  MailApp.sendEmail({
    to: SECRETS.NOTIFY_EMAILS,
    subject: `Weekly Payroll Summary - ${today.toLocaleDateString()}`,
    htmlBody: emailBody
  });
}

// ===== TASK MANAGEMENT FUNCTIONS =====

function getStaffList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let staff = [];

  sheets.forEach(sheet => {
    const pin = sheet.getRange("D3").getValue();
    if (!pin || sheet.getName().includes("Log") || sheet.getName().includes("Tasks")) return;

    const name = sheet.getRange("A3").getValue();
    const h3Value = String(sheet.getRange("H3").getValue() || "").trim().toUpperCase();
    const isAdmin = (h3Value === "ADMIN");

    if (!isAdmin) {
      staff.push({ name: name, pin: String(pin) });
    }
  });

  return { staff: staff };
}

function assignTasks(targetPin, tasksText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tasksSheet = ss.getSheetByName("Tasks");

  if (!tasksSheet) {
    tasksSheet = ss.insertSheet("Tasks");
    tasksSheet.appendRow(["PIN", "Task", "Date Assigned", "Status"]);
  }

  const targetSheet = findSheetByPin(targetPin);
  if (!targetSheet) return { error: "Employee not found" };

  const tasks = tasksText.split("\n").filter(t => t.trim());
  const now = new Date();

  tasks.forEach(task => {
    if (task.trim()) {
      tasksSheet.appendRow([targetPin, task.trim(), now, "Active"]);
    }
  });

  return { success: true, msg: `${tasks.length} task(s) assigned` };
}

function getTasks(pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName("Tasks");

  if (!tasksSheet || tasksSheet.getLastRow() < 2) {
    return { tasks: [] };
  }

  const data = tasksSheet.getRange(2, 1, tasksSheet.getLastRow() - 1, 4).getValues();
  const activeTasks = [];

  data.forEach(row => {
    if (String(row[0]).trim() === String(pin).trim() && row[3] === "Active") {
      activeTasks.push(row[1]);
    }
  });

  return { tasks: activeTasks };
}

function clearCompletedTasks(pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName("Tasks");
  if (!tasksSheet || tasksSheet.getLastRow() < 2) return;

  const lastRow = tasksSheet.getLastRow();

  for (let i = lastRow; i >= 2; i--) {
    const taskPin = String(tasksSheet.getRange(i, 1).getValue()).trim();
    const status = tasksSheet.getRange(i, 4).getValue();
    if (taskPin === String(pin).trim() && status === "Active") {
      tasksSheet.getRange(i, 4).setValue("Completed");
      tasksSheet.getRange(i, 5).setValue(new Date());
    }
  }
}

function getAllTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName("Tasks");

  if (!tasksSheet || tasksSheet.getLastRow() < 2) {
    return { tasks: [] };
  }

  const data = tasksSheet.getRange(2, 1, tasksSheet.getLastRow() - 1, 5).getValues();
  const allTasks = [];

  data.forEach((row, index) => {
    const userSheet = findSheetByPin(String(row[0]));
    const name = userSheet ? userSheet.getRange("A3").getValue() : "Unknown";
    allTasks.push({
      rowIndex: index + 2,
      pin: String(row[0]),
      name: name,
      task: row[1],
      dateAssigned: row[2] instanceof Date ? row[2].toLocaleDateString() : row[2],
      status: row[3],
      dateCompleted: row[4] instanceof Date ? row[4].toLocaleDateString() : (row[4] || "")
    });
  });

  allTasks.sort((a, b) => {
    if (a.status === "Active" && b.status !== "Active") return -1;
    if (a.status !== "Active" && b.status === "Active") return 1;
    return new Date(b.dateAssigned) - new Date(a.dateAssigned);
  });

  return { tasks: allTasks };
}

function deleteTask(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName("Tasks");

  if (!tasksSheet || !rowIndex || rowIndex < 2) {
    return { error: "Invalid task" };
  }

  tasksSheet.deleteRow(parseInt(rowIndex));
  return { success: true, msg: "Task deleted" };
}

function editTask(rowIndex, newTask) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName("Tasks");

  if (!tasksSheet || !rowIndex || rowIndex < 2 || !newTask) {
    return { error: "Invalid parameters" };
  }

  tasksSheet.getRange(parseInt(rowIndex), 2).setValue(newTask.trim());
  return { success: true, msg: "Task updated" };
}

function reassignTask(rowIndex, newPin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName("Tasks");

  if (!tasksSheet || !rowIndex || rowIndex < 2 || !newPin) {
    return { error: "Invalid parameters" };
  }

  const row = parseInt(rowIndex);
  const originalTask = tasksSheet.getRange(row, 2).getValue();

  tasksSheet.getRange(row, 4).setValue("Reassigned");
  tasksSheet.getRange(row, 5).setValue(new Date());

  tasksSheet.appendRow([
    String(newPin),
    originalTask,
    new Date(),
    "Active",
    ""
  ]);

  return { success: true, msg: "Task reassigned successfully" };
}

function addStaffMember(name, phone, email, pin, rate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!name || !pin || !rate) {
    return { error: "Name, PIN, and Rate are required" };
  }

  const existingSheet = findSheetByPin(String(pin).trim());
  if (existingSheet) {
    return { error: "PIN already exists. Please choose a different PIN." };
  }

  const templateSheet = ss.getSheetByName("TEMPLATE");
  if (!templateSheet) {
    return { error: "TEMPLATE sheet not found. Please create a template sheet first." };
  }

  const newSheet = templateSheet.copyTo(ss);
  newSheet.setName(name);

  newSheet.getRange("A3").setValue(name);
  newSheet.getRange("B3").setValue(phone || "");
  newSheet.getRange("C3").setValue(email || "");
  newSheet.getRange("D3").setValue(String(pin));
  newSheet.getRange("E3").setValue(parseFloat(rate));
  newSheet.getRange("F3").setValue(new Date());
  newSheet.getRange("G3").setValue("");

  newSheet.activate();
  ss.moveActiveSheet(1);

  return { success: true, msg: "Staff member added: " + name };
}