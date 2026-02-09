/** * SPARROW FARMS - SECURE BACKEND WITH TASK MANAGEMENT */

const SECRETS = {
  NOTIFY_EMAILS: "sparrowfarms@gmail.com,tina@sparrowfarms.ca, dave@sparrowfarms.ca" 
};

function doGet(e) {
  const callback = (e.parameter && e.parameter.callback) ? e.parameter.callback : "callback";
  let result = { error: "Initial State" }; 

  try {
    const pin = String(e.parameter.pin || "").trim();
    const action = e.parameter.action;
    const userSheet = findSheetByPin(pin);

    if (!userSheet) {
      result = { error: "PIN not found" };
    } else {
      const isAdmin = (String(userSheet.getRange("H3").getValue()).toUpperCase() === "ADMIN");

      if (action === "getWeeklyHistory") {
        result = getWeeklyHistory(pin);
      } else if (action === "handleClockAction") {
        result = handleClockAction(pin, e.parameter.type, e.parameter.lunchMinutes);
      } else if (action === "getAdminLogs") {
        result = isAdmin ? getAdminLogs() : { error: "Unauthorized" };
      } else if (action === "recalculate") {
        result = isAdmin ? recalculateAllSheets() : { error: "Unauthorized" };
      } else if (action === "sendLowStockAlert") {
        sendLowStockAlert(pin, e.parameter.item);
        result = { success: true };
      } else if (action === "sendEquipmentIssue") {
        sendEquipmentIssue(pin, e.parameter.machine, e.parameter.issue);
        result = { success: true };
      } else if (action === "submitShiftReport") {
        submitShiftReport(pin, e.parameter.work);
        result = { success: true };
      } else if (action === "assignTasks") {
        result = isAdmin ? assignTasks(e.parameter.targetPin, e.parameter.tasks) : { error: "Unauthorized" };
      } else if (action === "getTasks") {
        result = getTasks(pin);
      } else if (action === "getStaffList") {
        result = isAdmin ? getStaffList() : { error: "Unauthorized" };
      } else if (action === "getAllTasks") {
        result = isAdmin ? getAllTasks() : { error: "Unauthorized" };
      } else if (action === "deleteTask") {
        result = isAdmin ? deleteTask(e.parameter.rowIndex) : { error: "Unauthorized" };
      } else if (action === "editTask") {
        result = isAdmin ? editTask(e.parameter.rowIndex, e.parameter.newTask) : { error: "Unauthorized" };
      } else if (action === "reassignTask") {
        result = isAdmin ? reassignTask(e.parameter.rowIndex, e.parameter.newPin) : { error: "Unauthorized" };
      } else if (action === "addStaffMember") {
        result = isAdmin ? addStaffMember(e.parameter.name, e.parameter.phone, e.parameter.email, e.parameter.newPin, e.parameter.rate) : { error: "Unauthorized" };
      }
    }
  } catch (err) {
    result = { error: "Server Error: " + err.toString() };
  }

  const output = callback + "(" + JSON.stringify(result) + ")";
  return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

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
      const timeIn = row[5]; // Col F
      const lunchStr = String(row[6] || "0"); // Col G
      const timeOut = row[8]; // Col I

      if (timeIn instanceof Date && timeOut instanceof Date) {
        const lunchMins = parseInt(lunchStr.replace(/[^0-9]/g, "")) || 0;
        let diffMs = timeOut - timeIn;
        let netMs = diffMs - (lunchMins * 60 * 1000);
        
        const hours = Math.max(0, netMs / (1000 * 60 * 60));
        const earned = hours * rate;

        sheet.getRange(rowIndex, 4).setValue(hours.toFixed(2)); // Col D
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
  const data = userSheet.getLastRow() < 7 ? [] : userSheet.getRange(7, 1, userSheet.getLastRow() - 6, 5).getValues();
  let html = "<table style='width:100%; font-size: 14px; border-collapse: collapse;'>";
  data.slice(-5).reverse().forEach(row => {
    html += `<tr style='border-bottom: 1px solid #eee;'><td style='padding:5px;'>${row[0] ? new Date(row[0]).toLocaleDateString() : ''}</td><td style='text-align:right; padding:5px;'>${parseFloat(row[3]||0).toFixed(2)} hrs</td></tr>`;
  });
  
  // Get active tasks for this employee
  const tasks = getTasks(pin);
  
  return { 
    name: userSheet.getRange("A3").getValue(), 
    currentStatus: status, 
    pay: isAdmin ? "Administrator Access" : `Current Rate: $${rate}/hr`, 
    history: html + "</table>", 
    isAdmin: isAdmin,
    tasks: tasks.tasks || []
  };
}

function getAdminLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let html = "";

  // 1. PULL RECENT ALERTS (Last 3 Stock & Maint) - Left Aligned
  html += "<div style='text-align:left;'><b>Recent Alerts:</b>";
  html += "<div style='background:#fffcf0; padding:10px; border-radius:8px; margin:10px 0; border:1px solid #ffd54f; font-size:12px;'>";
  
  // Stock Alerts - Skip Header Row 1 & 2
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

  // Maintenance Alerts - Skip Header Row 1 & 2
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
          in: row[5].toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}),
          out: row[8] instanceof Date ? row[8].toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) : "Working..."
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
    
    // Get tasks for pre-filling
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
    
    // Clear completed tasks for this employee
    clearCompletedTasks(pin);
  }
}

function sendWeeklySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  let emailBody = "<html><body style='font-family: Arial, sans-serif;'>";
  emailBody += "<h2 style='color: #2b3d2b;'>Weekly Payroll Summary - Sparrow Farms</h2>";
  emailBody += "<p>Here's the summary of hours worked this week:</p>";
  
  let totalOwed = 0;
  let hasData = false;
  
  const today = new Date();
  const oneWeekAgo = new Date();
  oneWeekAgo.setDate(today.getDate() - 7);
  
  sheets.forEach(sheet => {
    const pin = sheet.getRange("D3").getValue();
    if (!pin || sheet.getName().includes("Log") || sheet.getName().includes("Tasks")) return;
    
    const name = sheet.getRange("A3").getValue();
    const email = sheet.getRange("C3").getValue();
    const rate = parseFloat(sheet.getRange("E3").getValue()) || 0;
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 7) return;
    
    const data = sheet.getRange(7, 1, lastRow - 6, 9).getValues();
    
    let weekHours = 0;
    let weekEarned = 0;
    let entries = [];
    
    data.forEach(row => {
      const dateVal = new Date(row[0]);
      const hours = parseFloat(row[3]) || 0;
      const earned = parseFloat(row[4]) || 0;
      
      if (dateVal >= oneWeekAgo && dateVal <= today && hours > 0) {
        weekHours += hours;
        weekEarned += earned;
        entries.push({
          date: dateVal.toLocaleDateString(),
          hours: hours.toFixed(2),
          earned: earned.toFixed(2),
          work: row[1] || "No description"
        });
      }
    });
    
    if (weekHours > 0) {
      hasData = true;
      totalOwed += weekEarned;
      
      emailBody += "<div style='background: #f9fbf9; padding: 15px; margin: 15px 0; border-left: 4px solid #388e3c; border-radius: 5px;'>";
      emailBody += `<h3 style='margin-top: 0; color: #2b3d2b;'>${name}</h3>`;
      emailBody += `<p><strong>Email:</strong> ${email}<br>`;
      emailBody += `<strong>Rate:</strong> $${rate.toFixed(2)}/hr<br>`;
      emailBody += `<strong>Total Hours:</strong> ${weekHours.toFixed(2)}<br>`;
      emailBody += `<strong>Amount Owed:</strong> <span style='color: #388e3c; font-size: 1.2em;'>$${weekEarned.toFixed(2)}</span></p>`;
      
      emailBody += "<table style='width: 100%; border-collapse: collapse; font-size: 0.9em;'>";
      emailBody += "<tr style='background: #e8f5e9;'><th style='padding: 8px; text-align: left;'>Date</th><th style='padding: 8px; text-align: right;'>Hours</th><th style='padding: 8px; text-align: right;'>Earned</th><th style='padding: 8px; text-align: left;'>Work Done</th></tr>";
      
      entries.forEach(entry => {
        emailBody += `<tr style='border-bottom: 1px solid #ddd;'>`;
        emailBody += `<td style='padding: 8px;'>${entry.date}</td>`;
        emailBody += `<td style='padding: 8px; text-align: right;'>${entry.hours}</td>`;
        emailBody += `<td style='padding: 8px; text-align: right;'>$${entry.earned}</td>`;
        emailBody += `<td style='padding: 8px; font-style: italic; color: #666;'>${entry.work}</td>`;
        emailBody += `</tr>`;
      });
      
      emailBody += "</table></div>";
    }
  });
  
  if (!hasData) {
    emailBody += "<p><em>No hours logged this week.</em></p>";
  } else {
    emailBody += `<div style='background: #2b3d2b; color: white; padding: 20px; margin: 20px 0; border-radius: 5px; text-align: center;'>`;
    emailBody += `<h3 style='margin: 0;'>Total Payroll This Week</h3>`;
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
    
    if (!isAdmin) { // Don't include admins in task assignment list
      staff.push({ name: name, pin: String(pin) });
    }
  });
  
  return { staff: staff };
}

function assignTasks(targetPin, tasksText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tasksSheet = ss.getSheetByName("Tasks");
  
  // Create Tasks sheet if it doesn't exist
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
      activeTasks.push(row[1]); // Task text
    }
  });
  
  return { tasks: activeTasks };
}

function clearCompletedTasks(pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = ss.getSheetByName("Tasks");
  
  if (!tasksSheet || tasksSheet.getLastRow() < 2) return;
  
  const lastRow = tasksSheet.getLastRow();
  // Work backwards to avoid row index shifting
  for (let i = lastRow; i >= 2; i--) {
    const taskPin = String(tasksSheet.getRange(i, 1).getValue()).trim();
    const status = tasksSheet.getRange(i, 4).getValue();
    
    if (taskPin === String(pin).trim() && status === "Active") {
      tasksSheet.getRange(i, 4).setValue("Completed");
      tasksSheet.getRange(i, 5).setValue(new Date()); // Add completion date in column E
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
      rowIndex: index + 2, // Actual row in sheet (accounting for header)
      pin: String(row[0]),
      name: name,
      task: row[1],
      dateAssigned: row[2] instanceof Date ? row[2].toLocaleDateString() : row[2],
      status: row[3],
      dateCompleted: row[4] instanceof Date ? row[4].toLocaleDateString() : (row[4] || "")
    });
  });
  
  // Sort: Active tasks first, then by date
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
  
  // Get the original task data
  const row = parseInt(rowIndex);
  const originalTask = tasksSheet.getRange(row, 2).getValue(); // Task text
  const originalDateAssigned = tasksSheet.getRange(row, 3).getValue(); // Original date
  
  // Mark the original as "Reassigned"
  tasksSheet.getRange(row, 4).setValue("Reassigned");
  tasksSheet.getRange(row, 5).setValue(new Date()); // Reassignment date in column E
  
  // Create new active task for the new person
  tasksSheet.appendRow([
    String(newPin),           // New staff PIN
    originalTask,             // Same task text
    new Date(),               // New assignment date (today)
    "Active",                 // Status
    ""                        // No completion date yet
  ]);
  
  return { success: true, msg: "Task reassigned successfully" };
}

function addStaffMember(name, phone, email, pin, rate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Validate inputs
  if (!name || !pin || !rate) {
    return { error: "Name, PIN, and Rate are required" };
  }
  
  // Check if PIN already exists
  const existingSheet = findSheetByPin(String(pin).trim());
  if (existingSheet) {
    return { error: "PIN already exists. Please choose a different PIN." };
  }
  
  // Find the TEMPLATE sheet
  const templateSheet = ss.getSheetByName("TEMPLATE");
  if (!templateSheet) {
    return { error: "TEMPLATE sheet not found. Please create a template sheet first." };
  }
  
  // Duplicate the template
  const newSheet = templateSheet.copyTo(ss);
  newSheet.setName(name);
  
  // Populate Row 3 with the staff info (Row 1 is heading, Row 2 is column labels, Row 3 is data)
  newSheet.getRange("A3").setValue(name);           // Employee Name
  newSheet.getRange("B3").setValue(phone || "");     // Phone Number
  newSheet.getRange("C3").setValue(email || "");     // Email Address
  newSheet.getRange("D3").setValue(String(pin));     // PIN
  newSheet.getRange("E3").setValue(parseFloat(rate)); // Current Hourly Rate
  newSheet.getRange("F3").setValue(new Date());      // Last Increase Date (today)
  newSheet.getRange("G3").setValue("");              // Previous Rate (empty for new hire)
  
  // Move the new sheet to the first position
  newSheet.activate();
  ss.moveActiveSheet(1);
  
  return { success: true, msg: "Staff member added: " + name };
}