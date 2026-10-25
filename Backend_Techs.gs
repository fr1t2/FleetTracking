/**
 * Run this function ONCE from the editor to upgrade your database.
 * It creates the Tech_Profiles sheet and migrates your existing roster.
 */
function upgradeDatabaseForTechs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let techSheet = ss.getSheetByName('Tech_Profiles');
  
  if (!techSheet) {
    techSheet = ss.insertSheet('Tech_Profiles');
    techSheet.appendRow(['Tech Name', 'Phone', 'Shift', 'Start Time', 'End Time', 'Work Days', 'Passed Test', 'Go-Fast', 'Restricted', 'Status', 'Notes']);
    techSheet.getRange('A1:K1').setFontWeight('bold').setBackground('#f0f2f5');
    
    // Migrate existing techs from _Config if they exist
    const configSheet = ss.getSheetByName('_Config');
    if (configSheet) {
      const roster = configSheet.getRange('K2:K' + configSheet.getLastRow()).getValues().flat().filter(String);
      roster.forEach(name => {
        techSheet.appendRow([name, '', 'First Shift', '07:00', '15:30', 'Mon-Fri', 'Yes', 'No', 'No', 'Available', 'Migrated']);
      });
    }
  }
}

/**
 * Updates a Tech's Profile attributes.
 */
function updateTechProfile(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Tech_Profiles');
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === String(data.origName).trim().toLowerCase()) {
      rowIndex = i + 1; break;
    }
  }

  const rowData = [
    data.name, data.phone, data.shift, data.startTime, data.endTime, 
    data.workDays, data.passedTest ? 'Yes' : 'No', data.goFast ? 'Yes' : 'No', 
    data.restricted ? 'Yes' : 'No', data.status, data.notes || ''
  ];

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 11).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  
  logAudit('TECH_UPDATED', 'Tech_Profiles', data.name, `Status: ${data.status} | Restricted: ${data.restricted}`);
  return { success: true };
}

/**
 * Permanently removes a tech from the system
 */
function processDeleteTech(techName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Remove from Tech Profiles
  const tSheet = ss.getSheetByName('Tech_Profiles');
  if (tSheet) {
    const tData = tSheet.getDataRange().getValues();
    for (let i = 1; i < tData.length; i++) {
      if (String(tData[i][0]).trim().toLowerCase() === techName.toLowerCase()) {
        tSheet.deleteRow(i + 1);
        break;
      }
    }
  }

  // 2. Remove from Config Roster
  const cSheet = ss.getSheetByName('_Config');
  if (cSheet) {
    const cData = cSheet.getDataRange().getValues();
    for (let i = 1; i < cData.length; i++) {
      if (String(cData[i][9]).trim().toLowerCase() === techName.toLowerCase()) { // Col J is Tech Roster
        cSheet.getRange(i + 1, 10).clearContent(); 
        break;
      }
    }
  }
  
  if (typeof logAudit === 'function') logAudit('TECH_DELETED', 'Tech_Profiles', techName, 'Terminated / Removed from service');
  return { success: true };
}

/**
 * Logs an attendance event (Sick, NCNS, PTO) and updates their profile status
 */
function logTechAttendance(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';
  
  // 1. Create Attendance_Log tab if it doesn't exist
  let aSheet = ss.getSheetByName('Attendance_Log');
  if (!aSheet) {
    aSheet = ss.insertSheet('Attendance_Log');
    aSheet.appendRow(['Date', 'Tech Name', 'Event Type', 'Notes', 'Logged By', 'Timestamp']);
    aSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#f0f2f5');
  }

  // 2. Record the Event
  aSheet.appendRow([ data.date, data.techName, data.type, data.notes, user, now ]);

  // 3. Auto-update their Tech Profile status
  const tSheet = ss.getSheetByName('Tech_Profiles');
  if (tSheet) {
    const tData = tSheet.getDataRange().getValues();
    for (let i = 1; i < tData.length; i++) {
      if (String(tData[i][0]).trim().toLowerCase() === data.techName.trim().toLowerCase()) {
        tSheet.getRange(i + 1, 10).setValue(data.type); // Col J is Status
        break;
      }
    }
  }

  return { success: true };
}

/**
 * Clears an attendance exception and sets the tech back to Available
 */
function clearTechException(techName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tSheet = ss.getSheetByName('Tech_Profiles');
    
    if (tSheet) {
      const tData = tSheet.getDataRange().getValues();
      for (let i = 1; i < tData.length; i++) {
        if (String(tData[i][0]).trim().toLowerCase() === techName.trim().toLowerCase()) {
          tSheet.getRange(i + 1, 10).setValue('Available'); // Col J is Status
          break;
        }
      }
    }
    
    if (typeof logAudit === 'function') logAudit('TECH_RETURNED_TO_DUTY', 'Tech_Profiles', techName, 'Exception cleared');
    return { success: true };
  } catch (err) {
    return { success: false, message: 'Failed to clear exception: ' + err.toString() };
  }
}