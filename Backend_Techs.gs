/**
 * Saves and updates the Tech Profile including new HR and EV constraints
 */
function updateTechProfile(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Tech_Profiles');
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === p.origName.toLowerCase()) {
      rowIndex = i + 1; break;
    }
  }
  
  const rowData = [
    p.name, p.phone, p.shift, p.startTime, p.endTime, p.workDays, 
    p.passedTest ? 'Yes' : 'No', p.goFast ? 'Yes' : 'No', p.restricted ? 'Yes' : 'No', 
    p.status, '', p.hrStatus, p.coastFuel ? 'Yes' : 'No', p.evGo ? 'Yes' : 'No', p.trainer ? 'Yes' : 'No'
  ];

  if (rowIndex === -1) {
    sheet.appendRow(rowData);
  } else {
    sheet.getRange(rowIndex, 1, 1, 15).setValues([rowData]);
  }
  return {success: true};
}

/**
 * Termination simply moves them to "Terminated" status so they vanish from the frontend.
 */
function processDeleteTech(techName) {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const tSheet = ss.getSheetByName('Tech_Profiles');
   if (tSheet) {
     const tData = tSheet.getDataRange().getValues();
     for (let i = 1; i < tData.length; i++) {
       if (String(tData[i][0]).trim().toLowerCase() === techName.toLowerCase()) {
         tSheet.getRange(i + 1, 12).setValue('Terminated'); // Set HR Status
         break;
       }
     }
   }
   if (typeof logAudit === 'function') logAudit('TECH_TERMINATED', 'Tech_Profiles', techName, 'Terminated');
   return { success: true };
}

/**
 * Helper to allow frontend override of the driving test!
 */
function overrideDrivingTest(techName) {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const tSheet = ss.getSheetByName('Tech_Profiles');
   if (tSheet) {
     const tData = tSheet.getDataRange().getValues();
     for (let i = 1; i < tData.length; i++) {
       if (String(tData[i][0]).trim().toLowerCase() === techName.toLowerCase()) {
         tSheet.getRange(i + 1, 7).setValue('Yes');
         break;
       }
     }
   }
   return { success: true };
}

/**
 * Logs exception to Tech Profile, Attendance Log, AND Deployments (Shift History)!
 */
function logTechAttendance(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  // 1. Log to Attendance DB
  let aSheet = ss.getSheetByName('Attendance_Log');
  if (!aSheet) { aSheet = ss.insertSheet('Attendance_Log'); aSheet.appendRow(['Date', 'Tech Name', 'Event Type', 'Notes', 'Logged By', 'Timestamp']); }
  aSheet.appendRow([ data.date, data.techName, data.type, data.notes, user, now ]);

  // 2. Change their Profile Status
  const tSheet = ss.getSheetByName('Tech_Profiles');
  if (tSheet) {
    const tData = tSheet.getDataRange().getValues();
    for (let i = 1; i < tData.length; i++) {
      if (String(tData[i][0]).trim().toLowerCase() === data.techName.trim().toLowerCase()) {
        tSheet.getRange(i + 1, 10).setValue(data.type); break;
      }
    }
  }

  // 3. INJECT INTO DEPLOYMENTS FOR SHIFT HISTORY TAB
  const dSheet = ss.getSheetByName('Deployments');
  const depId = 'ATT-' + Utilities.getUuid().substring(0, 6).toUpperCase();
  dSheet.appendRow([depId, data.date, data.shift, '—', '—', data.techName, '—', data.notes, '—', data.type, now, '', user, now, '', 'No']);

  return { success: true };
}

/**
 * Clears exception and wipes their fake history record!
 */
function clearTechException(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Restore Profile Status to Available
  const tSheet = ss.getSheetByName('Tech_Profiles');
  const tData = tSheet.getDataRange().getValues();
  for (let i = 1; i < tData.length; i++) {
    if (String(tData[i][0]).trim().toLowerCase() === data.techName.trim().toLowerCase()) {
      tSheet.getRange(i + 1, 10).setValue('Available'); break;
    }
  }
  
  // 2. Erase or complete the fake Deployment shift
  const dSheet = ss.getSheetByName('Deployments');
  const dData = dSheet.getDataRange().getValues();
  for (let j = 1; j < dData.length; j++) {
    if (String(dData[j][5]).trim().toLowerCase() === data.techName.trim().toLowerCase() && 
       ['Sick', 'PTO', 'NCNS', 'Late'].includes(String(dData[j][9]).trim())) {
      dSheet.getRange(j + 1, 10).setValue(data.isLate ? 'Late (Arrived)' : 'Cleared'); // Strike it out or mark arrived
    }
  }

  return { success: true };
}