function processCheckInVehicle(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    const vData = vSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        const row = i + 1;
        vSheet.getRange(row, 10).setValue('Warehouse (Ready)'); // Col J
        vSheet.getRange(row, 13).clearContent(); // Clear Zone
        vSheet.getRange(row, 14).clearContent(); // Clear Tech
        vSheet.getRange(row, 16).clearContent(); // Clear Notes (Col P)
        vSheet.getRange(row, 17).setValue(user);
        vSheet.getRange(row, 18).setValue(now);
        break;
      }
    }

    // Auto-resolve any missing incident resolution logs
    if (data.incidentId) {
      const iSheet = ss.getSheetByName('Incidents_Service');
      const iData = iSheet.getDataRange().getValues();
      for (let j = 1; j < iData.length; j++) {
        if (String(iData[j][0]).trim() === data.incidentId) {
          iSheet.getRange(j + 1, 9).setValue(data.resolutionNotes || 'Checked in via dashboard');
          iSheet.getRange(j + 1, 10).setValue(now);
          iSheet.getRange(j + 1, 11).setValue(user);
          break;
        }
      }
    }
    return { success: true };
  } catch (err) {
    throw new Error('Check-in failed: ' + err.message);
  }
}

function logNewServiceNeed(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const iSheet = ss.getSheetByName('Incidents_Service');
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    const incidentId = 'INC-' + Utilities.getUuid().substring(0, 6).toUpperCase();

    // Incidents Row: 14 Columns
    iSheet.appendRow([
      incidentId, now, data.cleanName, data.rentalId,
      data.category, data.description, data.isOos ? 'YES' : 'NO',
      '', '', '', '', user, now, data.isOos ? 'OOS' : 'Minor'
    ]);

    if (data.isOos) {
      const vData = vSheet.getDataRange().getValues();
      for (let i = 1; i < vData.length; i++) {
        if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
          const row = i + 1;
          vSheet.getRange(row, 10).setValue('Out of Service');
          vSheet.getRange(row, 13).clearContent();
          vSheet.getRange(row, 14).clearContent();
          vSheet.getRange(row, 16).setValue(data.description); // Set Notes
          vSheet.getRange(row, 17).setValue(user);
          vSheet.getRange(row, 18).setValue(now);
          break;
        }
      }
    }
    return { success: true };
  } catch (err) {
    throw new Error('Log service failed: ' + err.message);
  }
}

function closeServiceIssue(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const iSheet = ss.getSheetByName('Incidents_Service');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    const iData = iSheet.getDataRange().getValues();
    for (let i = 1; i < iData.length; i++) {
      if (String(iData[i][0]).trim() === data.incidentId) {
        iSheet.getRange(i + 1, 9).setValue(data.resolutionNotes);
        iSheet.getRange(i + 1, 10).setValue(now);
        iSheet.getRange(i + 1, 11).setValue(user);
        break;
      }
    }
    return { success: true };
  } catch (err) {
    throw new Error('Close ticket failed: ' + err.message);
  }
}