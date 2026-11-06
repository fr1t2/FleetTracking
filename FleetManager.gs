/**
 * @fileoverview Fleet Maintenance and Check-In Engine
 */

function processCheckInVehicle(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    if (!data.rentalId && !data.cleanName) {
      throw new Error('No vehicle identifier provided.');
    }

    const vData = vSheet.getDataRange().getValues();
    let found = false;
    
    // Vehicles: Status (J/10), Zone (M/13), Driver (N/14), Notes (P/16), SafeToDrive (false)
    for (let i = 1; i < vData.length; i++) {
      const rowRentalId = String(vData[i][1]).trim().toLowerCase();
      const rowCleanName = String(vData[i][0]).trim().toLowerCase();
      const targetRentalId = String(data.rentalId || '').trim().toLowerCase();
      const targetCleanName = String(data.cleanName || '').trim().toLowerCase();
      
      if ((targetRentalId && rowRentalId === targetRentalId) || (targetCleanName && rowCleanName === targetCleanName)) {
        vSheet.getRange(i + 1, 10).setValue('Warehouse (Ready)');
        vSheet.getRange(i + 1, 13).clearContent(); 
        vSheet.getRange(i + 1, 14).clearContent(); 
        vSheet.getRange(i + 1, 16).clearContent(); 
        vSheet.getRange(i + 1, 17).setValue(user); // Last Updated By
        vSheet.getRange(i + 1, 18).setValue(now);  // Last Updated At
        found = true;
        break;
      }
    }

    if (!found) throw new Error(`Vehicle ${data.cleanName || data.rentalId} not found.`);

    // Auto-Resolve open incidents
    let targetIncId = data.incidentId;
    const iSheet = ss.getSheetByName('Incidents_Service');
    
    if (iSheet) {
      const iData = iSheet.getDataRange().getValues();
      if (!targetIncId) {
         for (let j = iData.length - 1; j >= 1; j--) {
           if (String(iData[j][3]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase() && !iData[j][9]) {
             targetIncId = String(iData[j][0]).trim();
             break;
           }
         }
      }
      
      if (targetIncId) {
        for (let j = 1; j < iData.length; j++) {
          if (String(iData[j][0]).trim() === targetIncId) {
            iSheet.getRange(j + 1, 9).setValue(data.resolutionNotes || 'Checked in via dashboard');
            iSheet.getRange(j + 1, 10).setValue(now); 
            iSheet.getRange(j + 1, 11).setValue(user);
            break;
          }
        }
      }
    }
    
    return { success: true, message: 'Van is back in the Warehouse.' };
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

    // 14 Columns: [ID, Date, Name, Rental ID, Cat, Desc, OOS(YES/NO), Impact, ResNotes, DateRes, ResBy, LoggedBy, Timestamp, Severity]
    iSheet.appendRow([
      incidentId, now, data.cleanName, data.rentalId,
      data.category, data.description, data.isOos ? 'YES' : 'NO',
      data.statusImpact, '', '', '', user, now, data.severity || (data.isOos ? 'OOS' : 'Minor')
    ]);

    if (data.isOos) {
      const vData = vSheet.getDataRange().getValues();
      for (let i = 1; i < vData.length; i++) {
        if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
          const row = i + 1;
          vSheet.getRange(row, 10).setValue('Out of Service');
          vSheet.getRange(row, 13).clearContent();
          vSheet.getRange(row, 14).clearContent();
          vSheet.getRange(row, 16).setValue(data.description); 
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