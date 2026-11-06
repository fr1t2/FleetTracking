function processDispatchCheckout(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dSheet = ss.getSheetByName('Deployments');
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    const deployId = 'DEP-' + Utilities.getUuid().substring(0, 8).toUpperCase();

    // Deployments Row: 21 columns
    dSheet.appendRow([
      deployId, now, data.shift, data.cleanName, data.rentalId,
      data.assignedTech, data.workZone, data.goals || '', data.batteryTarget || '',
      'Dispatched', now, '', user, now, data.coDriver || '', data.isTraining ? 'Yes' : 'No',
      data.eKey || '', '', '', data.dispatchNotes || '', ''
    ]);

    const vData = vSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        const row = i + 1;
        vSheet.getRange(row, 10).setValue('Dispatched'); // Col J: Status
        vSheet.getRange(row, 13).setValue(data.workZone); // Col M: Zone
        vSheet.getRange(row, 14).setValue(data.assignedTech); // Col N: Tech
        vSheet.getRange(row, 17).setValue(user); // Col Q: Updated By
        vSheet.getRange(row, 18).setValue(now); // Col R: Updated At
        break;
      }
    }
    return { success: true };
  } catch (err) {
    throw new Error('Dispatch failed: ' + err.message);
  }
}

function endShift(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dSheet = ss.getSheetByName('Deployments');
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    const dRows = dSheet.getDataRange().getValues();
    for (let i = 1; i < dRows.length; i++) {
      if (String(dRows[i][0]).trim() === String(payload.deploymentId).trim()) {
        dSheet.getRange(i + 1, 10).setValue('Completed'); // Col J: Status
        dSheet.getRange(i + 1, 12).setValue(now); // Col L: Return Time
        break;
      }
    }

    const vRows = vSheet.getDataRange().getValues();
    for (let j = 1; j < vRows.length; j++) {
      if (String(vRows[j][1]).trim().toLowerCase() === String(payload.rentalId).trim().toLowerCase()) {
        const row = j + 1;
        vSheet.getRange(row, 10).setValue('Warehouse (Ready)'); // Col J
        vSheet.getRange(row, 13).clearContent(); // Clear Zone (Col M)
        vSheet.getRange(row, 14).clearContent(); // Clear Tech (Col N)
        vSheet.getRange(row, 17).setValue(user); // Col Q
        vSheet.getRange(row, 18).setValue(now); // Col R
        break;
      }
    }
    return { success: true };
  } catch (err) {
    throw new Error('End shift failed: ' + err.message);
  }
}