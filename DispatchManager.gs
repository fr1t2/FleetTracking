function processDispatchCheckout(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deploySheet = ss.getSheetByName('Deployments');
    const vehicleSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    const deployId = 'DEP-' + Utilities.getUuid().substring(0, 6).toUpperCase();

    // 1. AUTO-RESOLVE CONFLICTS (Enforce single-van/single-driver rules)
    // If the tech is already driving a DIFFERENT van, end that shift and release the old van.
    const dData = deploySheet.getDataRange().getValues();
    const vData = vehicleSheet.getDataRange().getValues();
    
    for (let j = 1; j < dData.length; j++) {
      if (String(dData[j][9]).trim() === 'Dispatched') {
        const depRental = String(dData[j][4]).trim().toLowerCase();
        const lead = String(dData[j][5]).trim().toLowerCase();
        
        const targetRental = String(data.rentalId).trim().toLowerCase();
        const targetLead = String(data.assignedTech).trim().toLowerCase();

        // If the tech is already out, but in a different van
        if (lead === targetLead && depRental !== targetRental) {
          // Complete the old conflicting shift
          deploySheet.getRange(j + 1, 10).setValue('Completed');
          deploySheet.getRange(j + 1, 12).setValue(now); // Return Time
          
          // Reset that old van back to Warehouse (Ready)
          for (let k = 1; k < vData.length; k++) {
            if (String(vData[k][1]).trim().toLowerCase() === depRental) {
              vehicleSheet.getRange(k + 1, 10).setValue('Warehouse (Ready)');
              vehicleSheet.getRange(k + 1, 13).clearContent(); // Clear Zone
              vehicleSheet.getRange(k + 1, 14).clearContent(); // Clear Driver
              break;
            }
          }
        }
      }
    }

    // 2. LOG NEW DEPLOYMENT
    let fullGoals = data.focusCategory ? `[${data.focusCategory}] ${data.goals || ''}` : (data.goals || '');

    deploySheet.appendRow([
      deployId, now, data.shift || 'First Shift', data.cleanName, data.rentalId,
      data.assignedTech, data.workZone, fullGoals, data.batteryTarget || '',
      'Dispatched', now, '', user, now, data.coDriver || '', data.isTraining ? 'Yes' : 'No'
    ]);

    // 3. UPDATE TARGET VEHICLE
    let foundTarget = false;
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        const row = i + 1;
        vehicleSheet.getRange(row, 10).setValue('Dispatched');
        vehicleSheet.getRange(row, 13).setValue(data.workZone);
        const fullTech = data.coDriver ? `${data.assignedTech} & ${data.coDriver}` : data.assignedTech;
        vehicleSheet.getRange(row, 14).setValue(fullTech);
        vehicleSheet.getRange(row, 17).setValue(user);
        vehicleSheet.getRange(row, 18).setValue(now);
        foundTarget = true;
        break;
      }
    }
    
    if (!foundTarget) {
      return { success: false, message: `Could not find Vehicle ${data.cleanName} in the sheet to update its status.` };
    }

    return { success: true, message: 'Van Dispatched successfully.' };
    
  } catch (err) {
    return { success: false, message: 'Dispatch failed: ' + err.toString() };
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

function updateDeploymentRecord(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deploySheet = ss.getSheetByName('Deployments');
    const vehicleSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    // 1. UPDATE DEPLOYMENT LOG
    const dData = deploySheet.getDataRange().getValues();
    let dRow = -1;
    for (let i = 1; i < dData.length; i++) {
      if (String(dData[i][0]).trim() === String(data.id).trim()) {
        dRow = i + 1;
        if (data.assignedTech !== undefined) deploySheet.getRange(dRow, 6).setValue(data.assignedTech);
        if (data.workZone !== undefined) deploySheet.getRange(dRow, 7).setValue(data.workZone);
        if (data.goals !== undefined) deploySheet.getRange(dRow, 8).setValue(data.goals);
        
        deploySheet.getRange(dRow, 13).setValue(user); // Logged By
        deploySheet.getRange(dRow, 14).setValue(now);  // Timestamp
        break;
      }
    }
    if (dRow === -1) throw new Error("Deployment record not found.");

    // 2. UPDATE ACTIVE VEHICLE
    const vData = vehicleSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        const row = i + 1;
        if (data.workZone !== undefined) vehicleSheet.getRange(row, 13).setValue(data.workZone);
        if (data.assignedTech !== undefined) vehicleSheet.getRange(row, 14).setValue(data.assignedTech);
        vehicleSheet.getRange(row, 17).setValue(user);
        vehicleSheet.getRange(row, 18).setValue(now);
        break;
      }
    }
    return { success: true };
  } catch (err) {
    throw new Error('Update failed: ' + err.toString());
  }
}