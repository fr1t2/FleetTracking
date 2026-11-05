function processDispatchCheckout(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const deploySheet = ss.getSheetByName('Deployments');
  const vehicleSheet = ss.getSheetByName('Vehicles');
  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';
  const deployId = 'DEP-' + Utilities.getUuid().substring(0, 8).toUpperCase();

  let fullGoals = data.focusCategory ? `[${data.focusCategory}] ${data.goals || ''}` : (data.goals || '');

  // LOG NEW DEPLOYMENT (Make sure this matches your column order, adding E-Key at the end)
  deploySheet.appendRow([
    deployId, now, data.shift, data.cleanName, data.rentalId,
    data.assignedTech, data.workZone, fullGoals, data.batteryTarget || '',
    'Dispatched', now, '', user, now, data.coDriver || '', data.isTraining ? 'Yes' : 'No',
    data.eKey || '', '', '' // E-Key Number, Batteries In (blank), E-Key Returned (blank)
  ]);

  // UPDATE TARGET VEHICLE
  const vData = vehicleSheet.getDataRange().getValues();
  for (let i = 1; i < vData.length; i++) {
    if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
      const row = i + 1;
      vehicleSheet.getRange(row, 9).setValue('Dispatched');
      vehicleSheet.getRange(row, 12).setValue(data.workZone);
      const fullTech = data.coDriver ? `${data.assignedTech} & ${data.coDriver}` : data.assignedTech;
      vehicleSheet.getRange(row, 13).setValue(fullTech);
      vehicleSheet.getRange(row, 16).setValue(user);
      vehicleSheet.getRange(row, 17).setValue(now);
      break;
    }
  }
  return { success: true };
}

function updateDeploymentRecord(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dSheet = ss.getSheetByName('Deployments');
  const vSheet = ss.getSheetByName('Vehicles');
  
  const dRows = dSheet.getDataRange().getValues();
  let dRow = -1;
  let oldRentalId = '';

  for (let i = 1; i < dRows.length; i++) {
    if (String(dRows[i][0]).trim() === String(data.id).trim()) {
      dRow = i + 1;
      oldRentalId = String(dRows[i][4]).trim();
      break;
    }
  }
  if (dRow === -1) throw new Error('Deployment not found.');

  // Update Deployment Sheet
  if (data.shift) dSheet.getRange(dRow, 3).setValue(data.shift);
  if (data.vehicle) {
    dSheet.getRange(dRow, 4).setValue(data.vehicle);
    dSheet.getRange(dRow, 5).setValue(data.rentalId);
  }
  if (data.zone) dSheet.getRange(dRow, 7).setValue(data.zone);
  if (data.goals !== undefined) dSheet.getRange(dRow, 8).setValue(data.goals);
  if (data.target !== undefined) dSheet.getRange(dRow, 9).setValue(data.target);
  if (data.eKey !== undefined) dSheet.getRange(dRow, 17).setValue(data.eKey); // Column Q

  // If vehicle changed, release old van and assign new van
  if (data.rentalId && oldRentalId !== data.rentalId) {
    const vRows = vSheet.getDataRange().getValues();
    // Free old van
    for (let j = 1; j < vRows.length; j++) {
      if (String(vRows[j][1]).trim() === oldRentalId) {
        vSheet.getRange(j + 1, 9).setValue('Warehouse');
        vSheet.getRange(j + 1, 12).clearContent();
        vSheet.getRange(j + 1, 13).clearContent();
        break;
      }
    }
    // Assign new van
    for (let j = 1; j < vRows.length; j++) {
      if (String(vRows[j][1]).trim() === String(data.rentalId).trim()) {
        vSheet.getRange(j + 1, 9).setValue('Dispatched');
        vSheet.getRange(j + 1, 12).setValue(data.zone);
        vSheet.getRange(j + 1, 13).setValue(data.tech);
        break;
      }
    }
  }
  return { success: true, message: 'Dispatch updated successfully.' };
}

function endShift(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dSheet = ss.getSheetByName('Deployments');
  const vSheet = ss.getSheetByName('Vehicles');
  const now = new Date();

  // 1. Update Deployments
  const dRows = dSheet.getDataRange().getValues();
  for (let i = 1; i < dRows.length; i++) {
    if (String(dRows[i][0]).trim() === String(payload.deploymentId).trim()) {
      dSheet.getRange(i + 1, 10).setValue('Completed'); 
      dSheet.getRange(i + 1, 12).setValue(now);
      dSheet.getRange(i + 1, 18).setValue(payload.batteriesIn || 0);     // Column R
      dSheet.getRange(i + 1, 19).setValue(payload.eKeyReturned ? 'Yes' : 'No'); // Column S
      break;
    }
  }

  // 2. Return Vehicle to Warehouse
  const vRows = vSheet.getDataRange().getValues();
  for (let j = 1; j < vRows.length; j++) {
    if (String(vRows[j][1]).trim() === String(payload.rentalId).trim()) {
      vSheet.getRange(j + 1, 9).setValue('Warehouse');
      vSheet.getRange(j + 1, 12).clearContent();
      vSheet.getRange(j + 1, 13).clearContent();
      break;
    }
  }
  return { success: true };
}

/**
 * Processes a drag-and-drop End Shift / Check-in.
 */
function processCheckIn(rentalId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vSheet = ss.getSheetByName('Vehicles');
    const dSheet = ss.getSheetByName('Deployments');
    const now = new Date();

    // 1. Clear the driver and zone, set back to Warehouse
    const vData = vSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(rentalId).trim().toLowerCase()) {
        vSheet.getRange(i + 1, 10).setValue('Warehouse (Ready)');
        vSheet.getRange(i + 1, 13).clearContent();
        vSheet.getRange(i + 1, 14).clearContent();
        found = true;
        break;
      }
    }

    if (!found) throw new Error("Vehicle not found.");

    // 2. Mark the active deployment as Completed
    const dData = dSheet.getDataRange().getValues();
    for (let i = 1; i < dData.length; i++) {
      if (String(dData[i][4]).trim().toLowerCase() === String(rentalId).trim().toLowerCase() && String(dData[i][9]) === 'Dispatched') {
        dSheet.getRange(i + 1, 10).setValue('Completed');
        dSheet.getRange(i + 1, 12).setValue(now); // Return Time
        break;
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}