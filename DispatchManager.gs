/**
 * @fileoverview Daily Shift Dispatch, Training Routes, and End-of-Shift Engine.
 */

function processDispatchCheckout(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deploySheet = ss.getSheetByName('Deployments');
    const vehicleSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    const deployId = 'DEP-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const dispatchDate = data.dispatchDate ? new Date(data.dispatchDate) : now;

    let fullGoals = '';
    if (data.focusCategory) {
      fullGoals = '[' + data.focusCategory + ']' + (data.goals ? ' ' + data.goals : '');
    } else {
      fullGoals = data.goals || '';
    }

    // Append Deployments (16 Cols)
    deploySheet.appendRow([
      deployId, dispatchDate, data.shift, data.cleanName, data.rentalId,
      data.assignedTech, data.workZone, fullGoals, data.batteryTarget || '',
      'Dispatched', now, '', user, now, 
      data.coDriver || '', data.isTraining ? 'Yes' : 'No'
    ]);

    // Update Vehicle to 'Dispatched'
    const vData = vehicleSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        const row = i + 1;
        vehicleSheet.getRange(row, 9).setValue('Dispatched');       // Status
        vehicleSheet.getRange(row, 12).setValue(data.workZone);     // Zone
        const fullTech = data.coDriver ? `${data.assignedTech} & ${data.coDriver}` : data.assignedTech;
        vehicleSheet.getRange(row, 13).setValue(fullTech);          // Assigned Tech
        vehicleSheet.getRange(row, 16).setValue(user);
        vehicleSheet.getRange(row, 17).setValue(now);
        break;
      }
    }

    logAudit('DISPATCH_CHECKOUT', 'Deployments', deployId, `Dispatched: ${data.assignedTech}`);
    return { success: true, message: `Vehicle dispatched to ${data.assignedTech}.` };
  } catch (err) {
    throw new Error(err.message);
  }
}

/**
 * Checks a van back into the Warehouse, clearing its driver and logging the return time.
 */
function endShift(deploymentId, rentalId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dSheet = ss.getSheetByName('Deployments');
  const vSheet = ss.getSheetByName('Vehicles');
  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  // 1. Mark Deployment Complete
  const dRows = dSheet.getDataRange().getValues();
  for (let i = 1; i < dRows.length; i++) {
    if (String(dRows[i][0]).trim() === String(deploymentId).trim()) {
      dSheet.getRange(i + 1, 10).setValue('Completed'); // Status
      dSheet.getRange(i + 1, 12).setValue(now);         // Return Time
      break;
    }
  }

  // 2. Return Vehicle to Warehouse (Ready)
  const vRows = vSheet.getDataRange().getValues();
  for (let j = 1; j < vRows.length; j++) {
    if (String(vRows[j][1]).trim() === String(rentalId).trim()) {
      const vRow = j + 1;
      vSheet.getRange(vRow, 9).setValue('Warehouse (Ready)');
      vSheet.getRange(vRow, 12).clearContent(); // Clear Zone
      vSheet.getRange(vRow, 13).clearContent(); // Clear Tech
      vSheet.getRange(vRow, 16).setValue(user);
      vSheet.getRange(vRow, 17).setValue(now);
      break;
    }
  }

  logAudit('SHIFT_ENDED', 'Deployments', deploymentId, 'Shift ended, van returned to Warehouse');
  return { success: true, message: 'Shift ended successfully. Van is back in the Warehouse.' };
}