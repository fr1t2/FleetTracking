/**
 * Processes the Drag-and-Drop quick dispatch from the unified dashboard.
 * Updates the vehicle status and logs the deployment.
 */
function processDispatchCheckout(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dSheet = ss.getSheetByName('Deployments');
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    const depId = 'DEP-' + Utilities.getUuid().substring(0, 6).toUpperCase();

    // 1. Log the new Deployment Ticket
    // Columns: Deployment ID | Date | Shift | Vehicle Name | Rental ID | Tech | Zone | Focus | Battery | Status | Depart | Return | User | Timestamp
    dSheet.appendRow([
      depId, now, 'First Shift', '—', data.rentalId, data.tech,
      data.zones, '', data.batteryCount, 'Dispatched',
      now, '', user, now
    ]);

    // 2. Update the Vehicle's Current Status in the Vehicles Tab
    // Based on the 19-column schema: Status=Col 10, Zone=Col 13, Tech=Col 14
    const vData = vSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        vSheet.getRange(i + 1, 10).setValue('Dispatched');     // Update Status
        vSheet.getRange(i + 1, 13).setValue(data.zones);       // Update Current Zone
        vSheet.getRange(i + 1, 14).setValue(data.tech);        // Update Assigned Tech
        found = true;
        break;
      }
    }

    if (!found) {
       return { success: false, message: `Could not find Unit ${data.rentalId} in the database.` };
    }

    return { success: true, message: 'Van Dispatched successfully.' };
    
  } catch (err) {
    return { success: false, message: 'Dispatch failed: ' + err.toString() };
  }
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