function processDispatchCheckout(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deploySheet = ss.getSheetByName('Deployments');
    const vehicleSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    const deployId = 'DEP-' + Utilities.getUuid().substring(0, 6).toUpperCase();

    // 1. AUTO-RESOLVE CONFLICTS (Enforce single-van/single-driver rules)
    const dData = deploySheet.getDataRange().getValues();
    const vData = vehicleSheet.getDataRange().getValues();
    
    for (let j = 1; j < dData.length; j++) {
      if (String(dData[j][9]).trim() === 'Dispatched') {
        const depRental = String(dData[j][4]).trim().toLowerCase();
        const lead = String(dData[j][5]).trim().toLowerCase();
        
        const targetRental = String(data.rentalId).trim().toLowerCase();
        const targetLead = String(data.assignedTech).trim().toLowerCase();

        // If the primary tech is already out in a DIFFERENT van, resolve it
        if (lead === targetLead && depRental !== targetRental) {
          deploySheet.getRange(j + 1, 10).setValue('Completed');
          deploySheet.getRange(j + 1, 12).setValue(now);
          
          for (let k = 1; k < vData.length; k++) {
            if (String(vData[k][1]).trim().toLowerCase() === depRental) {
              vehicleSheet.getRange(k + 1, 10).setValue('Warehouse (Ready)');
              vehicleSheet.getRange(k + 1, 13).clearContent(); 
              vehicleSheet.getRange(k + 1, 14).clearContent(); 
              break;
            }
          }
        }
      }
    }

    // 2. BUILD THE PAYLOAD ARRAY (29 Columns wide to reach AC)
    let fullGoals = data.goals || '';

    const rowData = [];
    rowData[0] = deployId;                      // Col A (0): ID
    rowData[1] = now;                           // Col B (1): Date
    rowData[2] = data.shift || 'First Shift';   // Col C (2): Shift
    rowData[3] = data.cleanName;                // Col D (3): Vehicle Name
    rowData[4] = data.rentalId;                 // Col E (4): Rental ID
    rowData[5] = data.assignedTech;             // Col F (5): Primary Tech
    rowData[6] = data.workZone;                 // Col G (6): Zone
    rowData[7] = fullGoals;                     // Col H (7): Goals 
    rowData[8] = data.batteryTarget || '';      // Col I (8): Battery Target
    rowData[9] = 'Dispatched';                  // Col J (9): Status
    rowData[10] = now;                          // Col K (10): Time Out
    rowData[11] = '';                           // Col L (11): Time In
    rowData[12] = user;                         // Col M (12): Logged By
    rowData[13] = now;                          // Col N (13): Timestamp
    rowData[14] = data.coDriver || '';          // Col O (14): Co-Driver
    rowData[15] = data.isTraining ? 'Yes' : 'No';// Col P (15): Is Training
    rowData[16] = data.eKey || '';              // Col Q (16): E-Key Number
    rowData[17] = '';                           // Col R (17): (Unused)
    rowData[18] = '';                           // Col S (18): (Unused)
    rowData[19] = data.dispatchNotes || '';     // Col T (19): Dispatch Notes
    rowData[20] = '';                           // Col U (20): (Unused)
    rowData[21] = data.tgtDM || '';             // Col V (21): Damaged
    rowData[22] = data.tgtER || '';             // Col W (22): Error
    rowData[23] = data.tgtIB || '';             // Col X (23): IOT Battery
    rowData[24] = data.tgtID || '';             // Col Y (24): Idle
    rowData[25] = data.tgtLB || '';             // Col Z (25): Low Battery
    rowData[26] = data.tgtOF || '';             // Col AA (26): Offline
    rowData[27] = data.tgtRT || '';             // Col AB (27): Retrieval
    rowData[28] = data.tgtTK || '';             // Col AC (28): Ticket

    // Ensure array has exactly 29 elements to prevent shifted columns
    for(let i=0; i<29; i++) { if(rowData[i] === undefined) rowData[i] = ''; }
    deploySheet.appendRow(rowData);

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
        if (data.batteryTarget !== undefined) deploySheet.getRange(dRow, 9).setValue(data.batteryTarget);
        
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