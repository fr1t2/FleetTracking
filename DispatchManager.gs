function processDispatchCheckout(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deploySheet = ss.getSheetByName('Deployments');
    const vehicleSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    const deployId = 'DEP-' + Utilities.getUuid().substring(0, 6).toUpperCase();

    // 1. AUTO-RESOLVE CONFLICTS
    const dData = deploySheet.getDataRange().getValues();
    const vData = vehicleSheet.getDataRange().getValues();
    
    for (let j = 1; j < dData.length; j++) {
      if (String(dData[j][9]).trim() === 'Dispatched') {
        const depRental = String(dData[j][4]).trim().toLowerCase();
        const lead = String(dData[j][5]).trim().toLowerCase();
        
        const targetRental = String(data.rentalId).trim().toLowerCase();
        const targetLead = String(data.assignedTech).trim().toLowerCase();

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

    let fullGoals = data.goals || '';

    const rowData = [];
    rowData[0] = deployId;                      
    rowData[1] = now;                           
    rowData[2] = data.shift || 'First Shift';   
    rowData[3] = data.cleanName;                
    rowData[4] = data.rentalId;                 
    rowData[5] = data.assignedTech;             
    rowData[6] = data.workZone;                 
    rowData[7] = fullGoals;                     
    rowData[8] = data.batteryTarget || '';      
    rowData[9] = 'Dispatched';                  
    rowData[10] = now;                          
    rowData[11] = '';                           
    rowData[12] = user;                         
    rowData[13] = now;                          
    rowData[14] = data.coDriver || '';          
    rowData[15] = data.isTraining ? 'Yes' : 'No';
    rowData[16] = data.eKey || '';              
    rowData[17] = '';                           
    rowData[18] = '';                           
    rowData[19] = data.dispatchNotes || '';     
    rowData[20] = '';                           
    rowData[21] = data.tgtDM || '';             
    rowData[22] = data.tgtER || '';             
    rowData[23] = data.tgtIB || '';             
    rowData[24] = data.tgtID || '';             
    rowData[25] = data.tgtLB || '';             
    rowData[26] = data.tgtOF || '';             
    rowData[27] = data.tgtRT || '';             
    rowData[28] = data.tgtTK || '';             

    for(let i=0; i<29; i++) { if(rowData[i] === undefined) rowData[i] = ''; }
    deploySheet.appendRow(rowData);

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
      return { success: false, message: `Could not find Vehicle ${data.cleanName} in the sheet.` };
    }

    logAudit('DISPATCH', 'Deployment', deployId, `Deployed ${data.rentalId} to Zone(s): ${data.workZone} | Tech: ${data.assignedTech} | Focus: ${fullGoals || 'Standard'}`);
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
        dSheet.getRange(i + 1, 10).setValue('Completed'); 
        dSheet.getRange(i + 1, 12).setValue(now); 
        
        if (payload.batteriesIn !== undefined) dSheet.getRange(i + 1, 18).setValue(payload.batteriesIn); 
        if (payload.eKeyReturned !== undefined) dSheet.getRange(i + 1, 19).setValue(payload.eKeyReturned); 
        if (payload.checkInNotes !== undefined) dSheet.getRange(i + 1, 21).setValue(payload.checkInNotes); 
        
        break;
      }
    }

    const vRows = vSheet.getDataRange().getValues();
    for (let j = 1; j < vRows.length; j++) {
      if (String(vRows[j][1]).trim().toLowerCase() === String(payload.rentalId).trim().toLowerCase()) {
        const row = j + 1;
        vSheet.getRange(row, 10).setValue('Warehouse (Ready)'); 
        vSheet.getRange(row, 13).clearContent(); 
        vSheet.getRange(row, 14).clearContent(); 
        vSheet.getRange(row, 17).setValue(user); 
        vSheet.getRange(row, 18).setValue(now); 
        break;
      }
    }
    
    logAudit('END_SHIFT', 'Deployment', payload.deploymentId, `Check-In for ${payload.rentalId} | BattIn: ${payload.batteriesIn || '0'} | Notes: ${payload.checkInNotes || 'None'}`);
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

    const dData = deploySheet.getDataRange().getValues();
    let dRow = -1;
    let oldData = null;
    
    for (let i = 1; i < dData.length; i++) {
      if (String(dData[i][0]).trim() === String(data.id).trim()) {
        dRow = i + 1;
        oldData = dData[i];
        break;
      }
    }
    if (dRow === -1) throw new Error("Deployment record not found.");

    // CALCULATE DELTA
    let changes = [];
    const checkChange = (newVal, oldVal, label) => {
      let cleanNew = newVal !== undefined && newVal !== null ? String(newVal).trim() : '';
      let cleanOld = oldVal !== undefined && oldVal !== null ? String(oldVal).trim() : '';
      if (cleanNew !== cleanOld) {
        changes.push(`${label}: '${cleanOld || 'empty'}' -> '${cleanNew || 'empty'}'`);
      }
    };
    checkChange(data.assignedTech, oldData[5], 'Driver');
    checkChange(data.workZone, oldData[6], 'Zone');
    checkChange(data.goals, oldData[7], 'Focus');
    checkChange(data.batteryTarget, oldData[8], 'Batt Target');

    // APPLY CHANGES
    if (data.assignedTech !== undefined) deploySheet.getRange(dRow, 6).setValue(data.assignedTech);
    if (data.workZone !== undefined) deploySheet.getRange(dRow, 7).setValue(data.workZone);
    if (data.goals !== undefined) deploySheet.getRange(dRow, 8).setValue(data.goals);
    if (data.batteryTarget !== undefined) deploySheet.getRange(dRow, 9).setValue(data.batteryTarget);
    deploySheet.getRange(dRow, 13).setValue(user); 
    deploySheet.getRange(dRow, 14).setValue(now);  

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
    
    const detailString = changes.length > 0 ? changes.join(' | ') : 'No fields changed';
    logAudit('EDIT_SHIFT', 'Deployment', data.id, detailString);
    
    return { success: true };
  } catch (err) {
    throw new Error('Update failed: ' + err.toString());
  }
}