/**
 * @fileoverview Vehicle Registry & Service Operations
 */

function processAddVehicle(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Vehicles');
  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  sheet.appendRow([
    formData.vanName || '', formData.rentalId || '', formData.rentalAgency || 'Enterprise',
    formData.makeModel || 'Cargo Van', formData.vin || '', formData.licensePlate || '',
    formData.plateExpDate || '', formData.samsaraSerial || '', 'Warehouse (Ready)',
    now, '', '', '', now, '', user, now, formData.addons || ''
  ]);
  
  logAudit('VEHICLE_ADDED', 'Vehicles', formData.rentalId, 'Added to fleet');
  return { success: true, message: `Vehicle added!` };
}

function updateVehicleDetails(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Vehicles');
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === String(formData.rentalId).trim().toLowerCase()) {
      rowIndex = i + 1; break;
    }
  }
  if (rowIndex === -1) throw new Error(`Vehicle not found.`);

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  sheet.getRange(rowIndex, 1).setValue(formData.vanName);
  sheet.getRange(rowIndex, 6).setValue(formData.licensePlate);
  sheet.getRange(rowIndex, 7).setValue(formData.plateExpDate);
  sheet.getRange(rowIndex, 8).setValue(formData.samsaraSerial);
  sheet.getRange(rowIndex, 18).setValue(formData.addons);
  sheet.getRange(rowIndex, 16).setValue(user); 
  sheet.getRange(rowIndex, 17).setValue(now);  

  logAudit('VEHICLE_UPDATED', 'Vehicles', formData.rentalId, 'Specs updated');
  return { success: true, message: `Vehicle updated.` };
}

function processReportIncident(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const incSheet = ss.getSheetByName('Incidents_Service');
  const vSheet = ss.getSheetByName('Vehicles');
  const dSheet = ss.getSheetByName('Deployments');
  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';
  const incId = 'INC-' + Utilities.getUuid().substring(0, 8).toUpperCase();

  incSheet.appendRow([
    incId, now, data.cleanName, data.rentalId, data.issueCategory, data.description,
    data.triggerOos ? 'YES' : 'NO', data.statusImpact, '', '', '', user, now
  ]);

  if (data.triggerOos) {
    const vData = vSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        const row = i + 1;
        vSheet.getRange(row, 9).setValue('Out of Service');
        vSheet.getRange(row, 12).clearContent();
        vSheet.getRange(row, 13).clearContent();
        vSheet.getRange(row, 16).setValue(user);
        vSheet.getRange(row, 17).setValue(now);
        break;
      }
    }
    
    // Auto-end the active deployment
    const dData = dSheet.getDataRange().getValues();
    for (let j = 1; j < dData.length; j++) {
      if (String(dData[j][4]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase() && String(dData[j][9]).trim() === 'Dispatched') {
        dSheet.getRange(j + 1, 10).setValue('Completed');
        dSheet.getRange(j + 1, 12).setValue(now);
      }
    }
  }
  logAudit('INCIDENT_REPORTED', 'Incidents_Service', incId, `OOS: ${data.triggerOos}`);
  return { success: true, message: 'Incident reported.' };
}

function processCheckInVehicle(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vSheet = ss.getSheetByName('Vehicles');
  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  if (!data.rentalId) {
    throw new Error('No Rental ID was provided to the database.');
  }

  const vData = vSheet.getDataRange().getValues();
  let found = false;
  
  for (let i = 1; i < vData.length; i++) {
    if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
      vSheet.getRange(i + 1, 9).setValue('Warehouse (Ready)');
      vSheet.getRange(i + 1, 12).clearContent(); // CLEAR ZONE
      vSheet.getRange(i + 1, 13).clearContent(); // CLEAR TECH
      vSheet.getRange(i + 1, 15).clearContent(); // CLEAR OOS NOTES
      vSheet.getRange(i + 1, 16).setValue(user);
      vSheet.getRange(i + 1, 17).setValue(now);
      found = true;
      break;
    }
  }

  if (!found) {
    throw new Error('Vehicle ' + data.rentalId + ' could not be found in the database.');
  }

  // 2. Automatically find and resolve any open incidents for this van
  let targetIncId = data.incidentId;
  const iSheet = ss.getSheetByName('Incidents_Service');
  const iData = iSheet.getDataRange().getValues();
  
  // Fallback: If frontend didn't pass the ID, find the open ticket manually
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
  
  logAudit('VEHICLE_CHECKED_IN', 'Vehicles', data.rentalId, 'Repaired and returned to Warehouse');
  return { success: true, message: 'Van is back in the Warehouse.' };
}

function decommissionVehicle(targetId, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Vehicles');
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === targetId.trim().toLowerCase() || String(data[i][1]).trim().toLowerCase() === targetId.trim().toLowerCase()) {
      rowIndex = i + 1; break;
    }
  }
  if (rowIndex === -1) throw new Error('Vehicle not found.');

  sheet.getRange(rowIndex, 9).setValue('Decommissioned'); 
  sheet.getRange(rowIndex, 12).clearContent(); 
  sheet.getRange(rowIndex, 13).clearContent(); 
  const curNotes = sheet.getRange(rowIndex, 15).getValue();
  sheet.getRange(rowIndex, 15).setValue((curNotes ? curNotes + ' | ' : '') + 'Decom: ' + reason);

  logAudit('VEHICLE_DECOMMISSIONED', 'Vehicles', targetId, reason);
  return { success: true, message: `Vehicle Decommissioned.` };
}