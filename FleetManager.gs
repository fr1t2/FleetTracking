/**
 * @fileoverview Vehicle Registry & Service Operations
 */

function processAddVehicle(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Vehicles');
  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';
  
  // Note the new parameter layout aligning with the 19 columns
  sheet.appendRow([
    formData.vanName || '', formData.rentalId || '', formData.rentalAgency || 'Enterprise',
    formData.makeModel || 'Cargo Van', formData.vin || '', formData.licensePlate || '',
    formData.plateExpDate || '', formData.samsaraSerial || '', formData.samsaraLink || '',
    'Warehouse (Ready)', formData.pickupDate || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    formData.serviceNumber || '', '', '', '', '', user, now, formData.addons || ''
  ]);
  
  if (typeof logAudit === 'function') logAudit('VEHICLE_ADDED', 'Vehicles', formData.rentalId, 'Added to fleet');
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

  // Aligning strictly to the 1-based columns
  sheet.getRange(rowIndex, 1).setValue(formData.vanName);
  sheet.getRange(rowIndex, 3).setValue(formData.rentalAgency);
  sheet.getRange(rowIndex, 4).setValue(formData.makeModel);
  sheet.getRange(rowIndex, 5).setValue(formData.vin);
  sheet.getRange(rowIndex, 6).setValue(formData.licensePlate);
  sheet.getRange(rowIndex, 7).setValue(formData.plateExpDate);
  sheet.getRange(rowIndex, 8).setValue(formData.samsaraSerial);
  sheet.getRange(rowIndex, 9).setValue(formData.samsaraLink);
  sheet.getRange(rowIndex, 11).setValue(formData.pickupDate);
  sheet.getRange(rowIndex, 12).setValue(formData.serviceNumber);
  sheet.getRange(rowIndex, 17).setValue(user); 
  sheet.getRange(rowIndex, 18).setValue(now); 
  sheet.getRange(rowIndex, 19).setValue(formData.addons);

  if (typeof logAudit === 'function') logAudit('VEHICLE_UPDATED', 'Vehicles', formData.rentalId, 'Specs updated');
  return { success: true, message: `Vehicle updated.` };
}

function processReportIncident(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const incSheet = ss.getSheetByName('Incidents_Service');
    const vSheet = ss.getSheetByName('Vehicles');
    const dSheet = ss.getSheetByName('Deployments');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    if (!data.rentalId && !data.cleanName) throw new Error("No Vehicle identifier was provided.");
    
    const incId = 'INC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    incSheet.appendRow([incId, now, data.cleanName, data.rentalId, data.issueCategory, data.description, data.triggerOos ? 'YES' : 'NO', data.statusImpact, '', '', '', user, now, data.driverInvolved || '']);

    if (data.triggerOos) {
      let vFound = false;
      const vData = vSheet.getDataRange().getValues();
      for (let i = 1; i < vData.length; i++) {
        const rowRentalId = String(vData[i][1]).trim().toLowerCase();
        const rowCleanName = String(vData[i][0]).trim().toLowerCase();
        const targetRentalId = String(data.rentalId || '').trim().toLowerCase();
        const targetCleanName = String(data.cleanName || '').trim().toLowerCase();

        if ((targetRentalId && rowRentalId === targetRentalId) || (targetCleanName && rowCleanName === targetCleanName)) {
          const row = i + 1;
          vSheet.getRange(row, 10).setValue('Out of Service'); // Status shifted to 10
          vSheet.getRange(row, 13).clearContent(); // Zone is 13
          vSheet.getRange(row, 14).clearContent(); // Tech is 14
          vSheet.getRange(row, 17).setValue(user);
          vSheet.getRange(row, 18).setValue(now);
          vFound = true; break;
        }
      }
      if (!vFound) throw new Error("Vehicle not found in registry to update status.");
      
      const dData = dSheet.getDataRange().getValues();
      for (let j = 1; j < dData.length; j++) {
        const depStatus = String(dData[j][9]).trim().toLowerCase();
        const rowRentalId = String(dData[j][4]).trim().toLowerCase();
        const rowCleanName = String(dData[j][3]).trim().toLowerCase();
        const targetRentalId = String(data.rentalId || '').trim().toLowerCase();
        const targetCleanName = String(data.cleanName || '').trim().toLowerCase();

        if ((depStatus === 'dispatched' || depStatus === 'active') && ((targetRentalId && rowRentalId === targetRentalId) || (targetCleanName && rowCleanName === targetCleanName))) {
          const row = j + 1;
          dSheet.getRange(row, 10).clearDataValidations(); 
          dSheet.getRange(row, 10).setValue('Completed'); 
          dSheet.getRange(row, 12).setValue(now); 
        }
      }
    }
    if (typeof logAudit === 'function') logAudit('INCIDENT_REPORTED', 'Incidents_Service', incId, `OOS: ${data.triggerOos}`);
    return { success: true, message: 'Incident reported successfully.' };
  } catch (err) {
    return { success: false, message: 'Failed to report incident: ' + err.toString() };
  }
}

function processCheckInVehicle(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    if (!data.rentalId && !data.cleanName) return { success: false, message: 'No vehicle identifier was provided to the database.' };

    const vData = vSheet.getDataRange().getValues();
    let found = false;
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
        vSheet.getRange(i + 1, 17).setValue(user);
        vSheet.getRange(i + 1, 18).setValue(now);
        found = true; break;
      }
    }

    if (!found) return { success: false, message: `Vehicle ${data.cleanName || data.rentalId} could not be found in the database.` };

    const iSheet = ss.getSheetByName('Incidents_Service');
    if (iSheet) {
      const iData = iSheet.getDataRange().getValues();
      let targetIncId = data.incidentId;
      if (!targetIncId) {
         for (let j = iData.length - 1; j >= 1; j--) {
           if (String(iData[j][3]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase() && !iData[j][9]) { targetIncId = String(iData[j][0]).trim(); break; }
         }
      }
      if (targetIncId) {
        for (let j = 1; j < iData.length; j++) {
          if (String(iData[j][0]).trim() === targetIncId) {
            iSheet.getRange(j + 1, 9).setValue(data.resolutionNotes || 'Checked in via dashboard');
            iSheet.getRange(j + 1, 10).setValue(now);
            iSheet.getRange(j + 1, 11).setValue(user); break;
          }
        }
      }
    }
    if (typeof logAudit === 'function') logAudit('VEHICLE_CHECKED_IN', 'Vehicles', data.rentalId || data.cleanName, 'Repaired and returned to Warehouse');
    return { success: true, message: 'Van is back in the Warehouse.' };
  } catch (err) {
    return { success: false, message: 'Check-in script failed: ' + err.toString() };
  }
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

  if (typeof logAudit === 'function') logAudit('VEHICLE_DECOMMISSIONED', 'Vehicles', targetId, reason);
  return { success: true, message: `Vehicle Decommissioned.` };
}