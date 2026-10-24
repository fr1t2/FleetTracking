/**
 * @fileoverview Vehicle Registry Operations
 */

function processAddVehicle(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  // 18-column format now includes Addons & Interior Builds at Col R
  const rowData = [
    formData.vanName || '',
    formData.rentalId || '',
    formData.rentalAgency || 'Enterprise',
    formData.makeModel || 'Cargo Van',
    formData.vin || '',
    formData.licensePlate || '',
    formData.plateExpDate || '',
    formData.samsaraSerial || '',
    'Warehouse (Ready)',          // New default status
    formData.pickupDate ? new Date(formData.pickupDate) : now,
    formData.serviceNum || '',
    '',   // Zone
    '',   // Tech
    now,  // Last service date
    formData.notes || '',
    user,
    now,
    formData.addons || ''         // Col R: Addons
  ];
  sheet.appendRow(rowData);
  logAudit('VEHICLE_ADDED', 'Vehicles', formData.rentalId, 'Added to fleet');
  return { success: true, message: `Vehicle ${formData.vanName} added!` };
}

function updateVehicleDetails(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  const targetRentalId = String(formData.rentalId).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === targetRentalId) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error(`Vehicle ${formData.rentalId} not found.`);

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  if (formData.vanName !== undefined) sheet.getRange(rowIndex, 1).setValue(formData.vanName);
  if (formData.licensePlate !== undefined) sheet.getRange(rowIndex, 6).setValue(formData.licensePlate);
  if (formData.plateExpDate !== undefined) sheet.getRange(rowIndex, 7).setValue(formData.plateExpDate);
  if (formData.samsaraSerial !== undefined) sheet.getRange(rowIndex, 8).setValue(formData.samsaraSerial);
  if (formData.addons !== undefined) sheet.getRange(rowIndex, 18).setValue(formData.addons);
  if (formData.notes) {
    const curNotes = sheet.getRange(rowIndex, 15).getValue();
    sheet.getRange(rowIndex, 15).setValue((curNotes ? curNotes + ' | ' : '') + formData.notes);
  }
  sheet.getRange(rowIndex, 16).setValue(user); 
  sheet.getRange(rowIndex, 17).setValue(now);  

  logAudit('VEHICLE_UPDATED', 'Vehicles', formData.rentalId, 'Vehicle specs updated');
  return { success: true, message: `Vehicle ${formData.vanName || formData.rentalId} updated successfully!` };
}

function decommissionVehicle(targetId, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === targetId.trim().toLowerCase() || 
        String(data[i][1]).trim().toLowerCase() === targetId.trim().toLowerCase()) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error('Vehicle not found.');

  const user = Session.getActiveUser().getEmail();
  const now = new Date();
  
  sheet.getRange(rowIndex, 9).setValue('Decommissioned'); // Col I is Status
  sheet.getRange(rowIndex, 12).clearContent(); // Clear Zone
  sheet.getRange(rowIndex, 13).clearContent(); // Clear Tech
  const curNotes = sheet.getRange(rowIndex, 15).getValue();
  sheet.getRange(rowIndex, 15).setValue((curNotes ? curNotes + ' | ' : '') + 'Decommissioned: ' + reason);
  sheet.getRange(rowIndex, 16).setValue(user);
  sheet.getRange(rowIndex, 17).setValue(now);

  logAudit('VEHICLE_DECOMMISSIONED', 'Vehicles', targetId, reason);
  return { success: true, message: `Vehicle marked as Decommissioned.` };
}

function processReportIncident(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const incSheet = ss.getSheetByName('Incidents_Service');
    const vSheet = ss.getSheetByName('Vehicles');
    const dSheet = ss.getSheetByName('Deployments');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    const incId = 'INC-' + Utilities.getUuid().substring(0, 8).toUpperCase();

    // 1. Log the Incident
    incSheet.appendRow([
      incId, now, data.cleanName, data.rentalId,
      data.issueCategory, data.description,
      data.triggerOos ? 'YES' : 'NO', data.statusImpact,
      '', '', '', user, now
    ]);

    // 2. Handle Workflow if OOS is Triggered
    if (data.triggerOos) {
      
      // A. Update Vehicle Status and Strip the Driver / Zone Assignment
      const vData = vSheet.getDataRange().getValues();
      for (let i = 1; i < vData.length; i++) {
        if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
          const row = i + 1;
          vSheet.getRange(row, 9).setValue('Out of Service'); // Set Status
          vSheet.getRange(row, 12).clearContent();            // Clear Zone
          vSheet.getRange(row, 13).clearContent();            // Clear Assigned Tech
          vSheet.getRange(row, 16).setValue(user);
          vSheet.getRange(row, 17).setValue(now);
          break;
        }
      }

      // B. Auto-End the Active Shift for this Van
      if (dSheet) {
        const dData = dSheet.getDataRange().getValues();
        for (let j = 1; j < dData.length; j++) {
          // If the rental ID matches AND the shift is currently 'Dispatched'
          if (String(dData[j][4]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase() && 
              String(dData[j][9]).trim() === 'Dispatched') {
            const dRow = j + 1;
            dSheet.getRange(dRow, 10).setValue('Completed'); // End the shift
            dSheet.getRange(dRow, 12).setValue(now);         // Log the exact return time
          }
        }
      }
    }

    logAudit('INCIDENT_REPORTED', 'Incidents_Service', incId, `OOS Triggered: ${data.triggerOos}`);
    return { success: true, message: 'Incident reported successfully.' };
  } catch (err) {
    throw new Error(err.message);
  }
}