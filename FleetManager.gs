/**
 * @fileoverview Vehicle Registry & Service Operations
 */

function processAddVehicle(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Vehicles');
  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';
  
  sheet.appendRow([
    formData.vanName || '', 
    formData.rentalId || '', 
    formData.rentalAgency || 'Enterprise',
    formData.makeModel || 'Cargo Van', 
    formData.vin || '', 
    formData.licensePlate || '',
    formData.plateExpDate || '', 
    formData.samsaraSerial || '', 
    formData.samsaraLink || '',
    'Warehouse (Ready)', 
    formData.pickupDate || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    formData.serviceNumber || '', 
    '', 
    '', 
    '', 
    '', 
    user, 
    now, 
    formData.addons || '',
    formData.coast || 'No'
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
  if (formData.coast) sheet.getRange(rowIndex, 20).setValue(formData.coast);

  if (typeof logAudit === 'function') logAudit('VEHICLE_UPDATED', 'Vehicles', formData.rentalId, 'Specs updated');
  return { success: true, message: `Vehicle updated.` };
}

function processReportIncident(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vSheet = ss.getSheetByName('Vehicles');
    const iSheet = ss.getSheetByName('Incidents_Service');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    const incidentId = 'INC-' + Utilities.getUuid().substring(0, 6).toUpperCase();

    // 1. Log to Incidents
    iSheet.appendRow([
      incidentId, now, data.cleanName, data.rentalId,
      data.issueCategory, data.description, data.triggerOos ? 'YES' : 'NO',
      data.statusImpact, '', '', ''
    ]);

    // 2. Mark OOS if triggered
    if (data.triggerOos) {
      const vData = vSheet.getDataRange().getValues();
      for (let i = 1; i < vData.length; i++) {
        if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
          const row = i + 1;
          vSheet.getRange(row, 10).setValue('Out of Service'); // Col J
          vSheet.getRange(row, 13).clearContent(); // Clear Zone (Col M)
          vSheet.getRange(row, 14).clearContent(); // Clear Tech (Col N)
          vSheet.getRange(row, 16).setValue(data.issueCategory + ': ' + data.description); // Col P
          vSheet.getRange(row, 17).setValue(user); // Col Q
          vSheet.getRange(row, 18).setValue(now);  // Col R
          break;
        }
      }
    }
    return { success: true, message: 'Incident reported.' };
  } catch (err) {
    return { success: false, message: 'Incident report failed: ' + err.toString() };
  }
}

function processCheckInVehicle(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    // 1. Reset Vehicle
    const vData = vSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        const row = i + 1;
        vSheet.getRange(row, 10).setValue('Warehouse (Ready)'); // Col J
        vSheet.getRange(row, 13).clearContent(); // Clear Zone (Col M)
        vSheet.getRange(row, 14).clearContent(); // Clear Tech (Col N)
        vSheet.getRange(row, 16).clearContent(); // Clear Notes (Col P)
        vSheet.getRange(row, 17).setValue(user); // Col Q
        vSheet.getRange(row, 18).setValue(now);  // Col R
        break;
      }
    }

    // 2. Resolve Incident
    if (data.incidentId) {
      const iSheet = ss.getSheetByName('Incidents_Service');
      const iData = iSheet.getDataRange().getValues();
      for (let j = 1; j < iData.length; j++) {
        if (String(iData[j][0]).trim() === data.incidentId) {
          iSheet.getRange(j + 1, 9).setValue(data.resolutionNotes);
          iSheet.getRange(j + 1, 10).setValue(now);
          iSheet.getRange(j + 1, 11).setValue(user);
          break;
        }
      }
    }
    return { success: true, message: 'Van is back in the Warehouse.' };
  } catch (err) {
    return { success: false, message: 'Check-in failed: ' + err.toString() };
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

  sheet.getRange(rowIndex, 10).setValue('Decommissioned'); 
  sheet.getRange(rowIndex, 13).clearContent(); 
  sheet.getRange(rowIndex, 14).clearContent(); 
  const curNotes = sheet.getRange(rowIndex, 15).getValue();
  sheet.getRange(rowIndex, 15).setValue((curNotes ? curNotes + ' | ' : '') + 'Decom: ' + reason);

  if (typeof logAudit === 'function') logAudit('VEHICLE_DECOMMISSIONED', 'Vehicles', targetId, reason);
  return { success: true, message: `Vehicle Decommissioned.` };
}