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