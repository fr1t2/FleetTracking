/**
 * @fileoverview Vehicle Registry Operations (Add, Edit, Decommission).
 */

function processAddVehicle(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet) throw new Error('Vehicles sheet not found');

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  // 17-column format:
  // [CleanName, RentalId, Agency, MakeModel, VIN, Plate, PlateExp, SamsaraSerial, Status, PickupDate, Service#, Zone, Tech, LastService, Notes, UpdatedBy, UpdatedAt]
  const rowData = [
    formData.vanName || '',
    formData.rentalId || '',
    formData.rentalAgency || 'Enterprise',
    formData.makeModel || 'Cargo Van',
    formData.vin || '',
    formData.licensePlate || '',
    formData.plateExpDate || '',
    formData.samsaraSerial || '',
    'Active',
    formData.pickupDate ? new Date(formData.pickupDate) : now,
    formData.serviceNum || '',
    '',   // Zone
    '',   // Tech
    now,  // Last service date
    formData.notes || '',
    user,
    now
  ];

  sheet.appendRow(rowData);

  logAudit('VEHICLE_ADDED', 'Vehicles', formData.rentalId || formData.vanName, {
    vanName: formData.vanName,
    rentalId: formData.rentalId,
    samsara: formData.samsaraSerial,
    plateExp: formData.plateExpDate,
    addedBy: user
  });

  return { success: true, message: `Vehicle ${formData.vanName} added successfully!` };
}

function updateVehicleDetails(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet) throw new Error('Vehicles sheet not found');

  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  const targetRentalId = String(formData.rentalId || '').trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === targetRentalId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error(`Vehicle ${formData.rentalId} not found in registry.`);

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  if (formData.vanName) sheet.getRange(rowIndex, 1).setValue(formData.vanName);
  if (formData.licensePlate !== undefined) sheet.getRange(rowIndex, 6).setValue(formData.licensePlate);
  if (formData.plateExpDate !== undefined) sheet.getRange(rowIndex, 7).setValue(formData.plateExpDate);
  if (formData.samsaraSerial !== undefined) sheet.getRange(rowIndex, 8).setValue(formData.samsaraSerial);

  if (formData.notes) {
    const curNotes = sheet.getRange(rowIndex, 15).getValue();
    sheet.getRange(rowIndex, 15).setValue((curNotes ? curNotes + ' | ' : '') + formData.notes);
  }

  sheet.getRange(rowIndex, 16).setValue(user);
  sheet.getRange(rowIndex, 17).setValue(now);

  logAudit('VEHICLE_UPDATED', 'Vehicles', formData.rentalId, {
    vanName: formData.vanName,
    plate: formData.licensePlate,
    plateExp: formData.plateExpDate,
    samsara: formData.samsaraSerial,
    updatedBy: user
  });

  return { success: true, message: `Vehicle ${formData.vanName || formData.rentalId} updated successfully!` };
}

function decommissionVehicle(targetId, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet) throw new Error('Vehicles sheet not found');

  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  const needle = String(targetId || '').trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const cleanName = String(data[i][0]).trim().toLowerCase();
    const rentalId = String(data[i][1]).trim().toLowerCase();

    if (cleanName === needle || rentalId === needle) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error('Vehicle "' + targetId + '" not found in registry.');

  const user = Session.getActiveUser().getEmail() || 'local-user';
  const now = new Date();

  sheet.getRange(rowIndex, 9).setValue('Decommissioned'); // Col I is Status in 17-col layout
  const existingNotes = sheet.getRange(rowIndex, 15).getValue();
  const updatedNotes = (existingNotes ? existingNotes + ' | ' : '') + 'Decommissioned: ' + (reason || 'No reason provided');
  sheet.getRange(rowIndex, 15).setValue(updatedNotes);
  sheet.getRange(rowIndex, 16).setValue(user);
  sheet.getRange(rowIndex, 17).setValue(now);

  logAudit('VEHICLE_DECOMMISSIONED', 'Vehicles', targetId, {
    reason: reason,
    decommissionedBy: user
  });

  return { success: true, message: `Vehicle ${targetId} marked as Decommissioned.` };
}