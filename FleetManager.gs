/**
 * @fileoverview Vehicle Registry Operations (Add, Decommission, Query).
 */

/**
 * Displays modal dialog for adding a new vehicle.
 */
function showAddVehicleModal() {
  const config = getSystemConfig();
  const template = HtmlService.createTemplateFromFile('Modal_AddVehicle');
  template.rentalAgencies = config.rentalAgencies;

  const html = template.evaluate()
    .setWidth(520)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '➕ Add New Vehicle to Fleet');
}

/**
 * Processes new vehicle submission from modal.
 */
function processAddVehicle(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet) throw new Error('Vehicles sheet not found');

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  // 15-column schema:
  // [CleanName, RentalId, Agency, MakeModel, VIN, Plate, Status, PickupDate, Service#, Zone, Tech, LastService, Notes, UpdatedBy, UpdatedAt]
  const rowData = [
    formData.vanName || '',
    formData.rentalId || '',
    formData.rentalAgency || 'Enterprise',
    formData.makeModel || 'Cargo Van',
    formData.vin || '',
    formData.licensePlate || '',
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
    agency: formData.rentalAgency,
    vin: formData.vin,
    plate: formData.licensePlate,
    addedBy: user
  });

  return { success: true, message: `Vehicle ${formData.vanName} successfully added!` };
}

/**
 * Decommissions a vehicle and marks status in registry.
 */
function decommissionVehicle(targetId, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet) throw new Error('Vehicles sheet not found');

  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    const cleanName = String(data[i][0]).trim().toLowerCase();
    const rentalId = String(data[i][1]).trim().toLowerCase();
    const needle = String(targetId).trim().toLowerCase();

    if (cleanName === needle || rentalId === needle) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('Vehicle "' + targetId + '" not found in registry.');
  }

  const user = Session.getActiveUser().getEmail() || 'local-user';
  const now = new Date();

  sheet.getRange(rowIndex, 7).setValue('Decommissioned'); // Col G
  const existingNotes = sheet.getRange(rowIndex, 13).getValue();
  const updatedNotes = (existingNotes ? existingNotes + ' | ' : '') + 'Decommissioned: ' + (reason || 'No reason provided');
  sheet.getRange(rowIndex, 13).setValue(updatedNotes);
  sheet.getRange(rowIndex, 14).setValue(user);
  sheet.getRange(rowIndex, 15).setValue(now);

  logAudit('VEHICLE_DECOMMISSIONED', 'Vehicles', targetId, {
    reason: reason,
    decommissionedBy: user
  });

  return { success: true, message: `Vehicle ${targetId} marked as Decommissioned.` };
}

/**
 * Helper to fetch active vehicles for dropdowns.
 */
function getActiveVehiclesList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  return data
    .filter(r => r[6] !== 'Decommissioned')
    .map(r => ({
      cleanName: r[0],
      rentalId: r[1],
      status: r[6],
      displayName: `${r[0]} (${r[1]}) - [${r[6]}]`
    }));
}