/**
 * @fileoverview Vehicle Registry Operations (Add, Decommission, Query).
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
 * Returns all vehicles that are NOT decommissioned for form dropdowns.
 */
function getActiveVehiclesList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const list = [];
  
  for (let i = 0; i < data.length; i++) {
    const cleanName = data[i][0];
    const rentalId = data[i][1];
    const status = data[i][6] || 'Active';

    if (status !== 'Decommissioned' && (cleanName || rentalId)) {
      list.push({
        cleanName: String(cleanName).trim(),
        rentalId: String(rentalId).trim(),
        status: status,
        displayName: `${cleanName} (${rentalId}) - [${status}]`
      });
    }
  }
  return list;
}

/**
 * One-click migration: Safely upgrades the Vehicles table to 17 columns
 * and preserves existing vehicle records (Van #19, Van #20, etc.).
 */
function upgradeVehiclesSchemaToV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet) throw new Error('Vehicles sheet not found');

  const v2Headers = [
    'Vehicle Clean Name',         // Col A
    'Rental Agreement / Unit #',  // Col B
    'Rental Agency',              // Col C
    'Make & Model',               // Col D
    'VIN',                        // Col E
    'License Plate',              // Col F
    'Plate Expiration Date',      // Col G (NEW)
    'Samsara Gateway Serial #',   // Col H (NEW)
    'Current Status',             // Col I
    'Pick-up Date',               // Col J
    'Service Number',             // Col K
    'Current Assigned Zone',      // Col L
    'Current Assigned Tech',      // Col M
    'Last Service Date',          // Col N
    'Notes',                      // Col O
    'Last Updated By',            // Col P
    'Last Updated At'             // Col Q
  ];

  const lastRow = sheet.getLastRow();
  const tz = Session.getScriptTimeZone();

  if (lastRow > 1) {
    const oldData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const upgradedRows = oldData.map(r => {
      // If already upgraded (17 cols), keep; otherwise migrate from 15 cols:
      if (r.length >= 17) return r;

      const cleanName   = r[0];
      const rentalId    = r[1];
      const agency      = r[2];
      const makeModel   = r[3];
      const vin         = r[4];
      const plate       = r[5];
      const status      = r[6] || 'Active';
      const pickupDate  = r[7];
      const serviceNum  = r[8];
      const zone        = r[9];
      const tech        = r[10];
      const lastService = r[11];
      const notes       = r[12];
      const updatedBy   = r[13];
      const updatedAt   = r[14];

      // Default sample Samsara & Tag for Van #19 if empty
      const samsaraId = (cleanName === 'Van #19') ? 'GDM3-683-XW2' : '';
      const tagExp    = (cleanName === 'Van #19') ? '2027-01-31' : '';

      return [
        cleanName, rentalId, agency, makeModel, vin, plate,
        tagExp, samsaraId, status, pickupDate, serviceNum,
        zone, tech, lastService, notes, updatedBy, updatedAt
      ];
    });

    sheet.clearContents();
    sheet.getRange(1, 1, 1, v2Headers.length).setValues([v2Headers]);
    sheet.getRange(1, 1, 1, v2Headers.length).setFontWeight('bold');
    sheet.getRange(2, 1, upgradedRows.length, v2Headers.length).setValues(upgradedRows);
  } else {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, v2Headers.length).setValues([v2Headers]);
    sheet.getRange(1, 1, 1, v2Headers.length).setFontWeight('bold');
  }

  logAudit('SCHEMA_UPGRADE', 'Vehicles', 'V2_SAMSARA', 'Upgraded schema with Samsara Gateway and Plate Expiration');
  SpreadsheetApp.getActiveSpreadsheet().toast('Vehicles sheet upgraded to 17-column Samsara schema!', 'Schema Updated', 4);
}

/**
 * Updates editable vehicle properties (Name, Plate, Samsara Serial, Expiration, Notes).
 */
function updateVehicleDetails(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet) throw new Error('Vehicles sheet not found');

  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  const targetRentalId = String(formData.rentalId).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === targetRentalId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error(`Vehicle with Rental Unit ${formData.rentalId} not found.`);

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  // Update specific fields:
  // Col A: Clean Name, Col F: License Plate, Col G: Plate Expiration, Col H: Samsara, Col O: Notes
  if (formData.vanName) sheet.getRange(rowIndex, 1).setValue(formData.vanName);
  if (formData.licensePlate !== undefined) sheet.getRange(rowIndex, 6).setValue(formData.licensePlate);
  if (formData.plateExpDate !== undefined) sheet.getRange(rowIndex, 7).setValue(formData.plateExpDate);
  if (formData.samsaraSerial !== undefined) sheet.getRange(rowIndex, 8).setValue(formData.samsaraSerial);
  
  if (formData.notes) {
    const curNotes = sheet.getRange(rowIndex, 15).getValue();
    sheet.getRange(rowIndex, 15).setValue((curNotes ? curNotes + ' | ' : '') + formData.notes);
  }

  sheet.getRange(rowIndex, 16).setValue(user); // Col P: Last Updated By
  sheet.getRange(rowIndex, 17).setValue(now);  // Col Q: Last Updated At

  logAudit('VEHICLE_UPDATED', 'Vehicles', formData.rentalId, {
    vanName: formData.vanName,
    licensePlate: formData.licensePlate,
    plateExpDate: formData.plateExpDate,
    samsaraSerial: formData.samsaraSerial,
    updatedBy: user
  });

  return { success: true, message: `Vehicle ${formData.vanName || formData.rentalId} updated successfully!` };
}