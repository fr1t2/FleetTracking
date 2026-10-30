function processAddVehicle(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    const row = Array(20).fill('');
    row[0] = data.cleanName || '';
    row[1] = data.rentalId || '';
    row[2] = data.agency || '';
    row[3] = data.makeModel || '';
    row[4] = data.vin || '';
    row[5] = data.plate || '';
    row[6] = data.plateExpDate || '';
    row[7] = data.samsaraSerial || '';
    row[8] = data.samsaraLink || '';
    row[9] = 'Warehouse (Ready)';
    row[10] = data.pickupDate || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    row[11] = data.serviceNum || '';
    row[14] = now; 
    row[16] = user;
    row[17] = now;
    row[18] = data.addons || '';
    row[19] = data.coast || 'No';

    sheet.appendRow(row);
    return { success: true, message: `Unit ${data.rentalId} added to the fleet.` };
  } catch (err) {
    return { success: false, message: 'Failed to add vehicle: ' + err.toString() };
  }
}

function updateVehicleDetails(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Vehicles');
    const rows = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) throw new Error("Vehicle not found.");

    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    // 1-based indexing for getRange(row, column)
    if(data.cleanName !== undefined) sheet.getRange(rowIndex, 1).setValue(data.cleanName);
    if(data.agency !== undefined) sheet.getRange(rowIndex, 3).setValue(data.agency);
    if(data.makeModel !== undefined) sheet.getRange(rowIndex, 4).setValue(data.makeModel);
    if(data.vin !== undefined) sheet.getRange(rowIndex, 5).setValue(data.vin);
    if(data.plate !== undefined) sheet.getRange(rowIndex, 6).setValue(data.plate);
    if(data.plateExpDate !== undefined) sheet.getRange(rowIndex, 7).setValue(data.plateExpDate);
    if(data.samsaraSerial !== undefined) sheet.getRange(rowIndex, 8).setValue(data.samsaraSerial);
    if(data.samsaraLink !== undefined) sheet.getRange(rowIndex, 9).setValue(data.samsaraLink);
    if(data.serviceNum !== undefined) sheet.getRange(rowIndex, 12).setValue(data.serviceNum);
    if(data.addons !== undefined) sheet.getRange(rowIndex, 19).setValue(data.addons);
    if(data.coast !== undefined) sheet.getRange(rowIndex, 20).setValue(data.coast);

    sheet.getRange(rowIndex, 17).setValue(user);
    sheet.getRange(rowIndex, 18).setValue(now);

    return { success: true, message: `Unit ${data.rentalId} updated.` };
  } catch (err) {
    return { success: false, message: 'Failed to update vehicle: ' + err.toString() };
  }
}

function decommissionVehicle(rentalId, reason) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Vehicles');
    const rows = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim().toLowerCase() === String(rentalId).trim().toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) throw new Error("Vehicle not found.");

    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    sheet.getRange(rowIndex, 10).setValue('Decommissioned'); 
    sheet.getRange(rowIndex, 13).clearContent(); 
    sheet.getRange(rowIndex, 14).clearContent(); 
    
    const curNotes = sheet.getRange(rowIndex, 16).getValue();
    sheet.getRange(rowIndex, 16).setValue((curNotes ? curNotes + ' | ' : '') + 'Decommissioned: ' + reason);
    sheet.getRange(rowIndex, 17).setValue(user);
    sheet.getRange(rowIndex, 18).setValue(now);

    return { success: true, message: `Unit ${rentalId} decommissioned.` };
  } catch (err) {
    return { success: false, message: 'Failed to decommission: ' + err.toString() };
  }
}