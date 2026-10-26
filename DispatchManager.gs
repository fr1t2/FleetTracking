/**
 * @fileoverview Shift Deployment & Export Operations
 */

function processDispatchCheckout(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dSheet = ss.getSheetByName('Deployments');
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    const depId = 'DEP-' + Utilities.getUuid().substring(0, 6).toUpperCase();

    // 1. Create the new Dispatch ticket
    dSheet.appendRow([
      depId, now, data.shift, data.cleanName, data.rentalId, data.assignedTech,
      data.workZone, data.goals, data.batteryTarget, 'Dispatched',
      now, '', user, now, data.coDriver || '', data.isTraining ? 'Yes' : 'No'
    ]);

    // 2. Update the Vehicle (Using new shifted columns!)
    const vData = vSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        const row = i + 1;
        vSheet.getRange(row, 10).setValue('Dispatched'); // Col J: Status
        vSheet.getRange(row, 13).setValue(data.workZone); // Col M: Zone
        vSheet.getRange(row, 14).setValue(data.coDriver ? data.assignedTech + ' & ' + data.coDriver : data.assignedTech); // Col N: Tech
        vSheet.getRange(row, 17).setValue(user); // Col Q: Updated By
        vSheet.getRange(row, 18).setValue(now); // Col R: Updated At
        found = true;
        break;
      }
    }

    if (!found) return { success: false, message: `Could not find Vehicle ${data.cleanName} in the sheet to update.` };

    if (typeof logAudit === 'function') logAudit('VEHICLE_DISPATCHED', 'Deployments', depId, `Tech: ${data.assignedTech}`);
    return { success: true, message: 'Van Dispatched successfully.' };
    
  } catch (err) {
    return { success: false, message: 'Dispatch failed: ' + err.toString() };
  }
}

function updateDeploymentRecord(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Deployments');
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(data.id).trim()) {
      sheet.getRange(i + 1, 7).setValue(data.zone);
      sheet.getRange(i + 1, 8).setValue(data.goals);
      sheet.getRange(i + 1, 9).setValue(data.target);
      logAudit('EDIT_DISPATCH', 'Deployments', data.id, 'Updated zone/target/goals');
      return { success: true };
    }
  }
  throw new Error('Deployment not found.');
}

function endShift(deploymentId, rentalId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dSheet = ss.getSheetByName('Deployments');
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();

    // 1. End the Deployment
    const dRows = dSheet.getDataRange().getValues();
    for (let i = 1; i < dRows.length; i++) {
      if (String(dRows[i][0]).trim() === String(deploymentId).trim()) {
        dSheet.getRange(i + 1, 10).clearDataValidations();
        dSheet.getRange(i + 1, 10).setValue('Completed');
        dSheet.getRange(i + 1, 12).setValue(now);
        break;
      }
    }

    // 2. Return Vehicle to Warehouse (Using new shifted columns!)
    const vRows = vSheet.getDataRange().getValues();
    for (let j = 1; j < vRows.length; j++) {
      if (String(vRows[j][1]).trim().toLowerCase() === String(rentalId).trim().toLowerCase()) {
        const row = j + 1;
        vSheet.getRange(row, 10).setValue('Warehouse (Ready)'); // Col J: Status
        vSheet.getRange(row, 13).clearContent(); // Col M: Clear Zone
        vSheet.getRange(row, 14).clearContent(); // Col N: Clear Driver
        break;
      }
    }
    if (typeof logAudit === 'function') logAudit('END_SHIFT', 'Deployments', deploymentId, 'Shift completed, tech cleared');
    return { success: true, message: 'Shift ended successfully.' };
  } catch (err) {
    return { success: false, message: 'End shift failed: ' + err.toString() };
  }
}

function exportShiftManifestCsv(dateVal, shiftName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Deployments');
  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  let csv = 'ID,Date,Shift,Vehicle,Rental ID,Lead Tech,Co-Driver,Zone,Target,Status,Time Out,Time In\n';

  for (let i = 1; i < data.length; i++) {
    let r = data[i];
    let dDate = (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]);
    if (dDate === dateVal && String(r[2]).trim() === shiftName.trim()) {
      let tOut = r[10] instanceof Date ? Utilities.formatDate(r[10], tz, 'HH:mm') : r[10];
      let tIn = r[11] instanceof Date ? Utilities.formatDate(r[11], tz, 'HH:mm') : r[11];
      let row = [r[0], dDate, r[2], r[3], r[4], r[5], r[14], r[6], r[8], r[9], tOut, tIn];
      csv += row.map(v => `"${String(v || '').replace(/"/g, '""')}"`).join(',') + '\n';
    }
  }
  return { success: true, csvContent: csv, filename: `Manifest_${dateVal}_${shiftName.replace(/\s+/g,'')}.csv` };
}

function getPrintableShiftManifestHtml(dateVal, shiftName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Deployments');
  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  
  let html = `<html style="font-family:sans-serif; padding:20px;"><h2>Shift Manifest: ${shiftName} (${dateVal})</h2>`;
  html += `<table border="1" cellpadding="8" style="border-collapse:collapse; width:100%; text-align:left;">`;
  html += `<tr style="background:#eee;"><th>Vehicle</th><th>Unit #</th><th>Lead Tech</th><th>Co-Driver</th><th>Zone</th><th>Time Out</th></tr>`;
  
  for (let i = 1; i < data.length; i++) {
    let r = data[i];
    let dDate = (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]);
    if (dDate === dateVal && String(r[2]).trim() === shiftName.trim()) {
      let tOut = r[10] instanceof Date ? Utilities.formatDate(r[10], tz, 'HH:mm') : r[10];
      html += `<tr><td>${r[3]}</td><td>${r[4]}</td><td>${r[5]}</td><td>${r[14] || ''}</td><td>${r[6]}</td><td>${tOut}</td></tr>`;
    }
  }
  html += `</table><script>window.print();</script></html>`;
  return html;
}