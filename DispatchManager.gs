/**
 * @fileoverview Shift Deployment & Export Operations
 */

function processDispatchCheckout(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const deploySheet = ss.getSheetByName('Deployments');
  const vehicleSheet = ss.getSheetByName('Vehicles');
  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';
  const deployId = 'DEP-' + Utilities.getUuid().substring(0, 8).toUpperCase();

  let fullGoals = data.focusCategory ? `[${data.focusCategory}] ${data.goals || ''}` : (data.goals || '');

  deploySheet.appendRow([
    deployId, now, data.shift, data.cleanName, data.rentalId,
    data.assignedTech, data.workZone, fullGoals, data.batteryTarget || '',
    'Dispatched', now, '', user, now, data.coDriver || '', data.isTraining ? 'Yes' : 'No'
  ]);

  const vData = vehicleSheet.getDataRange().getValues();
  for (let i = 1; i < vData.length; i++) {
    if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
      const row = i + 1;
      vehicleSheet.getRange(row, 9).setValue('Dispatched');
      vehicleSheet.getRange(row, 12).setValue(data.workZone);
      const fullTech = data.coDriver ? `${data.assignedTech} & ${data.coDriver}` : data.assignedTech;
      vehicleSheet.getRange(row, 13).setValue(fullTech);
      vehicleSheet.getRange(row, 16).setValue(user);
      vehicleSheet.getRange(row, 17).setValue(now);
      break;
    }
  }
  logAudit('DISPATCH', 'Deployments', deployId, `Dispatched: ${data.assignedTech}`);
  return { success: true };
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dSheet = ss.getSheetByName('Deployments');
  const vSheet = ss.getSheetByName('Vehicles');
  const now = new Date();

  const dRows = dSheet.getDataRange().getValues();
  for (let i = 1; i < dRows.length; i++) {
    if (String(dRows[i][0]).trim() === String(deploymentId).trim()) {
      dSheet.getRange(i + 1, 10).setValue('Completed');
      dSheet.getRange(i + 1, 12).setValue(now);
      break;
    }
  }

  const vRows = vSheet.getDataRange().getValues();
  for (let j = 1; j < vRows.length; j++) {
    if (String(vRows[j][1]).trim() === String(rentalId).trim()) {
      vSheet.getRange(j + 1, 9).setValue('Warehouse (Ready)');
      vSheet.getRange(j + 1, 12).clearContent();
      vSheet.getRange(j + 1, 13).clearContent();
      break;
    }
  }
  logAudit('END_SHIFT', 'Deployments', deploymentId, 'Shift completed');
  return { success: true };
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