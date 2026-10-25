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

  // 1. AUTO-RESOLVE CONFLICTS (Enforce single-van/single-driver rules)
  const dData = deploySheet.getDataRange().getValues();
  const vData = vehicleSheet.getDataRange().getValues();
  
  for (let j = 1; j < dData.length; j++) {
    if (String(dData[j][9]).trim() === 'Dispatched') {
      const depRental = String(dData[j][4]).trim().toLowerCase();
      const lead = String(dData[j][5]).trim().toLowerCase();
      const co = String(dData[j][14]).trim().toLowerCase();
      
      const targetRental = String(data.rentalId).trim().toLowerCase();
      const targetLead = String(data.assignedTech).trim().toLowerCase();
      const targetCo = String(data.coDriver).trim().toLowerCase();

      let isConflict = false;
      if (depRental === targetRental) isConflict = true; // Van is already assigned
      if (lead && (lead === targetLead || (targetCo && lead === targetCo))) isConflict = true; // Lead tech is already driving
      if (co && (co === targetLead || (targetCo && co === targetCo))) isConflict = true; // Co-driver is already driving

      if (isConflict) {
        // Complete the conflicting shift
        deploySheet.getRange(j + 1, 10).setValue('Completed');
        deploySheet.getRange(j + 1, 12).setValue(now);
        
        // If the tech was driving a DIFFERENT van, reset that old van back to Warehouse (Ready)
        if (depRental !== targetRental) {
          for (let k = 1; k < vData.length; k++) {
            if (String(vData[k][1]).trim().toLowerCase() === depRental) {
              vehicleSheet.getRange(k + 1, 9).setValue('Warehouse');
              vehicleSheet.getRange(k + 1, 12).clearContent();
              vehicleSheet.getRange(k + 1, 13).clearContent();
              break;
            }
          }
        }
      }
    }
  }

  // 2. LOG NEW DEPLOYMENT
  deploySheet.appendRow([
    deployId, now, data.shift, data.cleanName, data.rentalId,
    data.assignedTech, data.workZone, fullGoals, data.batteryTarget || '',
    'Dispatched', now, '', user, now, data.coDriver || '', data.isTraining ? 'Yes' : 'No'
  ]);

  // 3. UPDATE TARGET VEHICLE
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

  // 1. End the Deployment
  const dRows = dSheet.getDataRange().getValues();
  for (let i = 1; i < dRows.length; i++) {
    if (String(dRows[i][0]).trim() === String(deploymentId).trim()) {
      dSheet.getRange(i + 1, 10).setValue('Completed');
      dSheet.getRange(i + 1, 12).setValue(now); // Set Return Time
      break;
    }
  }

  // 2. Wipe the Driver and Zone from the Vehicle
  const vRows = vSheet.getDataRange().getValues();
  for (let j = 1; j < vRows.length; j++) {
    if (String(vRows[j][1]).trim().toLowerCase() === String(rentalId).trim().toLowerCase()) {
      vSheet.getRange(j + 1, 9).setValue('Warehouse');
      vSheet.getRange(j + 1, 12).clearContent(); // CLEAR ZONE
      vSheet.getRange(j + 1, 13).clearContent(); // CLEAR TECH
      break;
    }
  }
  logAudit('END_SHIFT', 'Deployments', deploymentId, 'Shift completed, tech cleared');
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