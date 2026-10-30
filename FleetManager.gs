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

/**
 * Resolves a specific incident ticket directly from the history logs.
 */
function resolveSpecificIncident(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const iSheet = ss.getSheetByName('Incidents_Service');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    if (!data.incidentId) throw new Error("No Incident ID provided.");

    const iData = iSheet.getDataRange().getValues();
    let found = false;
    for (let j = 1; j < iData.length; j++) {
      if (String(iData[j][0]).trim() === data.incidentId) {
        iSheet.getRange(j + 1, 9).setValue(data.resolutionNotes); // Col I: Resolution Notes
        iSheet.getRange(j + 1, 10).setValue(now);                 // Col J: Date Resolved
        iSheet.getRange(j + 1, 11).setValue(user);                // Col K: Resolved By
        found = true;
        break;
      }
    }
    if (!found) throw new Error("Incident not found.");
    
    if (typeof logAudit === 'function') {
      logAudit('INCIDENT_RESOLVED', 'Incidents_Service', data.incidentId, `Resolved by ${user}: ${data.resolutionNotes}`);
    }
    
    return { success: true, message: `Incident resolved.` };
  } catch (err) {
    return { success: false, message: 'Failed to resolve incident: ' + err.toString() };
  }
}

function processReportIncident(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const incSheet = ss.getSheetByName('Incidents_Service');
    const vSheet = ss.getSheetByName('Vehicles');
    const dSheet = ss.getSheetByName('Deployments');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    
    // Safety Checks
    if (!incSheet) throw new Error("Could not find the 'Incidents_Service' tab.");
    if (!vSheet) throw new Error("Could not find the 'Vehicles' tab.");
    if (!data.rentalId && !data.cleanName) throw new Error("No Vehicle identifier was provided.");
    
    const incId = 'INC-' + Utilities.getUuid().substring(0, 8).toUpperCase();

    // 1. Log the Incident
    incSheet.appendRow([
      incId, now, data.cleanName || '', data.rentalId || '', data.issueCategory || 'General', 
      data.description || '', data.triggerOos ? 'YES' : 'NO', data.statusImpact || '', 
      '', '', '', user, now, data.driverInvolved || ''
    ]);

    if (data.triggerOos) {
      // 2. Update Vehicle Sheet (Aggressively overwrite Zone and Tech)
      let vFound = false;
      const vData = vSheet.getDataRange().getValues();
      for (let i = 1; i < vData.length; i++) {
        const rowRentalId = String(vData[i][1]).trim().toLowerCase();
        const targetRentalId = String(data.rentalId || '').trim().toLowerCase();

        if (targetRentalId && rowRentalId === targetRentalId) {
          const row = i + 1;
          vSheet.getRange(row, 10).setValue('Out of Service'); // Status is Col 10
          vSheet.getRange(row, 13).clearContent(); // Zone is Col 13
          vSheet.getRange(row, 14).clearContent(); // Tech is Col 14
          vSheet.getRange(row, 17).setValue(user); // Updated By
          vSheet.getRange(row, 18).setValue(now);  // Updated At
          vFound = true;
          break;
        }
      }
      if (!vFound) throw new Error(`Vehicle Unit ${data.rentalId} not found in Vehicles tab.`);
      
      // 3. Auto-End Active Deployments
      if (dSheet) {
        const dData = dSheet.getDataRange().getValues();
        for (let j = 1; j < dData.length; j++) {
          const depStatus = String(dData[j][9]).trim().toLowerCase();
          const rowRentalId = String(dData[j][4]).trim().toLowerCase();
          const targetRentalId = String(data.rentalId || '').trim().toLowerCase();

          if ((depStatus === 'dispatched' || depStatus === 'active') && targetRentalId && rowRentalId === targetRentalId) {
            const row = j + 1;
            dSheet.getRange(row, 10).clearDataValidations(); 
            dSheet.getRange(row, 10).setValue('Completed'); // End shift
            dSheet.getRange(row, 12).setValue(now); // Return time
          }
        }
      }
    }
    
    if (typeof logAudit === 'function') logAudit('INCIDENT_REPORTED', 'Incidents_Service', incId, `OOS: ${data.triggerOos}`);
    return { success: true, message: 'Incident reported successfully.' };
    
  } catch (err) {
    // If it fails, bundle the exact error string so the UI can show it
    return { success: false, message: 'Failed to report incident: ' + err.message };
  }
}