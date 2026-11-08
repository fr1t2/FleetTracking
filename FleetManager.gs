/**
 * @fileoverview Fleet Maintenance and Check-In Engine
 */

function processCheckInVehicle(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    if (!data.rentalId && !data.cleanName) {
      throw new Error('No vehicle identifier provided.');
    }

    const vData = vSheet.getDataRange().getValues();
    let found = false;
    
    // Vehicles: Status (J/10), Zone (M/13), Driver (N/14), Notes (P/16), SafeToDrive (false)
    for (let i = 1; i < vData.length; i++) {
      const rowRentalId = String(vData[i][1]).trim().toLowerCase();
      const rowCleanName = String(vData[i][0]).trim().toLowerCase();
      const targetRentalId = String(data.rentalId || '').trim().toLowerCase();
      const targetCleanName = String(data.cleanName || '').trim().toLowerCase();
      
      if ((targetRentalId && rowRentalId === targetRentalId) || (targetCleanName && rowCleanName === targetCleanName)) {
        vSheet.getRange(i + 1, 10).setValue('Warehouse (Ready)');
        vSheet.getRange(i + 1, 13).clearContent(); 
        vSheet.getRange(i + 1, 14).clearContent(); 
        vSheet.getRange(i + 1, 16).clearContent(); 
        vSheet.getRange(i + 1, 17).setValue(user); // Last Updated By
        vSheet.getRange(i + 1, 18).setValue(now);  // Last Updated At
        found = true;
        break;
      }
    }

    if (!found) throw new Error(`Vehicle ${data.cleanName || data.rentalId} not found.`);

    // Auto-Resolve open incidents
    let targetIncId = data.incidentId;
    const iSheet = ss.getSheetByName('Incidents_Service');
    
    if (iSheet) {
      const iData = iSheet.getDataRange().getValues();
      if (!targetIncId) {
         for (let j = iData.length - 1; j >= 1; j--) {
           if (String(iData[j][3]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase() && !iData[j][9]) {
             targetIncId = String(iData[j][0]).trim();
             break;
           }
         }
      }
      
      if (targetIncId) {
        for (let j = 1; j < iData.length; j++) {
          if (String(iData[j][0]).trim() === targetIncId) {
            iSheet.getRange(j + 1, 9).setValue(data.resolutionNotes || 'Checked in via dashboard');
            iSheet.getRange(j + 1, 10).setValue(now); 
            iSheet.getRange(j + 1, 11).setValue(user);
            break;
          }
        }
      }
    }
    
    return { success: true, message: 'Van is back in the Warehouse.' };
  } catch (err) {
    throw new Error('Check-in failed: ' + err.message);
  }
}

function logNewServiceNeed(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const iSheet = ss.getSheetByName('Incidents_Service');
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    const incidentId = 'INC-' + Utilities.getUuid().substring(0, 6).toUpperCase();

    // Ensure we handle missing optional fields gracefully
    const statusImpact = data.statusImpact || '';
    const severity = data.severity || (data.isOos ? 'OOS' : 'Minor');

    // Incidents Row: 14 Columns
    // [ID, Date, Name, Rental ID, Category, Description, OOS Triggered, Status Impact, Resolution, Date Resolved, Resolved By, Logged By, Timestamp, Severity]
    iSheet.appendRow([
      incidentId, now, data.cleanName, data.rentalId,
      data.category, data.description, data.isOos ? 'YES' : 'NO',
      statusImpact, '', '', '', user, now, severity
    ]);

    if (data.isOos) {
      const vData = vSheet.getDataRange().getValues();
      for (let i = 1; i < vData.length; i++) {
        if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
          const row = i + 1;
          vSheet.getRange(row, 10).setValue('Out of Service'); // Status
          vSheet.getRange(row, 13).clearContent(); // Clear Zone
          vSheet.getRange(row, 14).clearContent(); // Clear Tech
          vSheet.getRange(row, 16).setValue(data.description); // Set Notes
          vSheet.getRange(row, 17).setValue(user); // Updated By
          vSheet.getRange(row, 18).setValue(now);  // Updated At
          break;
        }
      }
    }
    return { success: true };
  } catch (err) {
    throw new Error('Log service failed: ' + err.message);
  }
}

function closeServiceIssue(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const iSheet = ss.getSheetByName('Incidents_Service');
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    const iData = iSheet.getDataRange().getValues();
    
    // 1. Mark the specific ticket as resolved
    for (let i = 1; i < iData.length; i++) {
      if (String(iData[i][0]).trim() === data.incidentId) {
        iSheet.getRange(i + 1, 9).setValue(data.resolutionNotes);
        iSheet.getRange(i + 1, 10).setValue(now);
        iSheet.getRange(i + 1, 11).setValue(user);
        break;
      }
    }

    // 2. Check for any OTHER open tickets, specifically looking for critical OOS ones
    let hasOtherOos = false;
    for (let i = 1; i < iData.length; i++) {
      const rowRentalId = String(iData[i][3]).trim().toLowerCase();
      const dateResolved = String(iData[i][9]).trim();
      const incId = String(iData[i][0]).trim();
      const isOosTriggered = String(iData[i][6]).trim().toUpperCase() === 'YES';
      const severity = String(iData[i][13]).trim().toUpperCase();

      if (rowRentalId === String(data.rentalId).trim().toLowerCase() && 
          dateResolved === "" && 
          incId !== String(data.incidentId).trim()) {
        
        // If we find even one other ticket that triggers OOS, flag it
        if (isOosTriggered || severity === 'OOS') {
          hasOtherOos = true;
          break; 
        }
      }
    }

    // 3. Update the Vehicle status if there are no more OOS tickets holding it back
    const vData = vSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        const row = i + 1;
        
        if (!hasOtherOos) {
          const currentStatus = String(vData[i][9]).trim().toLowerCase();
          if (currentStatus === 'out of service') {
            vSheet.getRange(row, 10).setValue('Warehouse (Ready)');
          }
          vSheet.getRange(row, 16).clearContent(); // Clear the red OOS notes
        }
        
        vSheet.getRange(row, 17).setValue(user); // Updated By
        vSheet.getRange(row, 18).setValue(now);  // Updated At
        break;
      }
    }

    return { success: true };
  } catch (err) {
    throw new Error('Close ticket failed: ' + err.message);
  }
}

function processAddVehicle(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    const rowData = [
      data.vanName || '', data.rentalId || '', data.owner || '', data.makeModel || '',
      data.vin || '', data.plate || '', data.plateExp || '', data.samsaraSerial || '',
      data.samsaraLink || '', 'Warehouse (Ready)', now, '', '', '', now,
      data.notes || '', user, now, '', data.coast || 'No', data.type || '',
      data.lastOil || '', data.lastOilMiles || '', data.lastTires || '' // Cols V, W, X
    ];
    
    while (rowData.length < 24) rowData.push('');
    vSheet.appendRow(rowData);
    return { success: true };
  } catch (err) { throw new Error('Add vehicle failed: ' + err.message); }
}

function logRoutineMaintenance(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Vehicles');
    const vData = sheet.getDataRange().getValues();
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    let rowIndex = -1;

    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        rowIndex = i + 1; break;
      }
    }
    
    if (rowIndex === -1) throw new Error(`Vehicle ${data.rentalId} not found.`);

    if (data.type === 'oil') {
      sheet.getRange(rowIndex, 22).setValue(data.date); // Col V
      sheet.getRange(rowIndex, 23).setValue(data.miles); // Col W
    } else if (data.type === 'tires') {
      sheet.getRange(rowIndex, 24).setValue(data.date); // Col X
    }

    sheet.getRange(rowIndex, 17).setValue(user); // Updated By
    sheet.getRange(rowIndex, 18).setValue(now);  // Updated At

    return { success: true };
  } catch (err) {
    throw new Error('Failed to log maintenance: ' + err.message);
  }
}

function decommissionVehicle(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    
    const vData = vSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        const row = i + 1;
        vSheet.getRange(row, 10).setValue('Decommissioned'); 
        vSheet.getRange(row, 13).clearContent(); 
        vSheet.getRange(row, 14).clearContent(); 
        
        // Write the retirement checklist to Column Y (Index 25)
        const tz = Session.getScriptTimeZone();
        const dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
        const decomNotes = `RETIRED (${dateStr}) - Reason: ${data.reason} | Samsara Removed: ${data.samRemoved ? 'Yes' : 'No'} | Assets Removed: ${data.assetsRemoved ? 'Yes' : 'No'} | Ops Notified: ${data.opsNotified ? 'Yes' : 'No'} | Cleaned: ${data.clean ? 'Yes' : 'No'}`;
        
        const existingNotes = vSheet.getRange(row, 25).getValue();
        vSheet.getRange(row, 25).setValue((existingNotes ? existingNotes + '\n' : '') + decomNotes);
        
        vSheet.getRange(row, 17).setValue(user);
        vSheet.getRange(row, 18).setValue(now);
        break;
      }
    }
    return { success: true };
  } catch (err) {
    throw new Error('Decommission failed: ' + err.message);
  }
}

function updateVehicleDetails(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Vehicles');
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === String(formData.rentalId).trim().toLowerCase()) {
        rowIndex = i + 1; 
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error(`Vehicle ${formData.rentalId} not found.`);
    }

    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    // Map strictly to the 24-column layout 
    // (Note: getRange uses 1-based indexing, so Column P is 16)
    if (formData.vanName !== undefined) sheet.getRange(rowIndex, 1).setValue(formData.vanName);
    if (formData.owner !== undefined) sheet.getRange(rowIndex, 3).setValue(formData.owner);
    if (formData.makeModel !== undefined) sheet.getRange(rowIndex, 4).setValue(formData.makeModel);
    if (formData.vin !== undefined) sheet.getRange(rowIndex, 5).setValue(formData.vin);
    if (formData.licensePlate !== undefined) sheet.getRange(rowIndex, 6).setValue(formData.licensePlate);
    if (formData.plateExpDate !== undefined) sheet.getRange(rowIndex, 7).setValue(formData.plateExpDate);
    if (formData.samsaraSerial !== undefined) sheet.getRange(rowIndex, 8).setValue(formData.samsaraSerial);
    if (formData.samsaraLink !== undefined) sheet.getRange(rowIndex, 9).setValue(formData.samsaraLink);
    
    // Notes/Addons moved to Col P (16)
    if (formData.addons !== undefined) sheet.getRange(rowIndex, 16).setValue(formData.addons);
    
    sheet.getRange(rowIndex, 17).setValue(user); 
    sheet.getRange(rowIndex, 18).setValue(now); 
    
    if (formData.coast !== undefined) sheet.getRange(rowIndex, 20).setValue(formData.coast);
    if (formData.type !== undefined) sheet.getRange(rowIndex, 21).setValue(formData.type);

    return { success: true, message: `Vehicle ${formData.rentalId} updated successfully.` };
  } catch (err) {
    throw new Error('Failed to update vehicle details: ' + err.message);
  }
}

function unretireVehicle(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vSheet = ss.getSheetByName('Vehicles');
    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';

    const vData = vSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) {
        const row = i + 1;
        
        // 1. Reactivate Status
        vSheet.getRange(row, 10).setValue('Warehouse (Ready)');
        
        // 2. Apply Updated Specs
        if (data.vanName) vSheet.getRange(row, 1).setValue(data.vanName);
        if (data.owner) vSheet.getRange(row, 3).setValue(data.owner);
        if (data.makeModel) vSheet.getRange(row, 4).setValue(data.makeModel);
        if (data.type) vSheet.getRange(row, 21).setValue(data.type);
        if (data.vin) vSheet.getRange(row, 5).setValue(data.vin);
        if (data.licensePlate) vSheet.getRange(row, 6).setValue(data.licensePlate);
        if (data.plateExpDate) vSheet.getRange(row, 7).setValue(data.plateExpDate);
        if (data.samsaraSerial !== undefined) vSheet.getRange(row, 8).setValue(data.samsaraSerial);
        if (data.samsaraLink !== undefined) vSheet.getRange(row, 9).setValue(data.samsaraLink);
        if (data.coast) vSheet.getRange(row, 20).setValue(data.coast);
        if (data.addons) vSheet.getRange(row, 16).setValue(data.addons);
        
        // 3. Append to Lifecycle History (Column Y / Index 25)
        const tz = Session.getScriptTimeZone();
        const dateStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
        const unretireLog = `UN-RETIRED (${dateStr}) - Reason: ${data.reason}`;
        
        const existingNotes = vSheet.getRange(row, 25).getValue();
        vSheet.getRange(row, 25).setValue((existingNotes ? existingNotes + '\n' : '') + unretireLog);
        
        // 4. Stamp Audit 
        vSheet.getRange(row, 17).setValue(user);
        vSheet.getRange(row, 18).setValue(now);
        break;
      }
    }
    return { success: true };
  } catch (err) {
    throw new Error('Un-retire failed: ' + err.message);
  }
}

function processEditTech(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Tech_Profiles');
    const tData = sheet.getDataRange().getValues();
    let row = -1;

    for (let i = 1; i < tData.length; i++) {
      if (String(tData[i][0]).trim().toLowerCase() === String(data.origName).trim().toLowerCase()) {
        row = i + 1; break;
      }
    }

    if (row === -1) {
      row = sheet.getLastRow() + 1; // Add new tech if not found
    }

    sheet.getRange(row, 1).setValue(data.name);
    sheet.getRange(row, 2).setValue(data.phone);

    sheet.getRange(row, 3).setValue(data.shift);
    sheet.getRange(row, 4).setValue(data.start ? "'" + data.start : ""); // Forces plain text
    sheet.getRange(row, 5).setValue(data.end ? "'" + data.end : "");     // Forces plain text
    sheet.getRange(row, 6).setValue(data.workDays);
    
    sheet.getRange(row, 7).setValue(data.passedTest ? 'Yes' : 'No');
    sheet.getRange(row, 8).setValue(data.goFast ? 'Yes' : 'No');
    sheet.getRange(row, 9).setValue(data.restricted ? 'Yes' : 'No');
    sheet.getRange(row, 10).setValue(data.status);
    sheet.getRange(row, 11).setValue(data.notes);
    sheet.getRange(row, 12).setValue(data.hrStatus);
    sheet.getRange(row, 13).setValue(data.coast ? 'Yes' : 'No');
    sheet.getRange(row, 14).setValue(data.evgo ? 'Yes' : 'No');
    sheet.getRange(row, 15).setValue(data.trainer ? 'Yes' : 'No');
    sheet.getRange(row, 16).setValue(data.samsaraViolation ? 'Yes' : 'No');
    sheet.getRange(row, 17).setValue(data.flagged);

    return { success: true };
  } catch (err) { throw new Error('Tech update failed: ' + err.message); }
}