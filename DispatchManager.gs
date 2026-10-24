/**
 * @fileoverview Daily Shift Dispatch and Driver Checkout Engine.
 */

function processDispatchCheckout(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deploySheet = ss.getSheetByName(DEPLOYMENTS_SHEET);
    const vehicleSheet = ss.getSheetByName(VEHICLES_SHEET);
    if (!deploySheet || !vehicleSheet) throw new Error('Required sheets not found');

    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    const deployId = 'DEP-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const dispatchDate = data.dispatchDate ? new Date(data.dispatchDate) : now;

    const fullGoals = data.focusCategory 
      ? (data.goals ? `[${data.focusCategory}] ${data.goals}` : `[${data.focusCategory}]`)
      : (data.goals || '');

    // 1. Append record to Deployments
    // Headers: [Deployment ID, Date, Shift, Clean Name, Rental ID, Assigned Tech, Work Zone, Shift Focus / Goals, Battery Target / Manifest, Status, Departure Time, Return Time, Logged By, Logged Timestamp]
    deploySheet.appendRow([
      deployId,
      dispatchDate,
      data.shift,
      data.cleanName,
      data.rentalId,
      data.assignedTech,
      data.workZone,
      fullGoals,
      data.batteryTarget || '',
      'Dispatched',
      now, // Departure Time
      '',  // Return Time
      user,
      now
    ]);

    // 2. Update Vehicle row in Vehicles sheet (17 cols)
    const vData = vehicleSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      const vClean = String(vData[i][0]).trim().toLowerCase();
      const vRental = String(vData[i][1]).trim().toLowerCase();
      const tClean = String(data.cleanName).trim().toLowerCase();
      const tRental = String(data.rentalId).trim().toLowerCase();

      if ((tRental && vRental === tRental) || (tClean && vClean === tClean)) {
        const row = i + 1;
        vehicleSheet.getRange(row, 9).setValue('In Service');     // Col I: Current Status
        vehicleSheet.getRange(row, 12).setValue(data.workZone);   // Col L: Current Assigned Zone
        vehicleSheet.getRange(row, 13).setValue(data.assignedTech);// Col M: Current Assigned Tech
        vehicleSheet.getRange(row, 16).setValue(user);            // Col P: Last Updated By
        vehicleSheet.getRange(row, 17).setValue(now);             // Col Q: Last Updated At
        break;
      }
    }

    logAudit('DISPATCH_CHECKOUT', 'Deployments', deployId, {
      vehicle: data.cleanName,
      rentalId: data.rentalId,
      tech: data.assignedTech,
      shift: data.shift,
      zone: data.workZone,
      batteryTarget: data.batteryTarget,
      dispatchedBy: user
    });

    return {
      success: true,
      message: `Deployment ${deployId} logged!${data.cleanName} dispatched to ${data.assignedTech}.` }; } catch (err) { Logger.log('processDispatchCheckout error: ' + err.toString()); throw new Error(err.message); } }  function updateDeploymentRecord(data) { const ss = SpreadsheetApp.getActiveSpreadsheet(); const dSheet = ss.getSheetByName(DEPLOYMENTS_SHEET); const vSheet = ss.getSheetByName(VEHICLES_SHEET); if (!dSheet \vert{}\vert{} !vSheet) throw new Error('Sheets not found');  const rows = dSheet.getDataRange().getValues(); let rowIndex = -1;  for (let i = 1; i < rows.length; i++) { if (String(rows[i][0]).trim() === String(data.id).trim()) { rowIndex = i + 1; break; } }  if (rowIndex === -1) throw new Error('Deployment ID ' + data.id + ' not found.');  const now = new Date(); const user = Session.getActiveUser().getEmail() \vert{}\vert{} 'local-user';  if (data.shift) dSheet.getRange(rowIndex, 3).setValue(data.shift); if (data.tech) dSheet.getRange(rowIndex, 6).setValue(data.tech); if (data.zone) dSheet.getRange(rowIndex, 7).setValue(data.zone); if (data.goals !== undefined) dSheet.getRange(rowIndex, 8).setValue(data.goals); if (data.target !== undefined) dSheet.getRange(rowIndex, 9).setValue(data.target);  // Sync to Vehicles sheet (Cols L and M) const vRows = vSheet.getDataRange().getValues(); for (let j = 1; j < vRows.length; j++) { if (String(vRows[j][0]).trim().toLowerCase() === String(data.vehicle).trim().toLowerCase() \vert{}\vert{} String(vRows[j][1]).trim().toLowerCase() === String(data.rentalId).trim().toLowerCase()) { const vRowIndex = j + 1; if (data.zone) vSheet.getRange(vRowIndex, 12).setValue(data.zone); if (data.tech) vSheet.getRange(vRowIndex, 13).setValue(data.tech); vSheet.getRange(vRowIndex, 16).setValue(user); vSheet.getRange(vRowIndex, 17).setValue(now); break; } }  logAudit('DISPATCH_UPDATED', 'Deployments', data.id, { vehicle: data.vehicle, tech: data.tech, zone: data.zone, shift: data.shift, updatedBy: user });  return { success: true, message: `Deployment ${data.id} updated successfully!` };
}