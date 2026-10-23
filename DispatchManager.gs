/**
 * @fileoverview Daily Shift Dispatch and Driver Checkout Engine.
 */

/**
 * Opens the Daily Dispatch Checkout Modal.
 */
function showDispatchModal() {
  try {
    const config = getSystemConfig();
    const availableVehicles = getDispatchableVehicles();

    const template = HtmlService.createTemplateFromFile('Modal_Dispatch');
    template.vehicles = availableVehicles;
    template.shifts = config.shifts || ['Shift 1', 'Shift 2', 'Shift 3'];
    template.zones = config.zones || [];
    template.focusOptions = config.focusOptions || [];

    const html = template.evaluate()
      .setWidth(580)
      .setHeight(680);
    SpreadsheetApp.getUi().showModalDialog(html, '📋 Daily Driver & Shift Dispatch Checkout');
  } catch (err) {
    SpreadsheetApp.getUi().alert('Unable to open Dispatch modal: ' + err.message);
  }
}

/**
 * Returns only vehicles that are Active or In Service (safe for road deployment).
 */
function getDispatchableVehicles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  // Read: Col A (Clean Name), Col B (Rental ID), Col G (Status)
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const list = [];

  for (let i = 0; i < data.length; i++) {
    const cleanName = String(data[i][0]).trim();
    const rentalId = String(data[i][1]).trim();
    const status = String(data[i][6]).trim();

    if ((status === 'Active' || status === 'In Service') && (cleanName || rentalId)) {
      list.push({
        cleanName: cleanName,
        rentalId: rentalId,
        status: status,
        displayName: `${cleanName} (${rentalId}) - [${status}]`
      });
    }
  }
  return list;
}

/**
 * Backend processor for dispatch checkout submission.
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

    // Combine focus dropdown and custom goals text
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
      '',  // Return Time (empty until check-in/shift end)
      user,
      now
    ]);

    // 2. Update Vehicle row in Vehicles sheet with currently assigned Tech and Zone
    const vData = vehicleSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      const vClean = String(vData[i][0]).trim().toLowerCase();
      const vRental = String(vData[i][1]).trim().toLowerCase();
      const tClean = String(data.cleanName).trim().toLowerCase();
      const tRental = String(data.rentalId).trim().toLowerCase();

      if ((tRental && vRental === tRental) || (tClean && vClean === tClean)) {
        const row = i + 1;
        vehicleSheet.getRange(row, 7).setValue('In Service');     // Col G: Set to In Service
        vehicleSheet.getRange(row, 10).setValue(data.workZone);   // Col J: Current Assigned Zone
        vehicleSheet.getRange(row, 11).setValue(data.assignedTech);// Col K: Current Assigned Tech
        vehicleSheet.getRange(row, 14).setValue(user);            // Col N: Last Updated By
        vehicleSheet.getRange(row, 15).setValue(now);             // Col O: Last Updated At
        break;
      }
    }

    // 3. Centralized Audit Log
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
      message: `Deployment ${deployId} logged! ${data.cleanName} dispatched to ${data.assignedTech} in Zone ${data.workZone}.`
    };
  } catch (err) {
    Logger.log('processDispatchCheckout error: ' + err.toString());
    throw new Error(err.message);
  }
}