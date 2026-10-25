/**
 * @fileoverview Incident & Maintenance Reporting Engine.
 */

function processReportIncident(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const incidentSheet = ss.getSheetByName(INCIDENTS_SHEET);
    const vehicleSheet = ss.getSheetByName(VEHICLES_SHEET);
    if (!incidentSheet || !vehicleSheet) throw new Error('Operational sheets not found');

    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'local-user';
    const incidentId = 'INC-' + Utilities.getUuid().substring(0, 8).toUpperCase();

    // 1. Append Incident Row
    incidentSheet.appendRow([
      incidentId,
      now,
      data.cleanName || '',
      data.rentalId || '',
      data.issueCategory || '',
      data.description || '',
      data.triggerOos ? 'YES' : 'NO',
      data.statusImpact || '',
      '', // Resolution Notes
      '', // Date Resolved
      '', // Resolved By
      user,
      now
    ]);

    // 2. Automate Vehicle Status Update
    let newStatus = 'In Service';
    if (data.triggerOos || (data.statusImpact && data.statusImpact.indexOf('Out of Service') !== -1)) {
      newStatus = 'Out of Service';
    } else if (data.statusImpact && data.statusImpact.indexOf('In Service') !== -1) {
      newStatus = 'In Service';
    }

    const vData = vehicleSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      const vClean = String(vData[i][0]).trim().toLowerCase();
      const vRental = String(vData[i][1]).trim().toLowerCase();
      const tClean = String(data.cleanName).trim().toLowerCase();
      const tRental = String(data.rentalId).trim().toLowerCase();

      if ((tRental && vRental === tRental) || (tClean && vClean === tClean)) {
        const row = i + 1;
        vehicleSheet.getRange(row, 9).setValue(newStatus); // Col I: Status
        const curNotes = vehicleSheet.getRange(row, 15).getValue();
        vehicleSheet.getRange(row, 15).setValue((curNotes ? curNotes + ' | ' : '') + `[${incidentId}:${data.issueCategory} -> ${data.statusImpact}]`); vehicleSheet.getRange(row, 16).setValue(user); vehicleSheet.getRange(row, 17).setValue(now); break; } }  logAudit('INCIDENT_REPORTED', 'Incidents_Service', incidentId, { vehicle: data.cleanName, rentalId: data.rentalId, category: data.issueCategory, impact: data.statusImpact, oosTriggered: data.triggerOos, vehicleStatusSetTo: newStatus, reportedBy: user });  return { success: true, message: `Incident ${incidentId} logged successfully! Vehicle status set to "${newStatus}".`
    };
  } catch (err) {
    Logger.log('processReportIncident error: ' + err.toString());
    throw new Error(err.message);
  }
}

function processCheckInVehicle(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vSheet = ss.getSheetByName(VEHICLES_SHEET);
  const incSheet = ss.getSheetByName(INCIDENTS_SHEET);
  if (!vSheet || !incSheet) throw new Error('Operational sheets not found');

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  const vData = vSheet.getDataRange().getValues();
  let vehicleFound = false;

  for (let i = 1; i < vData.length; i++) {
    if (String(vData[i][1]).trim() === String(data.rentalId).trim() ||
        String(vData[i][0]).trim() === String(data.cleanName).trim()) {
      const row = i + 1;
      vSheet.getRange(row, 9).setValue('Warehouse (Ready)');  // Col I: back to Active
      vSheet.getRange(row, 14).setValue(now);     // Col N: Last Service Date
      const curNotes = vSheet.getRange(row, 15).getValue();
      vSheet.getRange(row, 15).setValue((curNotes ? curNotes + ' | ' : '') + `[Returned to Service: ${data.resolutionNotes}]`); vSheet.getRange(row, 16).setValue(user); vSheet.getRange(row, 17).setValue(now); vehicleFound = true; break; } }  if (!vehicleFound) throw new Error(`Vehicle ${data.cleanName} not found.`);

  if (data.incidentId) {
    const incData = incSheet.getDataRange().getValues();
    for (let j = 1; j < incData.length; j++) {
      if (String(incData[j][0]).trim() === String(data.incidentId).trim()) {
        const row = j + 1;
        incSheet.getRange(row, 9).setValue(data.resolutionNotes);
        incSheet.getRange(row, 10).setValue(now);
        incSheet.getRange(row, 11).setValue(user);
        break;
      }
    }
  }

  logAudit('VEHICLE_CHECKED_IN', 'Vehicles', data.rentalId || data.cleanName, {
    vehicle: data.cleanName,
    rentalId: data.rentalId,
    closedIncidentId: data.incidentId || 'None',
    resolutionNotes: data.resolutionNotes,
    servicedBy: user
  });

  return {
    success: true,
    message: `${data.cleanName} checked back into service and marked Active!`
  };
}