/**
 * @fileoverview Incident & Maintenance Reporting Engine.
 */



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