/**
 * Opens the Check-In / Return to Service dialog.
 */
function showCheckInModal() {
  const oosVehicles = getOosVehiclesWithIncidents();
  const template = HtmlService.createTemplateFromFile('Modal_CheckIn');
  template.oosVehicles = oosVehicles;

  const html = template.evaluate()
    .setWidth(540)
    .setHeight(580);
  SpreadsheetApp.getUi().showModalDialog(html, '✅ Check Vehicle Back Into Service');
}

/**
 * Fetches Out-of-Service vehicles with their open incident details.
 */
function getOosVehiclesWithIncidents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vSheet = ss.getSheetByName(VEHICLES_SHEET);
  const incSheet = ss.getSheetByName(INCIDENTS_SHEET);
  if (!vSheet || !incSheet) return [];

  const vData = vSheet.getDataRange().getValues();
  const incData = incSheet.getDataRange().getValues();

  // Find all Out of Service vehicles
  const list = [];
  for (let i = 1; i < vData.length; i++) {
    const status = String(vData[i][6]).trim(); // Col G: Current Status
    if (status === 'Out of Service') {
      const cleanName = vData[i][0];
      const rentalId = vData[i][1];

      // Find the most recent open incident for this vehicle
      let latestIncident = null;
      for (let j = incData.length - 1; j >= 1; j--) {
        const incRental = String(incData[j][3]).trim();
        const incClean = String(incData[j][2]).trim();
        const resolvedDate = incData[j][9]; // Col J: Date Resolved

        if ((incRental === String(rentalId).trim() || incClean === String(cleanName).trim()) && !resolvedDate) {
          latestIncident = {
            incidentId: incData[j][0],
            dateReported: Utilities.formatDate(new Date(incData[j][1]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            issueCategory: incData[j][4],
            description: incData[j][5],
            statusImpact: incData[j][7]
          };
          break;
        }
      }

      list.push({
        cleanName: cleanName,
        rentalId: rentalId,
        displayName: `${cleanName} (${rentalId})`,
        incident: latestIncident
      });
    }
  }
  return list;
}

/**
 * Processes vehicle return to active service.
 */
function processCheckInVehicle(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vSheet = ss.getSheetByName(VEHICLES_SHEET);
  const incSheet = ss.getSheetByName(INCIDENTS_SHEET);
  if (!vSheet || !incSheet) throw new Error('Operational sheets not found');

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  // 1. Update Vehicle Status to Active
  const vData = vSheet.getDataRange().getValues();
  let vehicleFound = false;
  for (let i = 1; i < vData.length; i++) {
    if (String(vData[i][1]).trim() === String(data.rentalId).trim() ||
        String(vData[i][0]).trim() === String(data.cleanName).trim()) {
      const row = i + 1;
      vSheet.getRange(row, 7).setValue('Active'); // Col G: Status back to Active
      vSheet.getRange(row, 12).setValue(now);     // Col L: Last Service Date
      const curNotes = vSheet.getRange(row, 13).getValue();
      vSheet.getRange(row, 13).setValue((curNotes ? curNotes + ' | ' : '') + `[Returned to Service: ${data.resolutionNotes}]`);
      vSheet.getRange(row, 14).setValue(user);    // Col N: Last Updated By
      vSheet.getRange(row, 15).setValue(now);     // Col O: Last Updated At
      vehicleFound = true;
      break;
    }
  }

  if (!vehicleFound) throw new Error(`Vehicle ${data.cleanName} not found.`);

  // 2. Close out Incident in Incidents_Service if linked
  if (data.incidentId) {
    const incData = incSheet.getDataRange().getValues();
    for (let j = 1; j < incData.length; j++) {
      if (String(incData[j][0]).trim() === String(data.incidentId).trim()) {
        const row = j + 1;
        incSheet.getRange(row, 9).setValue(data.resolutionNotes); // Col I: Resolution Notes
        incSheet.getRange(row, 10).setValue(now);                 // Col J: Date Resolved
        incSheet.getRange(row, 11).setValue(user);                // Col K: Resolved By
        break;
      }
    }
  }

  // 3. Log to Audit_Log
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