/**
 * @fileoverview Incident & Maintenance Reporting Engine.
 * Handles issue logging, status impact, and automated Out-of-Service transitions.
 */

/**
 * Opens the Report Incident / Service dialog.
 */
function showIncidentModal() {
  const config = getSystemConfig();
  const vehicles = getActiveVehiclesList();

  const template = HtmlService.createTemplateFromFile('Modal_Incident');
  template.vehicles = vehicles;
  template.oosReasons = config.oosReasons;

  const html = template.evaluate()
    .setWidth(540)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '⚠️ Report Incident / Vehicle Service');
}

/**
 * Backend handler for incident submissions.
 * Appends record to Incidents_Service and updates vehicle status if OOS triggered.
 */
function processReportIncident(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const incidentSheet = ss.getSheetByName(INCIDENTS_SHEET);
  const vehicleSheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!incidentSheet || !vehicleSheet) throw new Error('Operational sheets not found');

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';
  const incidentId = 'INC-' + Utilities.getUuid().substring(0, 8).toUpperCase();

  // 1. Log incident into Incidents_Service
  // Headers: [Incident ID, Date Reported, Vehicle Clean Name, Rental ID, Issue Category, Description, OOS Triggered, Status Impact, Resolution Notes, Date Resolved, Resolved By, Logged By, Created Timestamp]
  incidentSheet.appendRow([
    incidentId,
    now,
    data.cleanName,
    data.rentalId,
    data.issueCategory,
    data.description,
    data.triggerOos ? 'YES' : 'NO',
    data.statusImpact || 'Under Review',
    '',   // Resolution Notes (empty until resolved)
    '',   // Date Resolved
    '',   // Resolved By
    user,
    now
  ]);

  // 2. Automate Out-of-Service flip if checked
  let statusUpdated = false;
  if (data.triggerOos) {
    const vData = vehicleSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      if (String(vData[i][1]).trim() === String(data.rentalId).trim() ||
          String(vData[i][0]).trim() === String(data.cleanName).trim()) {
        const row = i + 1;
        vehicleSheet.getRange(row, 7).setValue('Out of Service'); // Col G: Current Status
        const curNotes = vehicleSheet.getRange(row, 13).getValue();
        vehicleSheet.getRange(row, 13).setValue((curNotes ? curNotes + ' | ' : '') + `[OOS: ${data.issueCategory} - ${incidentId}]`);
        vehicleSheet.getRange(row, 14).setValue(user);
        vehicleSheet.getRange(row, 15).setValue(now);
        statusUpdated = true;
        break;
      }
    }
  }

  // 3. Record in Audit Log
  logAudit('INCIDENT_REPORTED', 'Incidents_Service', incidentId, {
    vehicle: data.cleanName,
    rentalId: data.rentalId,
    issue: data.issueCategory,
    oosTriggered: data.triggerOos,
    impact: data.statusImpact,
    reportedBy: user
  });

  return {
    success: true,
    message: `Incident ${incidentId} recorded. ${statusUpdated ? 'Vehicle marked Out of Service.' : ''}`
  };
}