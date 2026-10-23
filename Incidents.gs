/**
 * @fileoverview Incident & Maintenance Reporting Engine.
 */

function showIncidentModal() {
  const config = getSystemConfig();
  const vehicles = getActiveVehiclesList();

  const template = HtmlService.createTemplateFromFile('Modal_Incident');
  template.vehicles = vehicles;
  template.issueCategories = config.issueCategories;
  template.statusImpacts = config.statusImpacts;

  const html = template.evaluate()
    .setWidth(560)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, '⚠️ Report Incident / Vehicle Service');
}

function processReportIncident(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const incidentSheet = ss.getSheetByName(INCIDENTS_SHEET);
  const vehicleSheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!incidentSheet || !vehicleSheet) throw new Error('Operational sheets not found');

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';
  const incidentId = 'INC-' + Utilities.getUuid().substring(0, 8).toUpperCase();

  // 1. Append Incident Row
  // [Incident ID, Date Reported, Vehicle Clean Name, Rental ID, Issue Category, Description, OOS Triggered, Status Impact, Resolution Notes, Date Resolved, Resolved By, Logged By, Created Timestamp]
  incidentSheet.appendRow([
    incidentId,
    now,
    data.cleanName,
    data.rentalId,
    data.issueCategory,
    data.description,
    data.triggerOos ? 'YES' : 'NO',
    data.statusImpact,
    '', // Resolution Notes
    '', // Date Resolved
    '', // Resolved By
    user,
    now
  ]);

  // 2. Automate Vehicle Status Update
  let newStatus = 'Active';
  if (data.triggerOos || data.statusImpact.indexOf('Out of Service') !== -1) {
    newStatus = 'Out of Service';
  } else if (data.statusImpact.indexOf('In Service') !== -1) {
    newStatus = 'In Service';
  }

  const vData = vehicleSheet.getDataRange().getValues();
  for (let i = 1; i < vData.length; i++) {
    if (String(vData[i][1]).trim() === String(data.rentalId).trim() ||
        String(vData[i][0]).trim() === String(data.cleanName).trim()) {
      const row = i + 1;
      vehicleSheet.getRange(row, 7).setValue(newStatus); // Col G: Current Status
      const curNotes = vehicleSheet.getRange(row, 13).getValue();
      vehicleSheet.getRange(row, 13).setValue((curNotes ? curNotes + ' | ' : '') + `[${incidentId}: ${data.issueCategory} -> ${data.statusImpact}]`);
      vehicleSheet.getRange(row, 14).setValue(user);
      vehicleSheet.getRange(row, 15).setValue(now);
      break;
    }
  }

  // 3. Log Audit
  logAudit('INCIDENT_REPORTED', 'Incidents_Service', incidentId, {
    vehicle: data.cleanName,
    rentalId: data.rentalId,
    category: data.issueCategory,
    impact: data.statusImpact,
    oosTriggered: data.triggerOos,
    reportedBy: user
  });

  return {
    success: true,
    message: `Incident ${incidentId} recorded. Vehicle status set to ${newStatus}.`
  };
}