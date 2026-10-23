/**
 * @fileoverview Incident & Maintenance Reporting Engine.
 */

function showIncidentModal() {
  try {
    const config = getSystemConfig();
    const vehicles = getActiveVehiclesList();

    const template = HtmlService.createTemplateFromFile('Modal_Incident');
    template.vehicles = vehicles || [];
    template.issueCategories = (config.issueCategories && config.issueCategories.length) ? config.issueCategories : [
      'Mechanical / Engine',
      'Flat Tire / Wheel',
      'Check Engine Light',
      'Body / Exterior Damage',
      'Scheduled Maintenance',
      'Battery / Electrical',
      'Safety Inspection Failure',
      'Samsara / Tech Equipment'
    ];
    template.statusImpacts = (config.statusImpacts && config.statusImpacts.length) ? config.statusImpacts : [
      'Out of Service - Needs Repair',
      'Out of Service - On Hold',
      'Out of Service - Dealer Return',
      'In Service - Roadworthy',
      'In Service - Monitor'
    ];

    const html = template.evaluate()
      .setWidth(560)
      .setHeight(640);
    SpreadsheetApp.getUi().showModalDialog(html, '⚠️ Report Incident / Vehicle Service');
  } catch (err) {
    SpreadsheetApp.getUi().alert('Unable to open Report Incident modal: ' + err.message + '\nMake sure Modal_Incident.html exists in your Apps Script files.');
  }
}