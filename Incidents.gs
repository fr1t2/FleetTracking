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
    SpreadsheetApp.getUi().alert('Unable to open Report Incident modal: ' + err.message);
  }
}

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
    // [Incident ID, Date Reported, Clean Name, Rental ID, Issue Category, Description, OOS Triggered, Status Impact, Resolution Notes, Date Resolved, Resolved By, Logged By, Created Timestamp]
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
    let updated = false;

    for (let i = 1; i < vData.length; i++) {
      const vClean = String(vData[i][0]).trim().toLowerCase();
      const vRental = String(vData[i][1]).trim().toLowerCase();
      const tClean = String(data.cleanName).trim().toLowerCase();
      const tRental = String(data.rentalId).trim().toLowerCase();

      if ((tRental && vRental === tRental) || (tClean && vClean === tClean)) {
        const row = i + 1;
        vehicleSheet.getRange(row, 7).setValue(newStatus); // Col G: Current Status
        const curNotes = vehicleSheet.getRange(row, 13).getValue();
        vehicleSheet.getRange(row, 13).setValue((curNotes ? curNotes + ' | ' : '') + `[${incidentId}: ${data.issueCategory} -> ${data.statusImpact}]`);
        vehicleSheet.getRange(row, 14).setValue(user);
        vehicleSheet.getRange(row, 15).setValue(now);
        updated = true;
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
      vehicleStatusSetTo: newStatus,
      reportedBy: user
    });

    return {
      success: true,
      message: `Incident ${incidentId} logged successfully! Vehicle status set to "${newStatus}".`
    };
  } catch (err) {
    Logger.log('processReportIncident error: ' + err.toString());
    throw new Error(err.message);
  }
}

function showCheckInModal() {
  try {
    const oosVehicles = getOosVehiclesWithIncidents();
    const template = HtmlService.createTemplateFromFile('Modal_CheckIn');
    template.oosVehicles = oosVehicles;

    const html = template.evaluate()
      .setWidth(540)
      .setHeight(580);
    SpreadsheetApp.getUi().showModalDialog(html, '✅ Check Vehicle Back Into Service');
  } catch (err) {
    SpreadsheetApp.getUi().alert('Unable to open Check In modal: ' + err.message);
  }
}

function getOosVehiclesWithIncidents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vSheet = ss.getSheetByName(VEHICLES_SHEET);
  const incSheet = ss.getSheetByName(INCIDENTS_SHEET);
  if (!vSheet || !incSheet) return [];

  const vData = vSheet.getDataRange().getValues();
  const incData = incSheet.getDataRange().getValues();

  const list = [];
  for (let i = 1; i < vData.length; i++) {
    const status = String(vData[i][6]).trim(); // Col G: Current Status
    if (status === 'Out of Service') {
      const cleanName = vData[i][0];
      const rentalId = vData[i][1];

      let latestIncident = null;
      for (let j = incData.length - 1; j >= 1; j--) {
        const incRental = String(incData[j][3]).trim();
        const incClean = String(incData[j][2]).trim();
        const resolvedDate = incData[j][9];

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
      vSheet.getRange(row, 7).setValue('Active'); // Col G: back to Active
      vSheet.getRange(row, 12).setValue(now);     // Col L: Last Service Date
      const curNotes = vSheet.getRange(row, 13).getValue();
      vSheet.getRange(row, 13).setValue((curNotes ? curNotes + ' | ' : '') + `[Returned to Service: ${data.resolutionNotes}]`);
      vSheet.getRange(row, 14).setValue(user);
      vSheet.getRange(row, 15).setValue(now);
      vehicleFound = true;
      break;
    }
  }

  if (!vehicleFound) throw new Error(`Vehicle ${data.cleanName} not found.`);

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