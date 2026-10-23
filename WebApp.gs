/**
 * @fileoverview Serves the standalone Web App dashboard and web-side endpoints.
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard_App')
    .evaluate()
    .setTitle('VEO Fleet & Dispatch Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Returns full state for the web app in one synchronized payload.
 */
function getDashboardState(dateStr, shift) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const vSheet = ss.getSheetByName(VEHICLES_SHEET);
  const dSheet = ss.getSheetByName(DEPLOYMENTS_SHEET);
  const iSheet = ss.getSheetByName(INCIDENTS_SHEET);
  const config = getSystemConfig();

  // 1. Vehicle Data & Metrics
  const vehicles = [];
  let counts = { total: 0, active: 0, oos: 0, decommissioned: 0 };

  if (vSheet && vSheet.getLastRow() > 1) {
    const vRows = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, 15).getValues();
    vRows.forEach(r => {
      const cleanName = r[0];
      const rentalId = r[1];
      const status = r[6] || 'Active';
      if (!cleanName && !rentalId) return;

      counts.total++;
      if (status === 'Active' || status === 'In Service') counts.active++;
      else if (status === 'Out of Service') counts.oos++;
      else if (status === 'Decommissioned') counts.decommissioned++;

      vehicles.push({
        cleanName: cleanName,
        rentalId: rentalId,
        agency: r[2],
        makeModel: r[3],
        vin: r[4],
        plate: r[5],
        status: status,
        pickupDate: r[7] ? Utilities.formatDate(new Date(r[7]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '—',
        serviceNum: r[8],
        zone: r[9] || 'Unassigned',
        driver: r[10] || 'Unassigned',
        lastService: r[11] ? Utilities.formatDate(new Date(r[11]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '—',
        notes: r[12]
      });
    });
  }

  // 2. Filtered Deployments for selected date and shift
  const targetDateStr = dateStr || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const targetShift = shift || 'Shift 1';
  const deployments = [];

  if (dSheet && dSheet.getLastRow() > 1) {
    const dRows = dSheet.getRange(2, 1, dSheet.getLastRow() - 1, 14).getValues();
    dRows.forEach(r => {
      const depDate = r[1] ? Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
      const depShift = String(r[2]).trim();

      if (depDate === targetDateStr && depShift.toLowerCase() === targetShift.toLowerCase()) {
        deployments.push({
          id: r[0],
          date: depDate,
          shift: depShift,
          vehicle: r[3],
          rentalId: r[4],
          tech: r[5],
          zone: r[6],
          goals: r[7],
          target: r[8],
          status: r[9],
          departureTime: r[10] ? Utilities.formatDate(new Date(r[10]), Session.getScriptTimeZone(), 'hh:mm a') : '—'
        });
      }
    });
  }

  // 3. Incidents
  const incidents = [];
  if (iSheet && iSheet.getLastRow() > 1) {
    const iRows = iSheet.getRange(2, 1, iSheet.getLastRow() - 1, 13).getValues();
    iRows.forEach(r => {
      incidents.push({
        id: r[0],
        date: r[1] ? Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        vehicle: r[2],
        rentalId: r[3],
        category: r[4],
        details: r[5],
        oos: r[6],
        impact: r[7],
        resolution: r[8],
        dateResolved: r[9] ? Utilities.formatDate(new Date(r[9]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        resolvedBy: r[10]
      });
    });
  }

  // 4. Out-of-Service Vehicles with open incident link
  const oosList = [];
  vehicles.forEach(v => {
    if (v.status === 'Out of Service') {
      let openInc = null;
      for (let j = incidents.length - 1; j >= 0; j--) {
        if ((incidents[j].vehicle === v.cleanName || incidents[j].rentalId === v.rentalId) && !incidents[j].dateResolved) {
          openInc = incidents[j];
          break;
        }
      }
      oosList.push({
        cleanName: v.cleanName,
        rentalId: v.rentalId,
        displayName: `${v.cleanName} (${v.rentalId})`,
        incident: openInc
      });
    }
  });

  return {
    config: config,
    counts: counts,
    vehicles: vehicles,
    deployments: deployments,
    incidents: incidents,
    oosVehicles: oosList,
    selectedDate: targetDateStr,
    selectedShift: targetShift
  };
}