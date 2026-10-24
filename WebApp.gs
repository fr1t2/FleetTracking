/**
 * @fileoverview High-Speed Web App Backend
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard_App')
    .evaluate()
    .setTitle('VEO Fleet & Dispatch Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const cfg = getSystemConfig(); // Uses the config we built in Phase 1

  // Vehicles (18 Columns)
  const vSheet = ss.getSheetByName('Vehicles');
  const vehicles = [];
  const counts = { total: 0, ready: 0, dispatched: 0, oos: 0, decommissioned: 0 };

  if (vSheet && vSheet.getLastRow() > 1) {
    const vRows = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, 18).getValues();
    vRows.forEach(r => {
      const cleanName = String(r[0] || '').trim();
      const rentalId = String(r[1] || '').trim();
      const status = String(r[8] || 'Warehouse (Ready)').trim(); // Col I
      if (!cleanName && !rentalId) return;

      if (status !== 'Decommissioned') counts.total++;
      if (status === 'Warehouse (Ready)') counts.ready++;
      else if (status === 'Dispatched') counts.dispatched++;
      else if (status === 'Out of Service') counts.oos++;
      else if (status === 'Decommissioned') counts.decommissioned++;

      vehicles.push({
        cleanName: cleanName,
        rentalId: rentalId,
        agency: String(r[2] || ''),
        makeModel: String(r[3] || ''),
        vin: String(r[4] || ''),
        plate: String(r[5] || ''),
        plateExpDate: (r[6] instanceof Date) ? Utilities.formatDate(r[6], tz, 'yyyy-MM-dd') : String(r[6] || '—'),
        samsaraSerial: String(r[7] || ''),
        status: status,
        zone: String(r[11] || '—'),
        driver: String(r[12] || '—'),
        lastService: (r[13] instanceof Date) ? Utilities.formatDate(r[13], tz, 'yyyy-MM-dd') : '—',
        notes: String(r[14] || ''),
        addons: String(r[17] || '') // Col R
      });
    });
  }

  // Deployments (16 Columns)
  const dSheet = ss.getSheetByName('Deployments');
  const deployments = [];
  if (dSheet && dSheet.getLastRow() > 1) {
    const dRows = dSheet.getRange(2, 1, dSheet.getLastRow() - 1, 16).getValues();
    dRows.forEach(r => {
      deployments.push({
        id: String(r[0] || ''),
        date: (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
        shift: String(r[2] || '').trim(),
        vehicle: String(r[3] || ''),
        rentalId: String(r[4] || ''),
        tech: String(r[5] || ''),
        zone: String(r[6] || ''),
        goals: String(r[7] || ''),
        target: String(r[8] || ''),
        status: String(r[9] || 'Dispatched'),
        departureTime: (r[10] instanceof Date) ? Utilities.formatDate(r[10], tz, 'hh:mm a') : String(r[10]),
        returnTime: (r[11] instanceof Date) ? Utilities.formatDate(r[11], tz, 'hh:mm a') : String(r[11] || ''),
        coDriver: String(r[14] || ''),
        isTraining: String(r[15] || '') === 'Yes'
      });
    });
  }

  // Incidents
  const iSheet = ss.getSheetByName('Incidents_Service');
  const incidents = [];
  if (iSheet && iSheet.getLastRow() > 1) {
    const iRows = iSheet.getRange(2, 1, iSheet.getLastRow() - 1, 13).getValues();
    iRows.forEach(r => {
      incidents.push({
        id: String(r[0] || ''),
        date: (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
        vehicle: String(r[2] || ''),
        rentalId: String(r[3] || ''),
        category: String(r[4] || ''),
        details: String(r[5] || ''),
        oos: String(r[6] || ''),
        impact: String(r[7] || ''),
        resolution: String(r[8] || ''),
        dateResolved: (r[9] instanceof Date) ? Utilities.formatDate(r[9], tz, 'yyyy-MM-dd') : String(r[9])
      });
    });
  }

  const oosList = vehicles.filter(v => v.status === 'Out of Service').map(v => {
    let openInc = incidents.find(i => (i.rentalId === v.rentalId) && !i.dateResolved);
    return { cleanName: v.cleanName, rentalId: v.rentalId, displayName: `${v.cleanName} (${v.rentalId})`, incident: openInc };
  });

  return { config: cfg, counts: counts, vehicles: vehicles, deployments: deployments, incidents: incidents, oosVehicles: oosList, todayDate: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd') };
}