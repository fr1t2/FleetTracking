/**
 * @fileoverview High-Speed Standalone Web App Backend
 * Batch reads all spreadsheet ranges in single calls to minimize latency.
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard_App')
    .evaluate()
    .setTitle('VEO Fleet & Dispatch Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Returns full state with single-pass batch range reads.
 */
function getDashboardState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();

  // 1. FAST BATCH READ: _Config in a single API call
  const cfgSheet = ss.getSheetByName(CONFIG_SHEET);
  const cfg = {
    settings: {},
    shifts: ['Shift 1', 'Shift 2', 'Shift 3'],
    zones: [],
    issueCategories: [],
    statuses: ['Active', 'In Service', 'Out of Service', 'Decommissioned'],
    focusOptions: [],
    rentalAgencies: ['Enterprise', 'Penske', 'U-Haul'],
    statusImpacts: []
  };

  if (cfgSheet && cfgSheet.getLastRow() > 1) {
    const rawCfg = cfgSheet.getRange(2, 1, cfgSheet.getLastRow() - 1, Math.min(cfgSheet.getLastColumn(), 10)).getValues();
    rawCfg.forEach(row => {
      // Key-Value (Cols A:B)
      if (row[0]) cfg.settings[String(row[0]).trim()] = row[1];
      // Lists (Cols D through J)
      if (row[3]) cfg.shifts.push(String(row[3]).trim());
      if (row[4]) cfg.zones.push(String(row[4]).trim());
      if (row[5]) cfg.issueCategories.push(String(row[5]).trim());
      if (row[6]) cfg.statuses.push(String(row[6]).trim());
      if (row[7]) cfg.focusOptions.push(String(row[7]).trim());
      if (row[8]) cfg.rentalAgencies.push(String(row[8]).trim());
      if (row[9]) cfg.statusImpacts.push(String(row[9]).trim());
    });
    // Deduplicate lists
    cfg.shifts = Array.from(new Set(cfg.shifts));
    cfg.statuses = Array.from(new Set(cfg.statuses));
    cfg.rentalAgencies = Array.from(new Set(cfg.rentalAgencies));
  }

  // 2. FAST BATCH READ: Vehicles (17 columns)
  const vSheet = ss.getSheetByName(VEHICLES_SHEET);
  const vehicles = [];
  const counts = { total: 0, active: 0, oos: 0, decommissioned: 0 };

  if (vSheet && vSheet.getLastRow() > 1) {
    const vRows = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, 17).getValues();
    vRows.forEach(r => {
      const cleanName = String(r[0] || '').trim();
      const rentalId = String(r[1] || '').trim();
      const status = String(r[8] || 'Active').trim(); // Col I is Status
      if (!cleanName && !rentalId) return;

      counts.total++;
      if (status === 'Active' || status === 'In Service') counts.active++;
      else if (status === 'Out of Service') counts.oos++;
      else if (status === 'Decommissioned') counts.decommissioned++;

      let expDate = '—';
      if (r[6] instanceof Date) expDate = Utilities.formatDate(r[6], tz, 'yyyy-MM-dd');
      else if (r[6]) expDate = String(r[6]).trim();

      vehicles.push({
        cleanName: cleanName,
        rentalId: rentalId,
        agency: String(r[2] || ''),
        makeModel: String(r[3] || 'Cargo Van'),
        vin: String(r[4] || ''),
        plate: String(r[5] || ''),
        plateExpDate: expDate,
        samsaraSerial: String(r[7] || 'Unassigned'),
        status: status,
        pickupDate: (r[9] instanceof Date) ? Utilities.formatDate(r[9], tz, 'yyyy-MM-dd') : String(r[9] || '—'),
        serviceNum: String(r[10] || ''),
        zone: String(r[11] || 'Unassigned'),
        driver: String(r[12] || 'Unassigned'),
        lastService: (r[13] instanceof Date) ? Utilities.formatDate(r[13], tz, 'yyyy-MM-dd') : '—',
        notes: String(r[14] || '')
      });
    });
  }

  // 3. FAST BATCH READ: Deployments
  const dSheet = ss.getSheetByName(DEPLOYMENTS_SHEET);
  const deployments = [];

  if (dSheet && dSheet.getLastRow() > 1) {
    const dRows = dSheet.getRange(2, 1, dSheet.getLastRow() - 1, 14).getValues();
    dRows.forEach(r => {
      let depDate = '';
      if (r[1] instanceof Date) {
        depDate = Utilities.formatDate(r[1], tz, 'yyyy-MM-dd');
      } else if (r[1]) {
        // Safe string parse fallback
        const parsed = new Date(r[1]);
        depDate = !isNaN(parsed) ? Utilities.formatDate(parsed, tz, 'yyyy-MM-dd') : String(r[1]);
      }

      let timeOut = '—';
      if (r[10] instanceof Date) {
        timeOut = Utilities.formatDate(r[10], tz, 'hh:mm a');
      } else if (r[10]) {
        timeOut = String(r[10]);
      }

      deployments.push({
        id: String(r[0] || ''),
        date: depDate,
        shift: String(r[2] || '').trim(),
        vehicle: String(r[3] || ''),
        rentalId: String(r[4] || ''),
        tech: String(r[5] || ''),
        zone: String(r[6] || ''),
        goals: String(r[7] || ''),
        target: String(r[8] || ''),
        status: String(r[9] || 'Dispatched'),
        departureTime: timeOut
      });
    });
  }

  // 4. FAST BATCH READ: Incidents
  const iSheet = ss.getSheetByName(INCIDENTS_SHEET);
  const incidents = [];

  if (iSheet && iSheet.getLastRow() > 1) {
    const iRows = iSheet.getRange(2, 1, iSheet.getLastRow() - 1, 13).getValues();
    iRows.forEach(r => {
      let incDate = (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1] || '');
      let resDate = (r[9] instanceof Date) ? Utilities.formatDate(r[9], tz, 'yyyy-MM-dd') : String(r[9] || '');

      incidents.push({
        id: String(r[0] || ''),
        date: incDate,
        vehicle: String(r[2] || ''),
        rentalId: String(r[3] || ''),
        category: String(r[4] || ''),
        details: String(r[5] || ''),
        oos: String(r[6] || ''),
        impact: String(r[7] || ''),
        resolution: String(r[8] || ''),
        dateResolved: resDate,
        resolvedBy: String(r[10] || '')
      });
    });
  }

  // 5. Build Out-of-Service Vehicles list
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
    config: cfg,
    counts: counts,
    vehicles: vehicles,
    deployments: deployments,
    incidents: incidents,
    oosVehicles: oosList,
    todayDate: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')
  };
}