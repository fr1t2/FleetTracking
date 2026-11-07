/**
 * @fileoverview Web App Router & Dashboard State Compiler
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard_App')
    .evaluate()
    .setTitle('VEO Dispatch OS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardState() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = Session.getScriptTimeZone();
    const vSheet = ss.getSheetByName('Vehicles');
    const tSheet = ss.getSheetByName('Tech_Profiles');
    const iSheet = ss.getSheetByName('Incidents_Service');
    const dSheet = ss.getSheetByName('Deployments');
    const cSheet = ss.getSheetByName('_Config');
    
    const zones = [];
    const issueCategories = []; 
    const rentalAgencies = [];
    const vehicleTypes = []; 
    
    if (cSheet && cSheet.getLastRow() > 1) {
      const cData = cSheet.getDataRange().getValues();
      for (let i = 1; i < cData.length; i++) {
        if (cData[i][4]) zones.push(String(cData[i][4]).trim()); // Col E: Zones
        if (cData[i][5]) issueCategories.push(String(cData[i][5]).trim()); // Col F: Issues
        if (cData[i][8]) rentalAgencies.push(String(cData[i][8]).trim()); // Col I: Agencies (Owner)
        if (cData[i][10]) vehicleTypes.push(String(cData[i][10]).trim()); // Col K: Vehicle Types
      }
    }

    // 1. Fetch Config Zones (Col E / Index 4)
    const cfg = { zones: [] };
    if (cSheet && cSheet.getLastRow() > 1) {
      const cData = cSheet.getDataRange().getValues();
      let normalZones = [];
      let ozGroups = [];
      for (let i = 1; i < cData.length; i++) {
        const z = String(cData[i][4] || '').trim();
        if (z) {
          if (z.toUpperCase().includes('DENVER OZ') || z.toUpperCase().includes('DEN OZ')) {
            ozGroups.push(z);
          } else {
            normalZones.push(z);
          }
        }
      }
      normalZones.sort();
      ozGroups.sort();
      cfg.zones = normalZones.concat(ozGroups); 
    }

    // 2. Fetch Vehicles (24 Columns)
    const vehicles = [];
    if (vSheet && vSheet.getLastRow() > 1) {
      const vData = vSheet.getDataRange().getValues();
      for (let i = 1; i < vData.length; i++) {
        const r = vData[i];
        while (r.length < 24) r.push(''); // Pad array to prevent out-of-bounds crash
        if (!r[0] && !r[1]) continue;
        
        vehicles.push({
          cleanName: String(r[0] || '').trim(),
          rentalId: String(r[1] || '').trim(),
          agency: String(r[2] || ''),
          makeModel: String(r[3] || ''),
          vin: String(r[4] || ''),
          plate: String(r[5] || ''),
          plateExpDate: (r[6] instanceof Date) ? Utilities.formatDate(r[6], tz, 'yyyy-MM-dd') : String(r[6] || '—'),
          samsaraSerial: String(r[7] || ''),
          samsaraLink: String(r[8] || ''),
          status: String(r[9] || 'Warehouse (Ready)').trim(),
          pickupDate: (r[10] instanceof Date) ? Utilities.formatDate(r[10], tz, 'yyyy-MM-dd') : String(r[10] || '—'),
          serviceNum: String(r[11] || ''),
          zone: String(r[12] || '—'),
          driver: String(r[13] || '—'),
          lastService: (r[14] instanceof Date) ? Utilities.formatDate(r[14], tz, 'yyyy-MM-dd') : '—',
          notes: String(r[15] || ''),
          addons: String(r[18] || ''),
          coast: String(r[19] || 'No'),
          type: String(r[20] || 'Cargo Van').trim(),
          lastOil: (r[21] instanceof Date) ? Utilities.formatDate(r[21], tz, 'yyyy-MM-dd') : String(r[21] || 'N/A'),
          lastOilMiles: String(r[22] || 'N/A'),
          lastTires: (r[23] instanceof Date) ? Utilities.formatDate(r[23], tz, 'yyyy-MM-dd') : String(r[23] || 'N/A'),
          safeToDrive: 'false'
        });
      }
    }

    // 3. Fetch Tech Profiles (17 Columns)
    const tSheet = ss.getSheetByName('Tech_Profiles');
    const techs = [];
    if (tSheet && tSheet.getLastRow() > 1) {
      const tRows = tSheet.getRange(2, 1, tSheet.getLastRow() - 1, 17).getValues();
      tRows.forEach(r => {
        if (!r[0]) return;
        techs.push({
          name: String(r[0]).trim(),
          phone: String(r[1] || '').trim(),
          shift: String(r[2] || '').trim(),
          start: String(r[3] || '').trim(),
          end: String(r[4] || '').trim(),
          workDays: String(r[5] || '').trim(),
          passedTest: (String(r[6] || '').trim().toLowerCase() === 'yes'),
          goFast: (String(r[7] || '').trim().toLowerCase() === 'yes'),
          restricted: (String(r[8] || '').trim().toLowerCase() === 'yes'),
          status: String(r[9] || 'Available').trim(),
          notes: String(r[10] || '').trim(),
          hrStatus: String(r[11] || 'Active').trim(),
          coastEnrolled: (String(r[12] || '').trim().toLowerCase() === 'yes'),
          evgo: (String(r[13] || '').trim().toLowerCase() === 'yes'),
          trainer: (String(r[14] || '').trim().toLowerCase() === 'yes'),
          samsaraViolation: (String(r[15] || '').trim().toLowerCase() === 'yes'),
          flagged: String(r[16] || '').trim()
        });
      });
    }

    // 4. Fetch Deployments (21 Columns)
    const deployments = [];
    if (dSheet && dSheet.getLastRow() > 1) {
      const dData = dSheet.getDataRange().getValues();
      const startRow = Math.max(1, dData.length - 200); // Only pull latest 200 for speed
      for (let i = startRow; i < dData.length; i++) {
        const r = dData[i];
        while (r.length < 21) r.push('');
        if (!r[0]) continue;
        
        let depDate = (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1] || '').split(' ')[0];
        let timeOut = (r[10] instanceof Date) ? Utilities.formatDate(r[10], tz, 'hh:mm a') : String(r[10] || '—');
        let timeIn = (r[11] instanceof Date) ? Utilities.formatDate(r[11], tz, 'hh:mm a') : String(r[11] || '');
        
        deployments.push({
          id: String(r[0] || ''), 
          date: depDate, 
          shift: String(r[2] || '').trim(), 
          vehicle: String(r[3] || ''), 
          rentalId: String(r[4] || ''), 
          tech: String(r[5] || ''), 
          zone: String(r[6] || ''), 
          goals: String(r[7] || ''),
          battery: String(r[8] || ''), 
          status: String(r[9] || 'Dispatched'), 
          departureTime: timeOut, 
          returnTime: timeIn,
          coDriver: String(r[14] || ''), 
          isTraining: String(r[15] || '') === 'Yes'
        });
      }
    }

    // 5. Fetch Incidents (14 Columns)
    const incidents = [];
    if (iSheet && iSheet.getLastRow() > 1) {
      const iData = iSheet.getDataRange().getValues();
      for (let i = 1; i < iData.length; i++) {
        const r = iData[i];
        while (r.length < 14) r.push('');
        if (!r[0]) continue;
        
        incidents.push({
          id: String(r[0] || ''), 
          dateReported: (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
          vehicle: String(r[2] || ''), 
          rentalId: String(r[3] || ''), 
          category: String(r[4] || ''), 
          issue: String(r[5] || ''),
          desc: String(r[5] || ''), 
          impact: String(r[7] || ''), 
          resolution: String(r[8] || ''), 
          dateResolved: (r[9] instanceof Date) ? Utilities.formatDate(r[9], tz, 'yyyy-MM-dd') : String(r[9]),
          severity: String(r[13] || 'Minor')
        });
      }
    }

    // Attach "safeToDrive" flag to vehicles based on incidents
    vehicles.forEach(v => {
      if (v.status === 'Out of Service') {
        v.safeToDrive = 'false';
      } else {
        const openInc = incidents.find(i => i.rentalId === v.rentalId && !i.dateResolved);
        if (openInc && openInc.severity !== 'OOS') v.safeToDrive = 'true';
      }
    });

    return JSON.stringify({ 
        vehicles, techs, incidents, deployments, zones, 
        issueCategories, rentalAgencies, vehicleTypes 
    });
  } catch (err) {
    throw new Error("Backend Crash: " + err.message);
  }
}