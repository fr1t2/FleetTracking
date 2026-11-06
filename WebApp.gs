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

/**
 * Helper function to include modular HTML files into the main Dashboard_App
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDashboardState() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = Session.getScriptTimeZone();

    // 1. Fetch Config & Sort Zones (Push DEN OZ to the bottom)
    const cfg = { zones: [], shifts: ['First Shift', 'Second Shift', 'Third Shift'] };
    const cSheet = ss.getSheetByName('_Config');
    if (cSheet && cSheet.getLastRow() > 1) {
      const cData = cSheet.getDataRange().getValues();
      let normalZones = [];
      let ozGroups = [];
      for (let i = 1; i < cData.length; i++) {
        const z = String(cData[i][4]).trim();
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

    // 2. Fetch Vehicles (Adding Type, Oil, Tires, and Severity)
    const vSheet = ss.getSheetByName('Vehicles');
    const vehicles = [];
    if (vSheet && vSheet.getLastRow() > 1) {
      const vData = vSheet.getDataRange().getValues();
      for (let i = 1; i < vData.length; i++) {
        const r = vData[i];
        if (!r[0] && !r[1]) continue;
        
        vehicles.push({
          cleanName: String(r[0] || '').trim(),
          rentalId: String(r[1] || '').trim(),
          type: String(r[2] || 'Cargo Van').trim(), 
          plate: String(r[5] || ''),
          samsaraLink: String(r[8] || ''), 
          status: String(r[9] || 'Warehouse (Ready)').trim(),
          zone: String(r[12] || ''),
          driver: String(r[13] || ''),
          notes: String(r[15] || ''),
          safeToDrive: String(r[16] || '').toLowerCase() === 'yes', 
          lastOil: String(r[17] || 'N/A'),
          lastTires: String(r[18] || 'N/A')
        });
      }
    }

    // 3. Fetch Tech Profiles (Adding Coast, Samsara Violation, Flagged)
    const tSheet = ss.getSheetByName('Tech_Profiles');
    const techs = [];
    if (tSheet && tSheet.getLastRow() > 1) {
      const tData = tSheet.getDataRange().getValues();
      for (let i = 1; i < tData.length; i++) {
        const r = tData[i];
        if (!r[0]) continue;
        techs.push({
          name: String(r[0] || '').trim(),
          passedTest: String(r[6]).trim().toLowerCase() === 'yes',
          coastEnrolled: String(r[12]).trim().toLowerCase() === 'yes',
          samsaraViolation: String(r[13]).trim().toLowerCase() === 'yes',
          flagged: String(r[14]).trim() !== '' ? String(r[14]) : null
        });
      }
    }

    // 4. Fetch Deployments
    const dSheet = ss.getSheetByName('Deployments');
    const deployments = [];
    if (dSheet && dSheet.getLastRow() > 1) {
      const dData = dSheet.getDataRange().getValues();
      const startRow = Math.max(1, dData.length - 200);
      for (let i = startRow; i < dData.length; i++) {
        const r = dData[i];
        if (!r[0]) continue;
        deployments.push({
          id: String(r[0] || ''), vehicle: String(r[3] || ''), rentalId: String(r[4] || ''),
          tech: String(r[5] || ''), zone: String(r[6] || ''), goals: String(r[7] || ''), target: String(r[8] || ''),
          status: String(r[9] || 'Dispatched'),
          departureTime: (r[10] instanceof Date) ? Utilities.formatDate(r[10], tz, 'hh:mm a') : String(r[10] || '')
        });
      }
    }
    
    // 5. Fetch Open Incidents
    const iSheet = ss.getSheetByName('Incidents_Service');
    const incidents = [];
    if (iSheet && iSheet.getLastRow() > 1) {
      const iData = iSheet.getDataRange().getValues();
      for (let i = 1; i < iData.length; i++) {
        if (!iData[i][0] || iData[i][8]) continue; // Skip if no ID or already resolved
        incidents.push({
          id: String(iData[i][0]),
          rentalId: String(iData[i][3]),
          issue: String(iData[i][4]) + ': ' + String(iData[i][5]),
          severity: String(iData[i][6]) === 'YES' ? 'OOS' : 'Minor'
        });
      }
    }

    return { config: cfg, vehicles: vehicles, deployments: deployments, techs: techs, incidents: incidents };
  } catch (err) {
    throw new Error("Backend Crash: " + err.message);
  }
}

/**
 * Helper function to include modular HTML files into the main Dashboard_App.
 * This simulates a folder structure by injecting file contents dynamically.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}