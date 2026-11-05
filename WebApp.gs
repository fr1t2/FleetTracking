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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const cfg = { zones: ['Zone 1A', 'Zone 1B', 'Zone 2A', 'Zone 2B', 'Zone 3A'] }; 
  
  // Try to load config dynamically if the sheet exists
  const cSheet = ss.getSheetByName('_Config');
  if (cSheet && cSheet.getLastRow() > 1) {
    cfg.zones = [];
    const cData = cSheet.getDataRange().getValues();
    for (let i = 1; i < cData.length; i++) {
      if (cData[i][4]) cfg.zones.push(String(cData[i][4]).trim());
    }
  }

  // 1. Fetch Vehicles Safely
  const vSheet = ss.getSheetByName('Vehicles');
  const vehicles = [];
  if (vSheet && vSheet.getLastRow() > 1) {
    const vData = vSheet.getDataRange().getValues(); // Dynamically sizes to exact columns
    for (let i = 1; i < vData.length; i++) {
      const r = vData[i];
      if (!r[0] && !r[1]) continue; 
      vehicles.push({
        cleanName: String(r[0] || '').trim(),
        rentalId: String(r[1] || '').trim(),
        plate: String(r[5] || ''),
        status: String(r[8] || r[9] || 'Warehouse (Ready)').trim(), // Covers multiple column layouts
        zone: String(r[11] || r[12] || '—'),
        driver: String(r[12] || r[13] || '—')
      });
    }
  }

  // 2. Fetch Deployments Safely
  const dSheet = ss.getSheetByName('Deployments');
  const deployments = [];
  if (dSheet && dSheet.getLastRow() > 1) {
    const dData = dSheet.getDataRange().getValues();
    const startRow = Math.max(1, dData.length - 200);
    for (let i = startRow; i < dData.length; i++) {
      const r = dData[i];
      deployments.push({
        id: String(r[0] || ''),
        vehicle: String(r[3] || ''),
        rentalId: String(r[4] || ''),
        tech: String(r[5] || ''),
        zone: String(r[6] || ''),
        target: String(r[8] || ''),
        status: String(r[9] || 'Dispatched'),
        departureTime: (r[10] instanceof Date) ? Utilities.formatDate(r[10], tz, 'hh:mm a') : String(r[10] || '')
      });
    }
  }

  // 3. Fetch Techs Safely
  const tSheet = ss.getSheetByName('Tech_Profiles');
  const techs = [];
  if (tSheet && tSheet.getLastRow() > 1) {
    const tData = tSheet.getDataRange().getValues();
    for (let i = 1; i < tData.length; i++) {
      const r = tData[i];
      if (!r[0]) continue;
      techs.push({
        name: String(r[0] || ''),
        phone: String(r[1] || 'No Phone'),
        shift: String(r[2] || 'First Shift'),
        startTime: String(r[3] || '—'),
        endTime: String(r[4] || '—'),
        workDays: String(r[5] || '—'),
        passedTest: String(r[6]).trim().toLowerCase() === 'yes',
        restricted: String(r[8]).trim().toLowerCase() === 'yes',
        status: String(r[9] || 'Available')
      });
    }
  }

  return { config: cfg, vehicles: vehicles, deployments: deployments, techs: techs };
}

/**
 * Helper function to include modular HTML files into the main Dashboard_App.
 * This simulates a folder structure by injecting file contents dynamically.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}