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
  const cfg = getSystemConfig(); 

  // 1. Vehicles
  const vSheet = ss.getSheetByName('Vehicles');
  const vehicles = [];
  if (vSheet && vSheet.getLastRow() > 1) {
    const vRows = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, 19).getValues();
    vRows.forEach(r => {
      const cleanName = String(r[0] || '').trim();
      const rentalId = String(r[1] || '').trim();
      if (!cleanName && !rentalId) return;
      vehicles.push({
        cleanName: cleanName, rentalId: rentalId, agency: String(r[2] || ''), makeModel: String(r[3] || ''), 
        vin: String(r[4] || ''), plate: String(r[5] || ''), 
        plateExpDate: (r[6] instanceof Date) ? Utilities.formatDate(r[6], tz, 'yyyy-MM-dd') : String(r[6] || '—'),
        samsaraSerial: String(r[7] || ''), samsaraLink: String(r[8] || ''),
        status: String(r[9] || 'Warehouse (Ready)').trim(), pickupDate: (r[10] instanceof Date) ? Utilities.formatDate(r[10], tz, 'yyyy-MM-dd') : String(r[10] || '—'),
        serviceNumber: String(r[11] || ''), zone: String(r[12] || '—'), driver: String(r[13] || '—'),
        lastService: (r[14] instanceof Date) ? Utilities.formatDate(r[14], tz, 'yyyy-MM-dd') : '—', notes: String(r[15] || ''), addons: String(r[18] || '')
      });
    });
  }

  // 2. Deployments (Now pulls fake attendance deployments too!)
  const dSheet = ss.getSheetByName('Deployments');
  const deployments = [];
  if (dSheet && dSheet.getLastRow() > 1) {
    const dRows = dSheet.getRange(2, 1, dSheet.getLastRow() - 1, 16).getValues();
    dRows.forEach(r => {
      let depDate = (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1] || '').split(' ')[0];
      let timeOut = (r[10] instanceof Date) ? Utilities.formatDate(r[10], tz, 'hh:mm a') : String(r[10] || '—');
      let timeIn = (r[11] instanceof Date) ? Utilities.formatDate(r[11], tz, 'hh:mm a') : String(r[11] || '');
      deployments.push({
        id: String(r[0] || ''), date: depDate, shift: String(r[2] || '').trim(), vehicle: String(r[3] || ''), 
        rentalId: String(r[4] || ''), tech: String(r[5] || ''), zone: String(r[6] || ''), goals: String(r[7] || ''),
        target: String(r[8] || ''), status: String(r[9] || 'Dispatched'), departureTime: timeOut, returnTime: timeIn,
        coDriver: String(r[14] || ''), isTraining: String(r[15] || '') === 'Yes'
      });
    });
  }

  // 3. Incidents
  const iSheet = ss.getSheetByName('Incidents_Service');
  const incidents = [];
  if (iSheet && iSheet.getLastRow() > 1) {
    const maxCol = iSheet.getLastColumn();
    const iRows = iSheet.getRange(2, 1, iSheet.getLastRow() - 1, maxCol < 14 ? 14 : maxCol).getValues();
    iRows.forEach(r => {
      incidents.push({
        id: String(r[0] || ''), dateReported: (r[1] instanceof Date) ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
        vehicle: String(r[2] || ''), rentalId: String(r[3] || ''), category: String(r[4] || ''), details: String(r[5] || ''), 
        impact: String(r[7] || ''), resolution: String(r[8] || ''), dateResolved: (r[9] instanceof Date) ? Utilities.formatDate(r[9], tz, 'yyyy-MM-dd') : String(r[9]),
        driverInvolved: String(r[13] || '')
      });
    });
  }

  // 4. Tech Profiles (Upgraded for accurate time parsing and new columns)
  const tSheet = ss.getSheetByName('Tech_Profiles');
  const techs = [];
  if (tSheet && tSheet.getLastRow() > 1) {
    const tRows = tSheet.getRange(2, 1, tSheet.getLastRow() - 1, 15).getValues();
    tRows.forEach(r => {
      if(r[0]) {
        // Enforce strict HH:mm formatting for the frontend timeline graph
        let sTime = r[3] instanceof Date ? Utilities.formatDate(r[3], tz, 'HH:mm') : String(r[3] || '');
        let eTime = r[4] instanceof Date ? Utilities.formatDate(r[4], tz, 'HH:mm') : String(r[4] || '');
        
        techs.push({
          name: String(r[0]).trim(), phone: String(r[1] || ''), shift: String(r[2] || 'First Shift'), 
          startTime: sTime, endTime: eTime, workDays: String(r[5] || 'Mon-Fri'), 
          passedTest: String(r[6] || 'No') === 'Yes', goFast: String(r[7] || 'No') === 'Yes', restricted: String(r[8] || 'No') === 'Yes', 
          status: String(r[9] || 'Available'), notes: String(r[10] || ''), hrStatus: String(r[11] || 'Active'),
          coastFuel: String(r[12] || 'No') === 'Yes', evGo: String(r[13] || 'No') === 'Yes', trainer: String(r[14] || 'No') === 'Yes'
        });
      }
    });
  }

  return { config: cfg, vehicles: vehicles, deployments: deployments, incidents: incidents, techs: techs };
}