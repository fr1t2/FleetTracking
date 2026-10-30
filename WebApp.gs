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
  const vSheet = ss.getSheetByName('Vehicles');
  const tSheet = ss.getSheetByName('Tech_Profiles');
  
  // 1. Fetch Vehicles
  const vehicles = [];
  if (vSheet && vSheet.getLastRow() > 1) {
    const vRows = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, 19).getValues();
    vRows.forEach(r => {
      if (!r[0] && !r[1]) return; 
      vehicles.push({
        cleanName: r[0],
        rentalId: r[1],
        plate: r[5],
        samsaraSerial: r[7],
        samsaraLink: r[8],
        status: r[9] || 'Warehouse (Ready)',
        zone: r[12] || '—',
        driver: r[13] || '—'
      });
    });
  }

  // 2. Fetch Techs from Tech_Profiles (Extended Attributes)
  const techs = [];
  if (tSheet && tSheet.getLastRow() > 1) {
    const tRows = tSheet.getRange(2, 1, tSheet.getLastRow() - 1, 15).getValues();
    tRows.forEach(r => {
      if (!r[0]) return;
      techs.push({
        name: r[0],
        phone: r[1] || 'No Phone',
        shift: r[2] || 'First Shift',
        startTime: r[3] || '—',
        endTime: r[4] || '—',
        workDays: r[5] || '—',
        passedTest: String(r[6]).trim().toLowerCase() === 'yes',
        restricted: String(r[8]).trim().toLowerCase() === 'yes',
        status: r[9] || 'Available'
      });
    });
  }

  return { vehicles: vehicles, techs: techs };
}

/**
 * Helper function to include modular HTML files into the main Dashboard_App.
 * This simulates a folder structure by injecting file contents dynamically.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}