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
  const dSheet = ss.getSheetByName('Deployments');
  
  // 1. Fetch Vehicles (Mapping all 19 columns)
  const vehicles = [];
  if (vSheet && vSheet.getLastRow() > 1) {
    const vRows = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, 19).getValues();
    vRows.forEach(r => {
      if (!r[0] && !r[1]) return; // Skip empty rows
      vehicles.push({
        cleanName: r[0],
        rentalId: r[1],
        agency: r[2],
        makeModel: r[3],
        vin: r[4],
        plate: r[5],
        plateExpDate: r[6],
        samsaraSerial: r[7],
        samsaraLink: r[8], // NEW COLUMN
        status: r[9] || 'Warehouse (Ready)',
        pickupDate: r[10],
        serviceNum: r[11],
        zone: r[12] || '—',
        driver: r[13] || '—',
        lastService: r[14],
        notes: r[15]
      });
    });
  }

  // 2. Fetch Techs & Deployments (Simplified for the unified view)
  const techs = []; // Pull from your _Config tab as before
  const configSheet = ss.getSheetByName('_Config');
  if (configSheet) {
    const roster = configSheet.getRange('K2:K' + configSheet.getLastRow()).getValues().flat().filter(String);
    roster.forEach(name => techs.push({ name: name, shift: "First Shift", restricted: false, passedTest: true }));
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