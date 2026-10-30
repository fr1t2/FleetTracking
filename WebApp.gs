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
  const iSheet = ss.getSheetByName('Incidents_Service');
  const dSheet = ss.getSheetByName('Deployments');
  const cSheet = ss.getSheetByName('_Config');
  
  const zones = [];
  if (cSheet && cSheet.getLastRow() > 1) {
    const cData = cSheet.getDataRange().getValues();
    for (let i = 1; i < cData.length; i++) {
      if (cData[i][4]) zones.push(String(cData[i][4]).trim());
    }
  }

  const vehicles = [];
  if (vSheet && vSheet.getLastRow() > 1) {
    const vData = vSheet.getDataRange().getValues();
    for (let i = 1; i < vData.length; i++) {
      if (!vData[i][0] && !vData[i][1]) continue; 
      vehicles.push({
        cleanName: vData[i][0] || '', rentalId: vData[i][1] || '',
        agency: vData[i][2] || '', makeModel: vData[i][3] || '',
        vin: vData[i][4] || '', plate: vData[i][5] || '',
        plateExpDate: vData[i][6] ? new Date(vData[i][6]).toLocaleDateString() : '—', // Col G
        samsaraSerial: vData[i][7] || '', samsaraLink: vData[i][8] || '',
        status: vData[i][9] || 'Warehouse (Ready)',
        pickupDate: vData[i][10] ? new Date(vData[i][10]).toLocaleDateString() : '',
        zone: vData[i][12] || '—', driver: vData[i][13] || '—', notes: vData[i][15] || '',
        coast: vData[i][19] || 'No' // Col T (Index 19)
      });
    }
  }

  const techs = [];
  if (tSheet && tSheet.getLastRow() > 1) {
    const tData = tSheet.getDataRange().getValues();
    for (let i = 1; i < tData.length; i++) {
      if (!tData[i][0]) continue;
      techs.push({
        name: tData[i][0] || '', phone: tData[i][1] || 'No Phone',
        shift: tData[i][2] || 'First Shift', startTime: tData[i][3] || '—', endTime: tData[i][4] || '—',
        workDays: tData[i][5] || '—', passedTest: String(tData[i][6]).trim().toLowerCase() === 'yes',
        restricted: String(tData[i][8]).trim().toLowerCase() === 'yes', status: tData[i][9] || 'Available'
      });
    }
  }

  const incidents = [];
  if (iSheet && iSheet.getLastRow() > 1) {
    const iData = iSheet.getDataRange().getValues();
    for (let i = 1; i < iData.length; i++) {
      if (!iData[i][0]) continue;
      incidents.push({
        id: iData[i][0], date: iData[i][1] ? new Date(iData[i][1]).toLocaleDateString() : '',
        cleanName: iData[i][2] || '', rentalId: iData[i][3] || '',
        category: iData[i][4] || '', desc: iData[i][5] || '',
        oos: iData[i][6] || 'NO', statusImpact: iData[i][7] || '',
        resolution: iData[i][8] || '', resolved: iData[i][9] ? true : false
      });
    }
  }

  const deployments = [];
  if (dSheet && dSheet.getLastRow() > 1) {
    const dData = dSheet.getDataRange().getValues();
    const startRow = Math.max(1, dData.length - 200);
    for (let i = startRow; i < dData.length; i++) {
      deployments.push({
        id: dData[i][0], date: dData[i][1] ? new Date(dData[i][1]).toLocaleDateString() : '',
        shift: dData[i][2] || '', cleanName: dData[i][3] || '', rentalId: dData[i][4] || '',
        tech: dData[i][5] || '', zone: dData[i][6] || '', goals: dData[i][7] || '',
        battery: dData[i][8] || '', status: dData[i][9] || '', departureTime: dData[i][10] ? new Date(dData[i][10]).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}) : ''
      });
    }
  }

  return JSON.stringify({ vehicles, techs, incidents, deployments, zones });
}

/**
 * Helper function to include modular HTML files into the main Dashboard_App.
 * This simulates a folder structure by injecting file contents dynamically.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}