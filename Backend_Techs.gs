/**
 * Run this function ONCE from the editor to upgrade your database.
 * It creates the Tech_Profiles sheet and migrates your existing roster.
 */
function upgradeDatabaseForTechs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let techSheet = ss.getSheetByName('Tech_Profiles');
  
  if (!techSheet) {
    techSheet = ss.insertSheet('Tech_Profiles');
    techSheet.appendRow(['Tech Name', 'Phone', 'Shift', 'Start Time', 'End Time', 'Work Days', 'Passed Test', 'Go-Fast', 'Restricted', 'Status', 'Notes']);
    techSheet.getRange('A1:K1').setFontWeight('bold').setBackground('#f0f2f5');
    
    // Migrate existing techs from _Config if they exist
    const configSheet = ss.getSheetByName('_Config');
    if (configSheet) {
      const roster = configSheet.getRange('K2:K' + configSheet.getLastRow()).getValues().flat().filter(String);
      roster.forEach(name => {
        techSheet.appendRow([name, '', 'First Shift', '07:00', '15:30', 'Mon-Fri', 'Yes', 'No', 'No', 'Available', 'Migrated']);
      });
    }
  }
}

/**
 * Updates a Tech's Profile attributes.
 */
function updateTechProfile(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Tech_Profiles');
  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === String(data.origName).trim().toLowerCase()) {
      rowIndex = i + 1; break;
    }
  }

  const rowData = [
    data.name, data.phone, data.shift, data.startTime, data.endTime, 
    data.workDays, data.passedTest ? 'Yes' : 'No', data.goFast ? 'Yes' : 'No', 
    data.restricted ? 'Yes' : 'No', data.status, data.notes || ''
  ];

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, 11).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  
  logAudit('TECH_UPDATED', 'Tech_Profiles', data.name, `Status: ${data.status} | Restricted: ${data.restricted}`);
  return { success: true };
}