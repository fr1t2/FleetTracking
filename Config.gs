/**
 * @fileoverview System Constants, Config Parser, Centralized Audit Logger, and Roster Mgmt.
 */

const CONFIG_SHEET = '_Config';
const AUDIT_SHEET = 'Audit_Log';
const VEHICLES_SHEET = 'Vehicles';
const DEPLOYMENTS_SHEET = 'Deployments';
const INCIDENTS_SHEET = 'Incidents_Service';

function getSystemConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) throw new Error('Missing sheet: ' + CONFIG_SHEET);

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const config = {
    settings: {},
    shifts: [],
    zones: [],
    issueCategories: [],
    statuses: [],
    focusOptions: [],
    rentalAgencies: [],
    statusImpacts: [],
    techRoster: [], // Col K
    debugLogging: true
  };

  // Read config block out to Col 11 (K) for the Tech Roster
  const rawCfg = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  rawCfg.forEach(row => {
    if (row[0]) config.settings[String(row[0]).trim()] = row[1];
    if (row[3]) config.shifts.push(String(row[3]).trim());
    if (row[4]) config.zones.push(String(row[4]).trim());
    if (row[5]) config.issueCategories.push(String(row[5]).trim());
    if (row[6]) config.statuses.push(String(row[6]).trim());
    if (row[7]) config.focusOptions.push(String(row[7]).trim());
    if (row[8]) config.rentalAgencies.push(String(row[8]).trim());
    if (row[9]) config.statusImpacts.push(String(row[9]).trim());
    if (row[10]) config.techRoster.push(String(row[10]).trim()); // Col K
  });

  // Deduplicate and remove blanks safely
  config.shifts = Array.from(new Set(config.shifts)).filter(Boolean);
  config.zones = Array.from(new Set(config.zones)).filter(Boolean);
  config.issueCategories = Array.from(new Set(config.issueCategories)).filter(Boolean);
  config.statuses = Array.from(new Set(config.statuses)).filter(Boolean);
  config.focusOptions = Array.from(new Set(config.focusOptions)).filter(Boolean);
  config.rentalAgencies = Array.from(new Set(config.rentalAgencies)).filter(Boolean);
  config.statusImpacts = Array.from(new Set(config.statusImpacts)).filter(Boolean);
  config.techRoster = Array.from(new Set(config.techRoster)).filter(Boolean);

  return config;
}

function logAudit(actionType, entity, recordId, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let auditSheet = ss.getSheetByName(AUDIT_SHEET);
    if (!auditSheet) {
      auditSheet = ss.insertSheet(AUDIT_SHEET);
      auditSheet.appendRow(['Log ID', 'Timestamp', 'User Email', 'Action Type', 'Target Entity', 'Record ID', 'Details / Change Delta']);
    }

    const logId = 'LOG-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const timestamp = new Date();
    const userEmail = Session.getActiveUser().getEmail() || 'local-user';
    const detailString = typeof details === 'object' ? JSON.stringify(details) : String(details);

    auditSheet.appendRow([
      logId, timestamp, userEmail, actionType, entity, recordId || 'N/A', detailString
    ]);
  } catch (err) {
    Logger.log('Audit logger failure: ' + err.toString());
  }
}

/**
 * Roster Management API for the Dashboard
 */
function addTechToRoster(name) {
  if (!name) throw new Error("Employee name is required.");
  const cleanName = name.trim();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const rosterData = sheet.getRange(2, 11, lastRow - 1, 1).getValues();
  
  // 1. Prevent exact duplicates from being added
  for (let i = 0; i < rosterData.length; i++) {
    if (String(rosterData[i][0]).trim().toLowerCase() === cleanName.toLowerCase()) {
      return { success: true, message: `${cleanName} is already on the roster.` };
    }
  }

  // 2. Find the first empty cell in Column K and insert
  let insertRow = lastRow + 1;
  for (let i = 0; i < rosterData.length; i++) {
    if (!String(rosterData[i][0]).trim()) {
      insertRow = i + 2;
      break;
    }
  }
  
  sheet.getRange(insertRow, 11).setValue(cleanName);
  logAudit('ROSTER_UPDATED', 'System', cleanName, 'Added to tech roster');
  return { success: true, message: `${cleanName} added to employee roster.` };
}

function removeTechFromRoster(name) {
  if (!name) throw new Error("Employee name is required.");
  const cleanName = name.trim().toLowerCase();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  
  // Read the current roster list
  const rosterData = sheet.getRange(2, 11, lastRow - 1, 1).getValues();
  const newRosterList = [];
  
  // 1. Filter out ALL instances of the removed name, keeping the rest
  for (let i = 0; i < rosterData.length; i++) {
    const cellValue = String(rosterData[i][0]).trim();
    if (cellValue && cellValue.toLowerCase() !== cleanName) {
      newRosterList.push([cellValue]);
    }
  }
  
  // 2. Clear the entire column K below the header
  sheet.getRange(2, 11, lastRow, 1).clearContent();
  
  // 3. Write the clean, gap-free list back to the sheet
  if (newRosterList.length > 0) {
    sheet.getRange(2, 11, newRosterList.length, 1).setValues(newRosterList);
  }
  
  logAudit('ROSTER_UPDATED', 'System', name, 'Removed from tech roster');
  return { success: true, message: `${name.trim()} removed from roster.` };
}