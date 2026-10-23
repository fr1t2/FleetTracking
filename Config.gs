/**
 * @fileoverview System Constants, Config Parser, and Audit Logger.
 */

const CONFIG_SHEET = '_Config';
const AUDIT_SHEET = 'Audit_Log';
const VEHICLES_SHEET = 'Vehicles';
const DEPLOYMENTS_SHEET = 'Deployments';
const INCIDENTS_SHEET = 'Incidents_Service';

/**
 * Reads dynamic system settings and dropdown lists from _Config.
 * @return {Object} System configuration map.
 */
function getSystemConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) throw new Error('Missing sheet: ' + CONFIG_SHEET);

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const config = {
    settings: {},
    shifts: [],
    zones: [],
    oosReasons: [],
    statuses: [],
    focusOptions: [],
    rentalAgencies: [],
    debugLogging: true
  };

  // Read Key-Value pairs (Cols A:B)
  const kvData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  kvData.forEach(([k, v]) => {
    if (k) config.settings[String(k).trim()] = v;
  });

  if (config.settings.hasOwnProperty('Debug_Logging')) {
    config.debugLogging = Boolean(config.settings['Debug_Logging']);
  }

  function getColumnValues(colIndex) {
    return sheet.getRange(2, colIndex, lastRow - 1, 1)
      .getValues()
      .map(r => r[0])
      .filter(v => v !== '' && v !== null && v !== undefined);
  }

  config.shifts = getColumnValues(4);         // Col D
  config.zones = getColumnValues(5);          // Col E
  config.oosReasons = getColumnValues(6);     // Col F
  config.statuses = getColumnValues(7);       // Col G
  config.focusOptions = getColumnValues(8);   // Col H
  config.rentalAgencies = getColumnValues(9); // Col I

  if (!config.rentalAgencies.length) {
    config.rentalAgencies = ['Enterprise', 'Hertz', 'Merchants Fleet', 'Budget', 'Penske', 'Company Owned'];
  }

  return config;
}

/**
 * Standardized audit logging.
 */
function logAudit(actionType, entity, recordId, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getSystemConfig();

    if (!config.debugLogging && actionType === 'DEBUG') return;

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
      logId,
      timestamp,
      userEmail,
      actionType,
      entity,
      recordId || 'N/A',
      detailString
    ]);
  } catch (err) {
    Logger.log('Audit logger failure: ' + err.toString());
  }
}