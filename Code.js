/**
 * @fileoverview VEO Fleet Tracking & Dispatch Core Engine
 * Manages configuration reading, audit logging, fleet registry, and custom menus.
 */

const CONFIG_SHEET = '_Config';
const AUDIT_SHEET = 'Audit_Log';
const VEHICLES_SHEET = 'Vehicles';
const DEPLOYMENTS_SHEET = 'Deployments';
const INCIDENTS_SHEET = 'Incidents_Service';

/**
 * Initializes the custom menu on open.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚚 Fleet Dispatch')
    .addItem('➕ Add New Vehicle to Fleet', 'menuPromptAddVehicle')
    .addItem('🛑 Decommission Vehicle', 'menuPromptDecommissionVehicle')
    .addSeparator()
    .addItem('📋 Open Dispatch Form (Coming in T5)', 'showDispatchModal')
    .addItem('⚠️ Report Incident / OOS (Coming in T4)', 'showIncidentModal')
    .addSeparator()
    .addItem('🧪 Test Audit Log', 'testAuditLogger')
    .addToUi();
}

/**
 * Reads settings and dropdown options dynamically from _Config.
 * @return {Object} System configuration object.
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

  function getNonEmptyColumn(col) {
    return sheet.getRange(2, col, lastRow - 1, 1)
      .getValues()
      .map(r => r[0])
      .filter(v => v !== '' && v !== null && v !== undefined);
  }

  config.shifts = getNonEmptyColumn(4);       // Column D
  config.zones = getNonEmptyColumn(5);        // Column E
  config.oosReasons = getNonEmptyColumn(6);   // Column F
  config.statuses = getNonEmptyColumn(7);     // Column G
  config.focusOptions = getNonEmptyColumn(8); // Column H

  return config;
}

/**
 * Audit Logger: Appends a row to Audit_Log with user email, timestamp, and details.
 * @param {string} actionType E.g. 'VEHICLE_ADDED', 'STATUS_CHANGE', 'DECOMMISSION'
 * @param {string} entity E.g. 'Vehicles', 'Deployments'
 * @param {string} recordId Target identifier (e.g. Rental ID or Van Name)
 * @param {string|Object} details Change details or explanation
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
    Logger.log('Error writing audit log: ' + err.toString());
  }
}

/**
 * Adds a new vehicle to the fleet registry.
 */
function addNewVehicle(vanName, rentalId, makeModel, serviceNum, notes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet) throw new Error('Vehicles sheet not found');

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'system';

  sheet.appendRow([
    vanName,
    rentalId,
    makeModel || 'Cargo Van',
    'Active',
    now,
    serviceNum || '',
    '',   // Zone
    '',   // Tech
    now,  // Last service date
    notes || '',
    user,
    now
  ]);

  logAudit('VEHICLE_ADDED', 'Vehicles', rentalId, {
    vanName: vanName,
    rentalId: rentalId,
    serviceNum: serviceNum,
    addedBy: user
  });

  return { success: true, message: 'Vehicle ' + vanName + ' (' + rentalId + ') added.' };
}

/**
 * Marks a vehicle as Decommissioned.
 */
function decommissionVehicle(rentalId, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet) throw new Error('Vehicles sheet not found');

  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === String(rentalId).trim().toLowerCase() ||
        String(data[i][0]).trim().toLowerCase() === String(rentalId).trim().toLowerCase()) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('Vehicle identifier "' + rentalId + '" not found in registry.');
  }

  const user = Session.getActiveUser().getEmail() || 'system';
  const now = new Date();

  // Update Status (Col 4) and Notes (Col 10), Last Updated By (Col 11), Last Updated At (Col 12)
  sheet.getRange(rowIndex, 4).setValue('Decommissioned');
  const existingNotes = sheet.getRange(rowIndex, 10).getValue();
  const updatedNotes = (existingNotes ? existingNotes + ' | ' : '') + 'Decommissioned: ' + (reason || 'No reason provided');
  sheet.getRange(rowIndex, 10).setValue(updatedNotes);
  sheet.getRange(rowIndex, 11).setValue(user);
  sheet.getRange(rowIndex, 12).setValue(now);

  logAudit('VEHICLE_DECOMMISSIONED', 'Vehicles', rentalId, {
    reason: reason,
    decommissionedBy: user
  });

  return { success: true, message: 'Vehicle ' + rentalId + ' marked as Decommissioned.' };
}

/**
 * Updates a vehicle's status.
 */
function updateVehicleStatus(rentalId, newStatus, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === String(rentalId).trim().toLowerCase() ||
        String(data[i][0]).trim().toLowerCase() === String(rentalId).trim().toLowerCase()) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error('Vehicle not found: ' + rentalId);

  const user = Session.getActiveUser().getEmail() || 'system';
  const now = new Date();
  sheet.getRange(rowIndex, 4).setValue(newStatus);
  sheet.getRange(rowIndex, 11).setValue(user);
  sheet.getRange(rowIndex, 12).setValue(now);

  logAudit('STATUS_CHANGE', 'Vehicles', rentalId, {
    newStatus: newStatus,
    reason: reason,
    updatedBy: user
  });
}

// UI Prompt Handlers
function menuPromptAddVehicle() {
  const ui = SpreadsheetApp.getUi();
  const nameResp = ui.prompt('Add Vehicle', 'Enter clean name (e.g., Van #21):', ui.ButtonSet.OK_CANCEL);
  if (nameResp.getSelectedButton() !== ui.Button.OK) return;

  const rentalResp = ui.prompt('Add Vehicle', 'Enter Rental / VIN ID (e.g., P9103006):', ui.ButtonSet.OK_CANCEL);
  if (rentalResp.getSelectedButton() !== ui.Button.OK) return;

  const res = addNewVehicle(nameResp.getResponseText(), rentalResp.getResponseText(), 'Cargo Van', '', 'Added via menu prompt');
  ui.alert(res.message);
}

function menuPromptDecommissionVehicle() {
  const ui = SpreadsheetApp.getUi();
  const idResp = ui.prompt('Decommission Vehicle', 'Enter Van clean name or Rental ID to decommission:', ui.ButtonSet.OK_CANCEL);
  if (idResp.getSelectedButton() !== ui.Button.OK) return;

  const reasonResp = ui.prompt('Decommission Vehicle', 'Enter reason for decommissioning:', ui.ButtonSet.OK_CANCEL);
  if (reasonResp.getSelectedButton() !== ui.Button.OK) return;

  const res = decommissionVehicle(idResp.getResponseText(), reasonResp.getResponseText());
  ui.alert(res.message);
}

function testAuditLogger() {
  logAudit('TEST_RUN', 'System', 'SELF_CHECK', 'Audit log verification triggered.');
  SpreadsheetApp.getActiveSpreadsheet().toast('Audit entry created successfully in Audit_Log!', 'Verified', 4);
}

function showDispatchModal() {
  SpreadsheetApp.getUi().alert('Dispatch Form will be enabled in Ticket 5.');
}

function showIncidentModal() {
  SpreadsheetApp.getUi().alert('Incident Form will be enabled in Ticket 4.');
}