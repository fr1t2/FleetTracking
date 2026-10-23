/**
 * @fileoverview VEO Fleet Tracking & Dispatch Core Engine
 * Version 1.2 - Expanded Fleet Attributes & Dashboard Control Center
 */

const CONFIG_SHEET = '_Config';
const AUDIT_SHEET = 'Audit_Log';
const VEHICLES_SHEET = 'Vehicles';
const DEPLOYMENTS_SHEET = 'Deployments';
const INCIDENTS_SHEET = 'Incidents_Service';

/**
 * Creates custom top menu on open and provides sidebar access.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚚 Fleet Dispatch')
    .addItem('🖥️ Open Dispatch Control Sidebar', 'openFleetSidebar')
    .addItem('➕ Add New Vehicle to Fleet', 'showAddVehicleModal')
    .addItem('🛑 Decommission Vehicle', 'menuPromptDecommissionVehicle')
    .addSeparator()
    .addItem('📋 Open Daily Dispatch Form', 'showDispatchModal')
    .addItem('⚠️ Report Incident / OOS', 'showIncidentModal')
    .addSeparator()
    .addItem('⚙️ Upgrade Vehicles Sheet Columns', 'updateVehiclesTableSchema')
    .addItem('🧪 Test Audit Log', 'testAuditLogger')
    .addToUi();
}

/**
 * Opens the Fleet Dispatch Control Panel directly on the right side of the sheet.
 */
function openFleetSidebar() {
  const html = HtmlService.createHtmlOutput(getSidebarHtml())
    .setTitle('🚚 Fleet Dispatch Panel')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Reads settings and dropdown options dynamically from _Config.
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

  config.shifts = getNonEmptyColumn(4);
  config.zones = getNonEmptyColumn(5);
  config.oosReasons = getNonEmptyColumn(6);
  config.statuses = getNonEmptyColumn(7);
  config.focusOptions = getNonEmptyColumn(8);

  return config;
}

/**
 * Standardized Audit Logger.
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
 * Upgrades the Vehicles sheet to include all new granular columns.
 * Can be run safely without erasing existing vehicle rows.
 */
function updateVehiclesTableSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet) throw new Error('Vehicles sheet not found');

  const newHeaders = [
    'Vehicle Internal Name',
    'Rental Agreement / Unit #',
    'Rental Agency',
    'Make & Model',
    'VIN',
    'License Plate',
    'Current Status',
    'Pick-up Date',
    'Service Number',
    'Current Assigned Zone',
    'Current Assigned Tech',
    'Last Service Date',
    'Notes',
    'Last Updated By',
    'Last Updated At'
  ];

  sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight('bold');
  SpreadsheetApp.getActiveSpreadsheet().toast('Vehicles columns updated with full attributes!', 'Schema Updated', 4);
}

/**
 * Displays the modal dialog form for adding a new vehicle.
 */
function showAddVehicleModal() {
  const html = HtmlService.createHtmlOutput(getAddVehicleModalHtml())
    .setWidth(520)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(html, '➕ Add New Vehicle to Fleet');
}

/**
 * Backend processor for adding a vehicle from the modal.
 */
function processAddVehicle(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(VEHICLES_SHEET);
  if (!sheet) throw new Error('Vehicles sheet not found');

  const now = new Date();
  const user = Session.getActiveUser().getEmail() || 'local-user';

  // Order matching newHeaders
  const rowData = [
    formData.vanName || '',
    formData.rentalId || '',
    formData.rentalAgency || 'Enterprise',
    formData.makeModel || 'Cargo Van',
    formData.vin || '',
    formData.licensePlate || '',
    'Active',
    formData.pickupDate ? new Date(formData.pickupDate) : now,
    formData.serviceNum || '',
    '', // Assigned Zone
    '', // Assigned Tech
    now, // Last service date
    formData.notes || '',
    user,
    now
  ];

  sheet.appendRow(rowData);

  logAudit('VEHICLE_ADDED', 'Vehicles', formData.rentalId || formData.vanName, {
    vanName: formData.vanName,
    rentalId: formData.rentalId,
    agency: formData.rentalAgency,
    vin: formData.vin,
    plate: formData.licensePlate,
    addedBy: user
  });

  return { success: true, message: `Vehicle ${formData.vanName} added successfully!` };
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
    const internalName = String(data[i][0]).trim().toLowerCase();
    const unitId = String(data[i][1]).trim().toLowerCase();
    const searchTarget = String(rentalId).trim().toLowerCase();

    if (internalName === searchTarget || unitId === searchTarget) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('Vehicle "' + rentalId + '" not found in registry.');
  }

  const user = Session.getActiveUser().getEmail() || 'system';
  const now = new Date();

  // Col 7 is Current Status, Col 13 is Notes, Col 14 is Last Updated By, Col 15 is Last Updated At
  sheet.getRange(rowIndex, 7).setValue('Decommissioned');
  const existingNotes = sheet.getRange(rowIndex, 13).getValue();
  const updatedNotes = (existingNotes ? existingNotes + ' | ' : '') + 'Decommissioned: ' + (reason || 'No reason provided');
  sheet.getRange(rowIndex, 13).setValue(updatedNotes);
  sheet.getRange(rowIndex, 14).setValue(user);
  sheet.getRange(rowIndex, 15).setValue(now);

  logAudit('VEHICLE_DECOMMISSIONED', 'Vehicles', rentalId, {
    reason: reason,
    decommissionedBy: user
  });

  return { success: true, message: `Vehicle ${rentalId} has been marked Decommissioned.` };
}

function menuPromptDecommissionVehicle() {
  const ui = SpreadsheetApp.getUi();
  const idResp = ui.prompt('Decommission Vehicle', 'Enter Van clean name (e.g. Van #01) or Rental ID to decommission:', ui.ButtonSet.OK_CANCEL);
  if (idResp.getSelectedButton() !== ui.Button.OK) return;

  const reasonResp = ui.prompt('Decommission Vehicle', 'Enter reason (e.g. Returned to Enterprise / End of lease):', ui.ButtonSet.OK_CANCEL);
  if (reasonResp.getSelectedButton() !== ui.Button.OK) return;

  const res = decommissionVehicle(idResp.getResponseText(), reasonResp.getResponseText());
  ui.alert(res.message);
}

function testAuditLogger() {
  logAudit('TEST_RUN', 'System', 'SELF_CHECK', 'Audit log verification completed.');
  SpreadsheetApp.getActiveSpreadsheet().toast('Audit entry created successfully in Audit_Log!', 'Verified', 4);
}

function showDispatchModal() {
  SpreadsheetApp.getUi().alert('Daily Dispatch Form will be enabled in Ticket 5.');
}

function showIncidentModal() {
  SpreadsheetApp.getUi().alert('Incident & OOS Modal will be enabled in Ticket 4.');
}

/**
 * HTML definition for Add Vehicle Dialog
 */
function getAddVehicleModalHtml() {
  return `
  <!DOCTYPE html>
  <html>
  <head>
    <style>
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 15px; color: #202124; background: #fafafa; }
      h3 { margin-top: 0; color: #1a73e8; font-size: 18px; border-bottom: 2px solid #e8eaed; padding-bottom: 8px; }
      .form-group { margin-bottom: 12px; }
      label { display: block; font-size: 12px; font-weight: 600; color: #5f6368; margin-bottom: 4px; }
      input, select, textarea { width: 100%; box-sizing: border-box; padding: 8px 10px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px; background: #fff; }
      input:focus, select:focus, textarea:focus { outline: none; border-color: #1a73e8; box-shadow: 0 0 0 2px rgba(26,115,232,0.2); }
      .row { display: flex; gap: 10px; }
      .col { flex: 1; }
      .btn-bar { display: flex; justify-content: flex-end; gap: 10px; margin-top: 18px; }
      .btn { padding: 9px 18px; border-radius: 4px; font-size: 13px; font-weight: 500; cursor: pointer; border: none; }
      .btn-primary { background: #1a73e8; color: #fff; }
      .btn-secondary { background: #f1f3f4; color: #3c4043; }
      .status-msg { margin-top: 8px; font-size: 12px; font-weight: 500; }
    </style>
  </head>
  <body>
    <h3>➕ New Fleet Vehicle Entry</h3>
    <form id="vehicleForm">
      <div class="row">
        <div class="col form-group">
          <label>Clean Name *</label>
          <input type="text" id="vanName" placeholder="e.g. Van #19" required />
        </div>
        <div class="col form-group">
          <label>Rental / Unit # *</label>
          <input type="text" id="rentalId" placeholder="e.g. P9103003" required />
        </div>
      </div>
      <div class="row">
        <div class="col form-group">
          <label>Rental Agency</label>
          <select id="rentalAgency">
            <option value="Enterprise">Enterprise</option>
            <option value="Hertz">Hertz</option>
            <option value="Merchants Fleet">Merchants Fleet</option>
            <option value="Budget">Budget</option>
            <option value="Company Owned">Company Owned</option>
          </select>
        </div>
        <div class="col form-group">
          <label>Make & Model</label>
          <input type="text" id="makeModel" placeholder="e.g. Ford Transit 250" />
        </div>
      </div>
      <div class="row">
        <div class="col form-group">
          <label>VIN (17 digits)</label>
          <input type="text" id="vin" placeholder="1FTBR1Y89PKA..." maxlength="17" />
        </div>
        <div class="col form-group">
          <label>License Plate</label>
          <input type="text" id="licensePlate" placeholder="e.g. CO - ABC-1234" />
        </div>
      </div>
      <div class="row">
        <div class="col form-group">
          <label>Pick-up Date</label>
          <input type="date" id="pickupDate" />
        </div>
        <div class="col form-group">
          <label>Service / Contract #</label>
          <input type="text" id="serviceNum" placeholder="e.g. RA-88291" />
        </div>
      </div>
      <div class="form-group">
        <label>Initial Notes / Condition</label>
        <textarea id="notes" rows="2" placeholder="Tires good, Samsara unit installed..."></textarea>
      </div>
      <div class="btn-bar">
        <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
        <button type="submit" id="submitBtn" class="btn btn-primary">Save to Fleet</button>
      </div>
      <div id="status" class="status-msg"></div>
    </form>

    <script>
      document.getElementById('pickupDate').value = new Date().toISOString().split('T')[0];
      document.getElementById('vehicleForm').addEventListener('submit', function(e) {
        e.preventDefault();
        const submitBtn = document.getElementById('submitBtn');
        const status = document.getElementById('status');
        submitBtn.disabled = true;
        status.style.color = '#5f6368';
        status.innerText = 'Registering vehicle and logging audit entry...';

        const payload = {
          vanName: document.getElementById('vanName').value,
          rentalId: document.getElementById('rentalId').value,
          rentalAgency: document.getElementById('rentalAgency').value,
          makeModel: document.getElementById('makeModel').value,
          vin: document.getElementById('vin').value,
          licensePlate: document.getElementById('licensePlate').value,
          pickupDate: document.getElementById('pickupDate').value,
          serviceNum: document.getElementById('serviceNum').value,
          notes: document.getElementById('notes').value
        };

        google.script.run
          .withSuccessHandler(function(res) {
            status.style.color = '#137333';
            status.innerText = res.message;
            setTimeout(() => { google.script.host.close(); }, 1200);
          })
          .withFailureHandler(function(err) {
            submitBtn.disabled = false;
            status.style.color = '#d93025';
            status.innerText = 'Error: ' + err.message;
          })
          .processAddVehicle(payload);
      });
    </script>
  </body>
  </html>
  `;
}

/**
 * HTML definition for Dashboard Control Sidebar
 */
function getSidebarHtml() {
  return `
  <!DOCTYPE html>
  <html>
  <head>
    <style>
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 14px; color: #202124; background: #fff; }
      h2 { font-size: 16px; margin: 0 0 10px 0; color: #1a73e8; }
      p { font-size: 12px; color: #5f6368; line-height: 1.4; }
      .action-card { background: #f8f9fa; border: 1px solid #dadce0; border-radius: 8px; padding: 12px; margin-bottom: 12px; }
      .action-card h4 { margin: 0 0 6px 0; font-size: 13px; color: #3c4043; }
      .btn { display: block; width: 100%; box-sizing: border-box; text-align: center; padding: 10px; margin: 6px 0; border-radius: 6px; font-size: 13px; font-weight: 500; cursor: pointer; border: none; text-decoration: none; }
      .btn-blue { background: #1a73e8; color: #fff; }
      .btn-green { background: #1e8e3e; color: #fff; }
      .btn-amber { background: #f29900; color: #fff; }
      .btn-red { background: #d93025; color: #fff; }
      .btn-outline { background: #fff; border: 1px solid #dadce0; color: #3c4043; }
      .footer-note { font-size: 11px; color: #80868b; text-align: center; margin-top: 20px; }
    </style>
  </head>
  <body>
    <h2>🚚 Fleet Operations</h2>
    <p>Market Dispatch and Fleet Management Control Center</p>

    <div class="action-card">
      <h4>Fleet Management</h4>
      <button class="btn btn-blue" onclick="google.script.run.showAddVehicleModal()">➕ Add New Vehicle</button>
      <button class="btn btn-red" onclick="google.script.run.menuPromptDecommissionVehicle()">🛑 Decommission Vehicle</button>
    </div>

    <div class="action-card">
      <h4>Daily Dispatch</h4>
      <button class="btn btn-green" onclick="google.script.run.showDispatchModal()">📋 Daily Dispatch Form</button>
      <button class="btn btn-amber" onclick="google.script.run.showIncidentModal()">⚠️ Report Issue / OOS</button>
    </div>

    <div class="action-card">
      <h4>System Utilities</h4>
      <button class="btn btn-outline" onclick="google.script.run.updateVehiclesTableSchema()">⚙️ Update Sheet Columns</button>
      <button class="btn btn-outline" onclick="google.script.run.testAuditLogger()">🧪 Test Audit Logger</button>
    </div>

    <div class="footer-note">All actions automatically logged with timestamp and user ID.</div>
  </body>
  </html>
  `;
}