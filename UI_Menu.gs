/**
 * @fileoverview Custom Menus, Sidebar Launchers, and UI Handlers.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚚 Fleet Dispatch')
    .addItem('🖥️ Open Dispatch Control Sidebar', 'openFleetSidebar')
    .addSeparator()
    .addItem('➕ Add New Vehicle to Fleet', 'showAddVehicleModal')
    .addItem('⚠️ Report Incident / OOS', 'showIncidentModal')
    .addItem('✅ Check Vehicle Back Into Service', 'showCheckInModal')
    .addItem('🛑 Decommission Vehicle', 'menuPromptDecommissionVehicle')
    .addSeparator()
    .addItem('📋 Daily Dispatch Checkout (Coming in T5)', 'showDispatchModal')
    .addItem('🧪 Test Audit Log', 'testAuditLogger')
    .addToUi();
}

function menuPromptAddVehicle() {
  showAddVehicleModal();
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

function openFleetSidebar() {
  const html = HtmlService.createHtmlOutput(getSidebarHtml())
    .setTitle('🚚 Fleet Operations')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}


function testAuditLogger() {
  logAudit('TEST_RUN', 'System', 'SELF_CHECK', 'Audit log verification triggered.');
  SpreadsheetApp.getActiveSpreadsheet().toast('Audit entry created successfully in Audit_Log!', 'Verified', 4);
}

function getSidebarHtml() {
  return `
  <!DOCTYPE html>
  <html>
  <head>
    <style>
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 14px; color: #202124; }
      h2 { font-size: 16px; margin: 0 0 10px 0; color: #1a73e8; }
      p { font-size: 12px; color: #5f6368; line-height: 1.4; margin-bottom: 12px; }
      .action-card { background: #f8f9fa; border: 1px solid #dadce0; border-radius: 8px; padding: 12px; margin-bottom: 12px; }
      .action-card h4 { margin: 0 0 8px 0; font-size: 13px; color: #3c4043; text-transform: uppercase; letter-spacing: 0.5px; }
      .btn { display: block; width: 100%; box-sizing: border-box; text-align: center; padding: 10px; margin: 6px 0; border-radius: 6px; font-size: 13px; font-weight: 500; cursor: pointer; border: none; transition: opacity 0.2s; }
      .btn:hover { opacity: 0.9; }
      .btn-blue { background: #1a73e8; color: #fff; }
      .btn-green { background: #188038; color: #fff; }
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
      <button class="btn btn-blue" onclick="runScript('showAddVehicleModal')">➕ Add New Vehicle</button>
      <button class="btn btn-red" onclick="runScript('menuPromptDecommissionVehicle')">🛑 Decommission Vehicle</button>
    </div>

    <div class="action-card">
      <h4>Service & Maintenance</h4>
      <button class="btn btn-amber" onclick="runScript('showIncidentModal')">⚠️ Report Issue / OOS</button>
      <button class="btn btn-green" onclick="runScript('showCheckInModal')">✅ Check In / Return to Service</button>
    </div>

    <div class="action-card">
      <h4>Daily Dispatch</h4>
      <button class="btn btn-outline" onclick="runScript('showDispatchModal')">📋 Daily Dispatch Form</button>
    </div>

    <div class="action-card">
      <h4>Utilities</h4>
      <button class="btn btn-outline" onclick="runScript('testAuditLogger')">🧪 Test Audit Logger</button>
    </div>

    <div class="footer-note">All actions logged with active user and timestamp.</div>

    <script>
      function runScript(funcName) {
        google.script.run
          .withFailureHandler(function(err) {
            alert('Error running ' + funcName + ': ' + err.message);
          })[funcName]();
      }
    </script>
  </body>
  </html>
  `;
}